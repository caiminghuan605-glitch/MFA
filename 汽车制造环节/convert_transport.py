# -*- coding: utf-8 -*-
"""新能源汽车服务量 → 车辆保有量转换。

读取 数据.xlsx 的：
  - transport_service（服务量，宽表，整数年份列，含 Scenario/Region/Sector/Subsector/Technology/Units）
  - annual travel per vehicle（单车年行驶里程 VKT，宽表，VKT_YYYY 列）
  - All factor（平均载运系数，宽表，文本年份列）

统一年份字段并筛选 2010–2050（含首尾，保留区间内实际存在的 GCAM 模型年份，
交通表保留 2020），按 Region+Sector+Subsector+Technology+Year 精确匹配因子，
按公式计算车辆保有量（辆）。

公式：保有量(辆) = 年服务量原值 × 1,000,000 ÷ (单车年行驶里程 × 平均载运系数)
  - 服务量原值单位：million pass-km（客运）/ million ton-km（货运）
  - 单车年行驶里程单位：车公里/辆/年（km/vehicle/year）
  - 平均载运系数单位：人/车（客运）/ 吨/车（货运）

结果写入 Output/_transport_vehicle_stock.csv（中间产物），并返回结果与异常供合成。
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g

SERVICE_SHEET = "transport_service"
VKT_SHEET = "annual travel per vehicle"
LOADFACTOR_SHEET = "All factor"

MID_CSV = os.path.join(g.OUTPUT_DIR, "_transport_vehicle_stock.csv")

# 结果列（长表）：活动量表全部原始字段 + 归一 Year + 计算列
RESULT_FIELDS = [
    "Scenario", "Region", "Sector", "Subsector", "Technology", "Units",
    "Year", "ServiceValue",
    "AnnualVKT_km_per_veh", "VKT_Source",
    "LoadFactor", "LoadFactor_Unit", "LoadFactor_Source",
    "VehicleStock_vehicles", "Status",
]
ANOMALY_FIELDS = [
    "Table", "Scenario", "Region", "Sector", "Subsector", "Technology",
    "Year", "Type", "Detail",
]


def _meta_index(meta, *names):
    """按列名（忽略大小写）在 meta 列中查找索引；找不到返回 None。"""
    lower = {str(n).strip().lower(): idx for idx, n in meta}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def load_factor_wide(wb, sheet_name, key_cols, units_col=None):
    """读取宽表因子表，返回 (data, conflicts)。

    data: dict[(region, sector, subsector, technology)] -> {year: float|None}
    conflicts: set of key 出现重复且取值不一致的键。
    """
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    idx = {n: _meta_index(meta, n) for n in key_cols}
    if units_col:
        idx["__units__"] = _meta_index(meta, units_col)

    data = {}
    seen = {}
    conflicts = set()
    for r in rows:
        if r is None or r[idx[key_cols[0]]] is None:
            continue
        key = tuple(r[idx[c]] for c in key_cols)
        values = {y: g.to_number(r[col]) for col, y in year_cols}
        if key in seen:
            if seen[key] != values:
                conflicts.add(key)
        else:
            seen[key] = values
            data[key] = values
    return data, conflicts


def load_service(wb):
    """读取 transport_service 宽表，返回按 (Scenario,Region,Sector,Subsector,Technology,Units)
    分组的 {year: value}。"""
    ws = wb[SERVICE_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    key_names = ["Scenario", "Region", "Sector", "Subsector", "Technology", "Units"]
    idx = {n: _meta_index(meta, n) for n in key_names}
    records = []
    for r in rows:
        if r is None or r[idx["Scenario"]] is None:
            continue
        meta_vals = {n: r[idx[n]] for n in key_names}
        for col, y in year_cols:
            if not g.in_range(y):
                continue
            records.append((meta_vals, y, g.to_number(r[col])))
    return records


def loadfactor_unit(units):
    """由服务量单位判断载运系数单位：客运 人/车，货运 吨/车。"""
    if units is None:
        return ""
    u = str(units).lower()
    if "pass" in u:
        return "人/车"
    if "ton" in u:
        return "吨/车"
    return ""


def compute():
    """执行交通转换，返回 (results, anomalies)。"""
    wb = g.read_workbook_safe(g.DATA_XLSX)

    vkt_data, vkt_conflicts = load_factor_wide(
        wb, VKT_SHEET,
        key_cols=["region", "sector", "subsector", "technology"],
        units_col="units",
    )
    lf_data, lf_conflicts = load_factor_wide(
        wb, LOADFACTOR_SHEET,
        key_cols=["region", "sector", "subsector", "technology"],
    )
    service_records = load_service(wb)
    wb.close()

    results = []
    anomalies = []

    for meta, year, svc in service_records:
        scenario = meta["Scenario"]
        region = meta["Region"]
        sector = meta["Sector"]
        subsector = meta["Subsector"]
        technology = meta["Technology"]
        units = meta["Units"]
        fkey = (region, sector, subsector, technology)

        vkt_year = vkt_data.get(fkey, {})
        lf_year = lf_data.get(fkey, {})
        vkt = vkt_year.get(year)
        lf = lf_year.get(year)

        # 因子来源（表级，与年份无关）
        vkt_src = VKT_SHEET
        lf_src = LOADFACTOR_SHEET
        lf_unit = loadfactor_unit(units)

        row = {
            "Scenario": scenario, "Region": region, "Sector": sector,
            "Subsector": subsector, "Technology": technology, "Units": units,
            "Year": year, "ServiceValue": svc,
            "AnnualVKT_km_per_veh": vkt, "VKT_Source": vkt_src,
            "LoadFactor": lf, "LoadFactor_Unit": lf_unit, "LoadFactor_Source": lf_src,
            "VehicleStock_vehicles": None, "Status": None,
        }

        # ---- 状态判定 ----
        if svc is None:
            row["Status"] = "blank_service"
            results.append(row)
            anomalies.append(_anomaly(meta, year, "blank_service",
                                      "服务量空白，保有量留空"))
            continue

        if g.is_zero(svc):
            row["VehicleStock_vehicles"] = 0
            row["Status"] = "zero_service"
            results.append(row)
            continue

        # 正服务量：检查因子
        if fkey in vkt_conflicts or fkey in lf_conflicts:
            row["Status"] = "conflict_factor"
            results.append(row)
            anomalies.append(_anomaly(meta, year, "conflict_factor",
                                      "因子表出现重复键且取值不一致"))
            continue

        if vkt is None or lf is None:
            which = "VKT" if vkt is None else "All factor"
            if vkt is None and lf is None:
                which = "VKT 与 All factor"
            row["Status"] = "missing_factor"
            results.append(row)
            anomalies.append(_anomaly(meta, year, "missing_factor",
                                      f"缺失因子：{which}"))
            continue

        if vkt <= 0 or lf <= 0:
            row["Status"] = "nonpositive_factor"
            results.append(row)
            anomalies.append(_anomaly(meta, year, "nonpositive_factor",
                                      f"非正因子：VKT={vkt}, LoadFactor={lf}"))
            continue

        vehicles = svc * 1_000_000.0 / (vkt * lf)
        row["VehicleStock_vehicles"] = g.round_half_up(vehicles)
        row["Status"] = "ok"
        results.append(row)

    return results, anomalies


def _anomaly(meta, year, typ, detail):
    return {
        "Table": "transport_service", "Scenario": meta["Scenario"],
        "Region": meta["Region"], "Sector": meta["Sector"],
        "Subsector": meta["Subsector"], "Technology": meta["Technology"],
        "Year": year, "Type": typ, "Detail": detail,
    }


def summary(results, anomalies):
    from collections import Counter
    by_status = Counter(r["Status"] for r in results)
    years = sorted({r["Year"] for r in results})
    scenarios = sorted({r["Scenario"] for r in results})
    return {
        "rows": len(results),
        "years": years,
        "scenarios": scenarios,
        "status": dict(by_status),
        "anomalies": len(anomalies),
        "n_ok": by_status.get("ok", 0),
        "n_zero": by_status.get("zero_service", 0),
    }


def main():
    g.ensure_dir(g.OUTPUT_DIR)
    results, anomalies = compute()
    with open(MID_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=RESULT_FIELDS)
        w.writeheader()
        for r in results:
            w.writerow(r)
    s = summary(results, anomalies)
    print("=== 交通转换核验摘要 ===")
    print("结果行数:", s["rows"])
    print("实际年份:", s["years"])
    print("情景数:", len(s["scenarios"]), s["scenarios"])
    print("状态分布:", s["status"])
    print("异常数:", s["anomalies"])
    print("中间产物:", MID_CSV)


if __name__ == "__main__":
    main()
