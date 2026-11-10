# -*- coding: utf-8 -*-
"""发电量（EJ）→ 估算装机容量（MW）转换。

读取 数据.xlsx 的：
  - power generation（发电量，宽表，整数年份列，含 Scenario/region/subcetor/output/technology/units）
  - CIPG（单位装机材料强度，宽表，technology × 年份，单位 t/GW）
以及 电力部门_容量因子.xlsx 的 electricity 页（region/sector/subsector/technology/year/capacity_factor）。

统一年份字段并筛选 2010–2050（含首尾，电力表保留 2021），
按 region+output+subcetor+technology+Year ↔ region+sector+subsector+technology+year 匹配容量因子 CF。
只保留能匹配到有效 CF（0<CF≤1）的记录；匹配不上的记录从结果表排除并登记异常。

公式：估算装机容量(MW) = 年发电量(EJ) × 10^9 ÷ (3.6 × 8760 × CF)
CIPG 按 technology+Year 精确关联，缺失年份留空并标记（不阻断容量计算）。
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g

POWER_SHEET = "power generation"
CIPG_SHEET = "CIPG"
CF_SHEET = "electricity"

MID_CSV = os.path.join(g.OUTPUT_DIR, "_power_installed_capacity.csv")

RESULT_FIELDS = [
    "Scenario", "region", "subcetor", "output", "technology", "units",
    "Year", "Generation_EJ",
    "CapacityFactor", "CF_Source", "AnnualHours",
    "InstalledCapacity_MW",
    "CIPG_t_per_GW", "CIPG_Unit", "CIPG_Status",
    "Status",
]
ANOMALY_FIELDS = [
    "Table", "Scenario", "region", "subcetor", "output", "technology",
    "Year", "Type", "Detail",
]


def _meta_index(meta, *names):
    lower = {str(n).strip().lower(): idx for idx, n in meta}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def load_power(wb):
    """读取 power generation 宽表，返回 [(meta, year, value), ...]（仅范围内年份）。"""
    ws = wb[POWER_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    key_names = ["Scenario", "region", "subcetor", "output", "technology", "units"]
    idx = {n: _meta_index(meta, n) for n in key_names}
    records = []
    for r in rows:
        if r is None or r[idx["Scenario"]] is None:
            continue
        meta_vals = {n: r[idx[n]] for n in key_names}
        for col, y in year_cols:
            if g.in_range(y):
                records.append((meta_vals, y, g.to_number(r[col])))
    return records


def load_cipg(wb):
    """读取 CIPG，返回 dict[technology] -> {year: float}。单位列单独返回。"""
    ws = wb[CIPG_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    tech_idx = _meta_index(meta, "technology")
    unit_idx = _meta_index(meta, "unit")
    data = {}
    unit_val = None
    for r in rows:
        if r is None or tech_idx is None or r[tech_idx] is None:
            continue
        tech = r[tech_idx]
        data[tech] = {y: g.to_number(r[col]) for col, y in year_cols}
        if unit_idx is not None and unit_val is None and r[unit_idx] is not None:
            unit_val = str(r[unit_idx]).strip()
    return data, unit_val or "t/GW"


def build_cf_lookup(wb):
    """构建容量因子三级查找表，并检测冲突键。

    返回 (regional, global_tech, global_sub, conflicts)：
    - regional:  (region, sector, subsector, technology, year) -> CF   （地区专用）
    - global_tech: (sector, subsector, technology, year) -> CF        （全局默认，技术级）
    - global_sub:  (sector, subsector, year) -> CF                    （全局默认，子部门级）
    - conflicts: 出现多个不同 CF 的键集合。
    """
    ws = wb[CF_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    # 该表列固定：region, sector, subsector, technology, year, capacity_factor
    col = {"region": 0, "sector": 1, "subsector": 2, "technology": 3,
           "year": 4, "cf": 5}

    regional = {}
    global_tech = {}
    global_sub = {}
    raw = {}          # 用于冲突检测：key -> set(CF)
    key_kind = {}     # key -> 'regional' | 'global_tech' | 'global_sub'

    def _insert(dst, key, cf, kind):
        if key not in raw:
            raw[key] = set()
            key_kind[key] = kind
        raw[key].add(cf)
        dst[key] = cf

    for r in rows[1:]:
        if r is None or r[col["year"]] is None:
            continue
        region = r[col["region"]]
        sector = r[col["sector"]]
        subsector = r[col["subsector"]]
        technology = r[col["technology"]]
        year = r[col["year"]]
        cf = g.to_number(r[col["cf"]])
        if cf is None:
            continue
        if region is not None and str(region).strip() != "":
            key = (region, sector, subsector, technology, year)
            _insert(regional, key, cf, "regional")
        elif technology is not None and str(technology).strip() != "":
            key = (sector, subsector, technology, year)
            _insert(global_tech, key, cf, "global_tech")
        else:
            key = (sector, subsector, year)
            _insert(global_sub, key, cf, "global_sub")

    conflicts = {k for k, v in raw.items() if len(v) > 1}
    return regional, global_tech, global_sub, conflicts


def lookup_cf(regional, global_tech, global_sub, conflicts,
              region, output, subcetor, technology, year):
    """三级查找容量因子，返回 (cf, source) 或 (None, 未命中/冲突说明)。"""
    key1 = (region, output, subcetor, technology, year)
    if key1 in regional:
        if key1 in conflicts:
            return None, "conflict_regional"
        return regional[key1], "electricity.xml (regional)"

    key2 = (output, subcetor, technology, year)
    if key2 in global_tech:
        if key2 in conflicts:
            return None, "conflict_global_tech"
        return global_tech[key2], "electricity.xml (global)"

    key3 = (output, subcetor, year)
    if key3 in global_sub:
        if key3 in conflicts:
            return None, "conflict_global_sub"
        return global_sub[key3], "electricity.xml (global, subsector)"

    return None, "no_cf"


def compute():
    """执行电力转换，返回 (results, anomalies)。"""
    wb = g.read_workbook_safe(g.DATA_XLSX)
    power_records = load_power(wb)
    cipg_data, cipg_unit = load_cipg(wb)
    wb.close()

    cf_wb = g.read_workbook_safe(g.CF_XLSX)
    regional, global_tech, global_sub, cf_conflicts = build_cf_lookup(cf_wb)
    cf_wb.close()

    results = []
    anomalies = []

    for meta, year, gen in power_records:
        scenario = meta["Scenario"]
        region = meta["region"]
        subcetor = meta["subcetor"]
        output = meta["output"]
        technology = meta["technology"]
        units = meta["units"]

        cf, cf_src = lookup_cf(regional, global_tech, global_sub, cf_conflicts,
                               region, output, subcetor, technology, year)

        # 匹配不上（含冲突）→ 排除，登记异常
        if cf is None:
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "output": output, "technology": technology,
                "Year": year, "Type": cf_src,
                "Detail": "无适用容量因子，记录从结果表排除",
            })
            continue

        if not (0 < cf <= 1):
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "output": output, "technology": technology,
                "Year": year, "Type": "invalid_cf",
                "Detail": f"容量因子不在 (0,1]：CF={cf}，记录排除",
            })
            continue

        # CIPG 按 technology+Year 关联
        cipg_year = cipg_data.get(technology, {})
        cipg = cipg_year.get(year)
        cipg_status = "missing" if cipg is None else "ok"

        row = {
            "Scenario": scenario, "region": region, "subcetor": subcetor,
            "output": output, "technology": technology, "units": units,
            "Year": year, "Generation_EJ": gen,
            "CapacityFactor": cf, "CF_Source": cf_src,
            "AnnualHours": g.HOURS_PER_YEAR,
            "InstalledCapacity_MW": None,
            "CIPG_t_per_GW": cipg, "CIPG_Unit": cipg_unit,
            "CIPG_Status": cipg_status,
            "Status": None,
        }

        if gen is None:
            row["Status"] = "blank_generation"
            results.append(row)
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "output": output, "technology": technology,
                "Year": year, "Type": "blank_generation",
                "Detail": "发电量空白，装机留空",
            })
            continue

        if g.is_zero(gen):
            row["InstalledCapacity_MW"] = 0
            row["Status"] = "zero_generation"
            results.append(row)
            continue

        mw = gen * g.GIGA / (g.KWH_PER_J * g.HOURS_PER_YEAR * cf)
        row["InstalledCapacity_MW"] = g.round_half_up(mw)
        row["Status"] = "ok"
        results.append(row)

    return results, anomalies


def summary(results, anomalies):
    from collections import Counter
    by_status = Counter(r["Status"] for r in results)
    years = sorted({r["Year"] for r in results})
    scenarios = sorted({r["Scenario"] for r in results})
    drop_types = Counter(a["Type"] for a in anomalies)
    return {
        "rows": len(results),
        "years": years,
        "scenarios": scenarios,
        "status": dict(by_status),
        "anomalies": len(anomalies),
        "drop_types": dict(drop_types),
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
    print("=== 电力转换核验摘要 ===")
    print("结果行数:", s["rows"])
    print("实际年份:", s["years"])
    print("情景数:", len(s["scenarios"]), s["scenarios"])
    print("状态分布:", s["status"])
    print("异常/排除数:", s["anomalies"], "类型:", s["drop_types"])
    print("中间产物:", MID_CSV)


if __name__ == "__main__":
    main()
