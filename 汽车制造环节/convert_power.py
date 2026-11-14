# -*- coding: utf-8 -*-
"""发电量（EJ）→ 估算装机容量（MW）转换，并按 tech / Subtech 两级分类。

读取 数据.xlsx 的：
  - power generation（发电量，宽表，整数年份列，Scenario/region/subcetor/technology/units）
  - CF（容量因子主表，宽表，region/subcetor/technology × 年份）
  - CIPG（单位装机材料强度，technology × 年份，t/GW）
以及补充工作簿 GCAM_发电量_容量因子匹配_2010-2050.xlsx 的 匹配明细（六种 CCS 因子）。

统一年份字段并筛选 1990–2050（含首尾，电力表保留 2021）。
- 按 region+subcetor+technology+year 匹配主表 CF（保留发电表 Scenario）；
- 2010 年前的 1990、2005 年，主表 CF 无对应列，改从
  `电力部门_容量因子.xlsx` 的 `全部` 长表按同一来源文件补齐（见 load_cf_history）；
- 仅对主表缺失的六种 CCS 技术，按 Scenario+region+subcetor+technology+year 从 匹配明细 补齐；
- 水电直接使用主表 CF 中匹配到的 hydro 因子。

公式：装机(MW) = 年发电量(EJ) × 10^9 ÷ (3.6 × 8760 × CF)，CF 用 0<CF≤1 小数。
每个原始技术记录计算未取整 MW；再按 Scenario+region+Year+tech+Subtech 加总，
最后对 Subtech 合计四舍五入。tech 只作两级标签，不生成 tech 合计行。

分类口径（本次修正）：
- tech 为旧归并分类，但 wind、wind_offshore 统一归入 wind；
- Subtech 为修正后的细分：biomass/coal 的 IGCC 与 conv 分别保存，
  wind 与 wind_offshore 虽同属 tech=wind，也分别保存。
- 原始 technology 名称保持不变，继续用于匹配 CF / CIPG。
"""
import csv
import math
import os
import sys
from collections import defaultdict, Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g

POWER_SHEET = "power generation"
CF_SHEET = "CF"
CIPG_SHEET = "CIPG"
CCS_MATCH_SHEET = "匹配明细"
CF_HISTORY_SHEET = "全部"
CF_HISTORY_SOURCE = "CF表(历史长表)"
# 主表 CF 从 2010 开始；更早的年份从长表补齐
CF_HISTORY_YEARS = tuple(y for y in g.POWER_PANEL_YEARS if y < 2010)

MID_DETAIL_CSV = os.path.join(g.OUTPUT_DIR, "_power_detail.csv")
MID_SUBTECH_CSV = os.path.join(g.OUTPUT_DIR, "_power_subtechs.csv")

# 六种需要从补充工作簿补齐因子的 CCS 技术
CCS_TECHS = [
    "biomass (IGCC CCS)", "biomass (conv CCS)",
    "coal (IGCC CCS)", "coal (conv pul CCS)",
    "gas (CC CCS)", "refined liquids (CC CCS)",
]

# 原始技术 -> (tech, Subtech)。字段大小写严格按任务口径。
TECH_MAP = {
    "biomass (IGCC CCS)": ("biomass_CCS", "biomass_IGCC_CCS"),
    "biomass (conv CCS)": ("biomass_CCS", "biomass_conv_CCS"),
    "coal (IGCC CCS)": ("coal_CCS", "coal_IGCC_CCS"),
    "coal (conv pul CCS)": ("coal_CCS", "coal_conv_pul_CCS"),
    "gas (CC CCS)": ("gas_CCS", "gas_CC_CCS"),
    "refined liquids (CC CCS)": ("refined liquids_CCS", "refined liquids_CC_CCS"),
    "hydro": ("hydro", "hydro"),
    "Gen_III": ("nuclear", "nuclear"),
    "Gen_II_LWR": ("nuclear", "nuclear"),
    "rooftop_pv": ("PV", "PV"),
    "PV": ("PV", "PV"),
    "PV_storage": ("PV", "PV"),
    "wind": ("wind", "wind"),
    "wind_storage": ("wind", "wind"),
    "wind_offshore": ("wind", "wind_offshore"),
}

# 八类 tech
TECH_ORDER = [
    "biomass_CCS", "coal_CCS", "gas_CCS", "hydro", "nuclear",
    "refined liquids_CCS", "PV", "wind",
]

# 十一类 Subtech（聚合与面板输出的固定顺序）
SUBTECH_ORDER = [
    "biomass_IGCC_CCS", "biomass_conv_CCS",
    "coal_IGCC_CCS", "coal_conv_pul_CCS",
    "gas_CC_CCS", "refined liquids_CC_CCS",
    "hydro", "nuclear", "PV", "wind", "wind_offshore",
]

SUBTECH_TO_TECH = {sub: tech for tech, sub in TECH_MAP.values()}

# 装机换算明细列（长表）
RESULT_FIELDS = [
    "Scenario", "region", "subcetor", "technology", "units",
    "Year", "Generation_EJ",
    "CapacityFactor", "CF_Source", "CF_MatchStatus", "AnnualHours",
    "InstalledCapacity_MW_unrounded",
    "CIPG_t_per_GW", "CIPG_Unit", "CIPG_Status",
    "tech", "Subtech", "Status",
]
# 十一类 Subtech 聚合列
SUBTECH_FIELDS = [
    "Scenario", "region", "Year", "tech", "Subtech",
    "InstalledCapacity_MW", "Status",
]
ANOMALY_FIELDS = [
    "Table", "Scenario", "region", "subcetor", "technology",
    "Year", "Type", "Detail",
]


def _meta_index(meta, *names):
    lower = {str(n).strip().lower(): idx for idx, n in meta}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def _header_index(header, *names):
    lower = {str(h).strip().lower(): i for i, h in enumerate(header)}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def load_power(wb):
    """读取 power generation 宽表，返回 [(meta, year, gen), ...]（仅范围内年份）。"""
    ws = wb[POWER_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    key_names = ["Scenario", "region", "subcetor", "technology", "units"]
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


def load_cf_sheet(wb):
    """读取 CF 主表，返回 dict[(region, subcetor, technology)] -> {year: cf}。"""
    ws = wb[CF_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    meta, year_cols = g.split_year_columns(header)
    key_names = ["region", "subcetor", "technology"]
    idx = {n: _meta_index(meta, n) for n in key_names}
    data = {}
    for r in rows:
        if r is None or r[idx["region"]] is None:
            continue
        key = (r[idx["region"]], r[idx["subcetor"]], r[idx["technology"]])
        data[key] = {y: g.to_number(r[col]) for col, y in year_cols}
    return data


def _hist_value(all_vals, key, year, files):
    """从长表中取某键某年的唯一 CF；同键多来源多值则不猜。"""
    data = all_vals.get((key, year), {})
    if files is None:
        values = {v for v in data.values() if v is not None}
    else:
        values = {data[f] for f in files if f in data and data[f] is not None}
    if len(values) == 1:
        return next(iter(values))
    return None


def load_cf_history(cf_sheet):
    """从 电力部门_容量因子.xlsx 的 全部 长表补齐 2010 年前的 CF。

    主表 CF 从 2010 开始，1990/2005 没有列。长表若只按
    region+subcetor+technology+year 直接取，会因同一键存在多个来源文件而冲突。
    因此先用主表 2010–2050 的数值反推每个键对应的来源文件（要求该文件在
    主表全部年份都与主表一致），再用同一文件取 1990/2005 值；地区级缺失时
    退回全局（region=None）同技术。仅补齐 g.POWER_PANEL_YEARS 中小于 2010
    的年份，不覆盖主表已有值，不插值、不填零。
    """
    if not CF_HISTORY_YEARS:
        return {}
    wide_years = sorted({y for vals in cf_sheet.values() for y in vals})
    wb = g.read_workbook_safe(g.CF_XLSX)
    ws = wb[CF_HISTORY_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {}
    for n in ["file", "region", "sector", "subsector", "technology",
              "year", "capacity_factor"]:
        idx[n] = _header_index(header, n)

    wanted = set(wide_years) | set(CF_HISTORY_YEARS)
    all_vals = defaultdict(dict)
    for r in rows:
        if r is None or r[idx["year"]] is None:
            continue
        y = r[idx["year"]]
        if y not in wanted:
            continue
        region = r[idx["region"]]
        if isinstance(region, str) and not region.strip():
            region = None
        key = (region, r[idx["subsector"]], r[idx["technology"]])
        cf = g.to_number(r[idx["capacity_factor"]])
        if cf is not None:
            all_vals[(key, y)][r[idx["file"]]] = cf
    wb.close()

    history = {}
    for key, vals in cf_sheet.items():
        # 候选来源文件：在任一主表年份出现过
        candidates = set()
        for y in wide_years:
            candidates |= set(all_vals.get((key, y), {}))
        # 保留在主表全部年份都与主表值一致的来源文件
        good = []
        for f in candidates:
            ok = True
            for y in wide_years:
                wv = vals.get(y)
                if wv is None:
                    continue
                fv = all_vals.get((key, y), {}).get(f)
                if fv is None or not math.isclose(
                        fv, wv, rel_tol=1e-12, abs_tol=1e-12):
                    ok = False
                    break
            if ok:
                good.append(f)

        for hy in CF_HISTORY_YEARS:
            v = _hist_value(all_vals, key, hy, good)
            if v is None:
                # 全局技术（region=None），如核电 Gen_II_LWR
                v = _hist_value(all_vals, (None, key[1], key[2]), hy, None)
            if v is not None:
                history[(key[0], key[1], key[2], hy)] = v
    return history


def load_ccs_supplement(path):
    """读取补充工作簿 匹配明细，仅取六种 CCS 技术的 CF。

    返回 (data, conflicts)：
    - data: dict[(scenario, region, subcetor, technology, year)] -> (cf, source_file, match_status)
    - conflicts: 同键多值冲突的键集合。
    """
    wb = g.read_workbook_safe(path)
    ws = wb[CCS_MATCH_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    ci = {}
    for n in ["scenario", "region", "subcetor", "technology", "year",
              "capacity_factor", "source_file", "match_status"]:
        ci[n] = _header_index(header, n)
    data = {}
    seen = {}
    conflicts = set()
    for r in rows:
        if r is None or r[ci["technology"]] is None:
            continue
        tech = r[ci["technology"]]
        if tech not in CCS_TECHS:
            continue
        key = (r[ci["scenario"]], r[ci["region"]], r[ci["subcetor"]],
               tech, r[ci["year"]])
        val = (g.to_number(r[ci["capacity_factor"]]),
               r[ci["source_file"]], r[ci["match_status"]])
        if key in seen:
            if seen[key] != val:
                conflicts.add(key)
        else:
            seen[key] = val
            data[key] = val
    wb.close()
    return data, conflicts


def load_cipg(wb):
    """读取 CIPG，返回 dict[technology] -> {year: float}，以及单位。"""
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


def match_cf(technology, scenario, region, subcetor, year,
             cf_sheet, cf_history, ccs_data, ccs_conflicts):
    """返回 (cf, cf_source, cf_match_status, conflict_flag)。"""
    if technology in CCS_TECHS:
        key = (scenario, region, subcetor, technology, year)
        if key in ccs_conflicts:
            return None, "匹配明细", "冲突", True
        if key in ccs_data:
            cf, src_file, match_status = ccs_data[key]
            return cf, f"匹配明细({src_file})", match_status or "补充匹配", False
        return None, "匹配明细", "缺失", False

    cf_year = cf_sheet.get((region, subcetor, technology), {})
    cf = cf_year.get(year)
    if cf is None and year in CF_HISTORY_YEARS:
        cf = cf_history.get((region, subcetor, technology, year))
        if cf is not None:
            return cf, CF_HISTORY_SOURCE, "历史匹配", False
        return None, CF_HISTORY_SOURCE, "缺失", False
    if cf is None:
        return None, "CF表", "缺失", False
    return cf, "CF表(主表)", "主表匹配", False


def compute():
    """执行电力转换，返回 (detail, subtechs, anomalies)。"""
    wb = g.read_workbook_safe(g.DATA_XLSX)
    power_records = load_power(wb)
    cf_sheet = load_cf_sheet(wb)
    cipg_data, cipg_unit = load_cipg(wb)
    wb.close()

    cf_history = load_cf_history(cf_sheet)
    ccs_data, ccs_conflicts = load_ccs_supplement(g.CCS_MATCH_XLSX)

    detail = []
    anomalies = []

    for meta, year, gen in power_records:
        scenario = meta["Scenario"]
        region = meta["region"]
        subcetor = meta["subcetor"]
        technology = meta["technology"]
        units = meta["units"]

        cf, cf_src, cf_status, cf_conflict = match_cf(
            technology, scenario, region, subcetor, year,
            cf_sheet, cf_history, ccs_data, ccs_conflicts)

        tech, subtech = TECH_MAP.get(technology, (None, None))
        if tech is None:
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "technology": technology, "Year": year,
                "Type": "unmapped_technology",
                "Detail": "技术不在 tech/Subtech 映射表中",
            })

        cipg_year = cipg_data.get(technology, {})
        cipg = cipg_year.get(year)
        cipg_status = "missing" if cipg is None else "ok"

        row = {
            "Scenario": scenario, "region": region, "subcetor": subcetor,
            "technology": technology, "units": units,
            "Year": year, "Generation_EJ": gen,
            "CapacityFactor": cf, "CF_Source": cf_src,
            "CF_MatchStatus": cf_status, "AnnualHours": g.HOURS_PER_YEAR,
            "InstalledCapacity_MW_unrounded": None,
            "CIPG_t_per_GW": cipg, "CIPG_Unit": cipg_unit,
            "CIPG_Status": cipg_status,
            "tech": tech, "Subtech": subtech, "Status": None,
        }

        if gen is None:
            row["Status"] = "blank_generation"
            detail.append(row)
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "technology": technology, "Year": year,
                "Type": "blank_generation", "Detail": "发电量空白，装机留空",
            })
            continue

        if g.is_zero(gen):
            row["InstalledCapacity_MW_unrounded"] = 0.0
            row["Status"] = "zero_generation"
            detail.append(row)
            continue

        # 正发电量：检查因子
        if cf_conflict:
            row["Status"] = "conflict_factor"
            detail.append(row)
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "technology": technology, "Year": year,
                "Type": "conflict_factor", "Detail": "补充因子同键多值冲突",
            })
            continue

        if cf is None:
            row["Status"] = "missing_cf"
            detail.append(row)
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "technology": technology, "Year": year,
                "Type": "missing_cf", "Detail": f"缺少有效容量因子（{cf_src}）",
            })
            continue

        if not (0 < cf <= 1):
            row["Status"] = "invalid_cf"
            detail.append(row)
            anomalies.append({
                "Table": POWER_SHEET, "Scenario": scenario, "region": region,
                "subcetor": subcetor, "technology": technology, "Year": year,
                "Type": "invalid_cf", "Detail": f"容量因子不在(0,1]：CF={cf}",
            })
            continue

        mw = gen * g.GIGA / (g.KWH_PER_J * g.HOURS_PER_YEAR * cf)
        row["InstalledCapacity_MW_unrounded"] = mw
        row["Status"] = "ok"
        detail.append(row)

    subtechs = aggregate(detail)
    return detail, subtechs, anomalies


def aggregate(detail):
    """按 Scenario+region+Year+tech+Subtech 汇总未取整装机，最后对 Subtech 取整。

    汇总键包含 tech 与 Subtech；因每个 Subtech 只对应一个 tech，
    该键与 Scenario+region+Year+Subtech 等价，且不产生 tech 合计行。
    """
    # 确定 (scenario, region, year) 全域
    universe = sorted({(r["Scenario"], r["region"], r["Year"]) for r in detail})
    groups = defaultdict(lambda: {"sum": 0.0, "has_none": False, "has_record": False})
    for r in detail:
        key = (r["Scenario"], r["region"], r["Year"], r["tech"], r["Subtech"])
        groups[key]["has_record"] = True
        mw = r["InstalledCapacity_MW_unrounded"]
        if mw is None:
            groups[key]["has_none"] = True
        else:
            groups[key]["sum"] += mw

    rows = []
    for scenario, region, year in universe:
        for subtech in SUBTECH_ORDER:
            tech = SUBTECH_TO_TECH[subtech]
            key = (scenario, region, year, tech, subtech)
            grp = groups.get(key)
            if grp is None or not grp["has_record"]:
                rows.append({
                    "Scenario": scenario, "region": region, "Year": year,
                    "tech": tech, "Subtech": subtech,
                    "InstalledCapacity_MW": None,
                    "Status": "源表无记录",
                })
            elif grp["has_none"]:
                rows.append({
                    "Scenario": scenario, "region": region, "Year": year,
                    "tech": tech, "Subtech": subtech,
                    "InstalledCapacity_MW": None,
                    "Status": "缺项（部分源记录无法计算）",
                })
            else:
                rows.append({
                    "Scenario": scenario, "region": region, "Year": year,
                    "tech": tech, "Subtech": subtech,
                    "InstalledCapacity_MW": g.round_half_up(grp["sum"]),
                    "Status": "ok",
                })
    return rows


def summary(detail, subtechs, anomalies):
    by_status = Counter(r["Status"] for r in detail)
    years = sorted({r["Year"] for r in detail})
    scenarios = sorted({r["Scenario"] for r in detail})
    n_supplement = sum(1 for r in detail if r["CF_Source"] and r["CF_Source"].startswith("匹配明细"))
    sub_status = Counter(c["Status"] for c in subtechs)
    return {
        "detail_rows": len(detail),
        "subtech_rows": len(subtechs),
        "years": years,
        "scenarios": scenarios,
        "detail_status": dict(by_status),
        "subtech_status": dict(sub_status),
        "anomalies": len(anomalies),
        "n_supplement": n_supplement,
    }


def main():
    g.ensure_dir(g.OUTPUT_DIR)
    detail, subtechs, anomalies = compute()
    with open(MID_DETAIL_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=RESULT_FIELDS)
        w.writeheader()
        for r in detail:
            w.writerow(r)
    with open(MID_SUBTECH_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=SUBTECH_FIELDS)
        w.writeheader()
        for r in subtechs:
            w.writerow(r)
    s = summary(detail, subtechs, anomalies)
    print("=== 电力转换核验摘要 ===")
    print("换算明细行数:", s["detail_rows"])
    print("十一类 Subtech 行数:", s["subtech_rows"])
    print("实际年份:", s["years"])
    print("情景数:", len(s["scenarios"]), s["scenarios"])
    print("明细状态:", s["detail_status"])
    print("Subtech 状态:", s["subtech_status"])
    print("异常数:", s["anomalies"])
    print("补充因子使用记录数:", s["n_supplement"])


if __name__ == "__main__":
    main()
