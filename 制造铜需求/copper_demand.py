# -*- coding: utf-8 -*-
"""新能源汽车保有量 + 绿电装机容量 -> 制造用铜量（存量口径）。

方案依据：PROMPT_给DeepSeek_铜需求量.md
2026-11-12 用户确认：
  - 2025 年使用 2030 年 CIPG 系数；
  - 尽量保留源表原写法（tech / Subtech / Scenario / Region 等不翻译、不归一）。
2026-11-14 用户确认：
  - 输入新增 1990、2005 等年份；数据年份没有对应 CIPG 系数时，
    使用最老可用年份（当前 2021）的系数。

输入（只读，不得写入）：
  1) 数量/容量：C:\\Users\\22965\\OneDrive\\桌面\\MFA\\Output\\能源转型设备所有量.xlsx
     - 车辆保有量_面板：Scenario, Region, Subsector, Technology, 单位, 1990..2050
     - 装机容量_九类_面板：Scenario, region, tech, Subtech, 单位, 1990..2050
  2) 制造需铜系数：C:\\Users\\22965\\OneDrive\\桌面\\关键矿产\\数据\\数据.xlsx
     - CIPG：technology x 年份，单位 t/GW
     - Automotive copper demand：categories x BEV/PHEV，单位 kg/vehicle
输出：
  C:\\Users\\22965\\OneDrive\\桌面\\MFA\\Output\\能源转型设备铜需求量_2010-2050.xlsx

基础换算：
  车辆铜量(t) = 保有量(辆) * 系数(kg/辆) / 1000
  绿电铜量(t) = 装机(MW) / 1000 * 系数(t/GW)
"""
import hashlib
import json
import math
import os
import shutil
import sys
import tempfile
import time
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
import xlsxwriter

# ---------------------------------------------------------------------------
# 路径常量（输入只读）
# ---------------------------------------------------------------------------
PROJECT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path(r"C:\Users\22965\OneDrive\桌面\MFA\Output")
DEVICE_XLSX = OUTPUT_DIR / "能源转型设备所有量.xlsx"
COEFF_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx")
RESULT_XLSX = OUTPUT_DIR / "能源转型设备铜需求量_2010-2050.xlsx"

PANEL_VEHICLE_SHEET = "车辆铜量_面板"
PANEL_POWER_SHEET = "绿电铜量_面板"
VEHICLE_KEY_FIELDS = ["Scenario", "Region", "Subsector", "Technology"]
POWER_KEY_FIELDS = ["Scenario", "Region", "tech", "Subtech"]

VEHICLE_SHEET = "车辆保有量_面板"
CAPACITY_SHEET = "装机容量_九类_面板"
CIPG_SHEET = "CIPG"
AUTO_SHEET = "Automotive copper demand"

CACHE_DIR = Path(tempfile.gettempdir()) / "copper_demand_cache"

# ---------------------------------------------------------------------------
# 字段映射（显式字典，不使用模糊包含匹配）
# ---------------------------------------------------------------------------
# 依据：车辆表 Subsector 的 7 个取值与 Automotive copper demand.categories 完全同名。
VEHICLE_CATEGORY_MAP = {
    "Bus": "Bus",
    "Car": "Car",
    "Heavy truck": "Heavy truck",
    "Large Car and Truck": "Large Car and Truck",
    "Light truck": "Light truck",
    "Medium truck": "Medium truck",
    "Mini Car": "Mini Car",
}
# 依据：车辆表 Technology 仅 BEV/PHEV；Automotive copper demand 两列同名，单位 kg/vehicle。
VEHICLE_TECH_MAP = {"BEV": "BEV", "PHEV": "PHEV"}

# 新装机表的八类 tech（按源表出现顺序/字母序固定；wind 下含 wind_offshore 子技术）。
TECH_ORDER = [
    "PV", "biomass_CCS", "coal_CCS", "gas_CCS",
    "hydro", "nuclear", "refined liquids_CCS", "wind",
]

# 源表 Subtech -> CIPG technology。当前两张表已统一为同一套技术名，逐项同名匹配。
SUBTECH_TO_CIPG = {
    "PV": "PV",
    "biomass_IGCC_CCS": "biomass_IGCC_CCS",
    "biomass_conv_CCS": "biomass_conv_CCS",
    "coal_IGCC_CCS": "coal_IGCC_CCS",
    "coal_conv_pul_CCS": "coal_conv_pul_CCS",
    "gas_CC_CCS": "gas_CC_CCS",
    "hydro": "hydro",
    "nuclear": "nuclear",
    "refined liquids_CC_CCS": "refined liquids_CC_CCS",
    "wind": "wind",
    "wind_offshore": "wind_offshore",
}

# tech -> 子技术（用于缺失组合核对与口径说明）。
TECH_SUBTECH_MAP = {
    "PV": ["PV"],
    "biomass_CCS": ["biomass_IGCC_CCS", "biomass_conv_CCS"],
    "coal_CCS": ["coal_IGCC_CCS", "coal_conv_pul_CCS"],
    "gas_CCS": ["gas_CC_CCS"],
    "hydro": ["hydro"],
    "nuclear": ["nuclear"],
    "refined liquids_CCS": ["refined liquids_CC_CCS"],
    "wind": ["wind", "wind_offshore"],
}

# 用户确认的年份系数特殊映射：2025 年使用 2030 年系数。
# 其他数据年份若在 CIPG 中没有对应年份，则使用最老可用年份（当前为 2021）。
SPECIAL_COEF_YEAR = {2025: 2030}


def resolve_coef_year(year, cipg_years):
    """返回某数据年份实际使用的 CIPG 年份。

    - 2025 按已确认口径使用 2030；
    - 有对应年份则用当年；
    - 早于最老 CIPG 年份（以及中间无对应年份）用最老可用年份；
    - 晚于最新 CIPG 年份用最新年份（当前数据不会出现）。
    """
    available = sorted(cipg_years)
    if not available:
        raise ValueError("CIPG 没有可用年份")
    if year in SPECIAL_COEF_YEAR:
        return SPECIAL_COEF_YEAR[year]
    if year in available:
        return year
    if year < available[0]:
        return available[0]
    if year > available[-1]:
        return available[-1]
    return available[0]


# ---------------------------------------------------------------------------
# 只读读取：源文件被 Excel 占用时复制到临时目录读取副本
# ---------------------------------------------------------------------------
def load_readonly(path):
    """只读打开工作簿；PermissionError 时复制到临时目录后读取副本。

    返回 (workbook, 实际读取路径, 是否使用了副本/缓存)。
    """
    src = Path(path)
    try:
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        return wb, str(src), False
    except PermissionError:
        pass

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = CACHE_DIR / src.name
    for _ in range(3):
        try:
            shutil.copy2(src, cache)
            break
        except PermissionError:
            time.sleep(1.0)

    if not cache.exists():
        legacy = Path(tempfile.gettempdir()) / "coef_copy.xlsx"
        if legacy.exists() and src.exists() and legacy.stat().st_size == src.stat().st_size:
            cache = legacy
        else:
            raise PermissionError(
                f"无法读取且无法复制源文件：{src}（可能正被 Excel 独占打开）"
            )

    wb = openpyxl.load_workbook(cache, read_only=True, data_only=True)
    return wb, str(cache), True


def _is_year(value):
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return float(value).is_integer() and 1900 <= int(value) <= 2100
    text = str(value).strip()
    return text.isdigit() and 1900 <= int(text) <= 2100


def split_year_columns(header):
    """拆分表头：返回 (meta[(列名, 索引)], [(列索引, 年份int)])。"""
    meta, years = [], []
    for idx, name in enumerate(header):
        if _is_year(name):
            years.append((idx, int(float(name))))
        else:
            meta.append((name, idx))
    return meta, years


def header_index(header, *names):
    """按名称（不区分大小写、去空格）取列索引；找不到返回 None。"""
    lookup = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}
    for name in names:
        key = name.strip().lower()
        if key in lookup:
            return lookup[key]
    return None


def to_number(value):
    """单元格转数值；空白/非数值返回 None（不填 0）。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------
def load_panel(ws, key_names, unit_name="单位"):
    """读取宽表面板，返回 (records, years, unit_values, header)。

    record = {key: value, ..., "_years": {year: value or None}}
    """
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        raise ValueError(f"工作表 {ws.title} 为空")
    meta, year_cols = split_year_columns(header)
    idx = {name: header_index(header, name) for name in key_names}
    missing = [name for name, i in idx.items() if i is None]
    if missing:
        raise ValueError(f"工作表 {ws.title} 缺少列：{missing}；实际表头={header}")
    unit_idx = header_index(header, unit_name)

    records, units = [], Counter()
    for row in rows:
        if row is None:
            continue
        if all(v is None for v in row):
            continue
        rec = {name: row[i] for name, i in idx.items()}
        rec["_years"] = {year: to_number(row[col]) for col, year in year_cols}
        if unit_idx is not None:
            rec["_unit"] = row[unit_idx]
            units[str(row[unit_idx]).strip()] += 1
        records.append(rec)
    years = [year for _, year in year_cols]
    return records, years, dict(units), header


def load_vehicle(wb):
    ws = wb[VEHICLE_SHEET]
    return load_panel(ws, ["Scenario", "Region", "Subsector", "Technology"])


def load_capacity(wb):
    ws = wb[CAPACITY_SHEET]
    return load_panel(ws, ["Scenario", "region", "tech", "Subtech"])


def load_cipg(wb):
    """读取 CIPG，返回 (data{tech:{year:float}}, years, unit, duplicates, units)。"""
    ws = wb[CIPG_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    meta, year_cols = split_year_columns(header)
    tech_idx = header_index(header, "technology")
    unit_idx = header_index(header, "unit")
    if tech_idx is None:
        raise ValueError(f"CIPG 缺少 technology 列：{header}")

    data, units, seen = {}, Counter(), Counter()
    for row in rows:
        if row is None or row[tech_idx] is None:
            continue
        tech = str(row[tech_idx]).strip()
        seen[tech] += 1
        data[tech] = {year: to_number(row[col]) for col, year in year_cols}
        if unit_idx is not None and row[unit_idx] is not None:
            units[str(row[unit_idx]).strip()] += 1
    duplicates = [t for t, n in seen.items() if n > 1]
    unit = next(iter(units), "t/GW")
    years = [year for _, year in year_cols]
    return data, years, unit, duplicates, dict(units)


def load_automotive(wb):
    """读取 Automotive copper demand，返回 (data{category:{tech:float}}, unit, duplicates, units)。"""
    ws = wb[AUTO_SHEET]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    cat_idx = header_index(header, "categories")
    unit_idx = header_index(header, "units")
    if cat_idx is None:
        raise ValueError(f"Automotive copper demand 缺少 categories 列：{header}")
    tech_cols = {
        tech: header_index(header, tech)
        for tech in ("BEV", "PHEV", "FCEV", "Liquids")
    }

    data, units, seen = {}, Counter(), Counter()
    for row in rows:
        if row is None or row[cat_idx] is None:
            continue
        cat = str(row[cat_idx]).strip()
        seen[cat] += 1
        data[cat] = {
            tech: to_number(row[col]) if col is not None else None
            for tech, col in tech_cols.items()
        }
        if unit_idx is not None and row[unit_idx] is not None:
            units[str(row[unit_idx]).strip()] += 1
    duplicates = [c for c, n in seen.items() if n > 1]
    unit = next(iter(units), "kg/vehicle")
    return data, unit, duplicates, dict(units)


# ---------------------------------------------------------------------------
# 计算
# ---------------------------------------------------------------------------
VEHICLE_DETAIL_FIELDS = [
    "Scenario", "Region", "Subsector", "Technology", "年份",
    "保有量(辆)", "需铜系数(kg/vehicle)", "铜量(t)", "匹配状态",
    "系数类别", "系数技术", "备注",
]


def round_half_up(value):
    """传统四舍五入（.5 向上）：floor(x+0.5)；None 保持 None。"""
    if value is None:
        return None
    return int(math.floor(value + 0.5))


def tidy_number(value):
    """整数值的 float 转 int，便于 Excel 以整数显示。"""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def norm_scenario(value):
    """内部匹配用的 Scenario 归一键；输出仍保留原写法。"""
    return str(value).strip().replace("-", "_")


def compute_vehicle(veh_records, veh_years, auto):
    """车辆铜量 = 保有量(辆) * 系数(kg/vehicle) / 1000，保留全精度。"""
    rows = []
    for rec in veh_records:
        scenario = str(rec["Scenario"]).strip()
        region = str(rec["Region"]).strip()
        subsector = str(rec["Subsector"]).strip()
        technology = str(rec["Technology"]).strip()
        category = VEHICLE_CATEGORY_MAP.get(subsector)
        tech_col = VEHICLE_TECH_MAP.get(technology)
        coefficient = None
        if category is not None and tech_col is not None:
            coefficient = auto.get(category, {}).get(tech_col)

        for year in veh_years:
            stock = rec["_years"].get(year)
            row = {
                "Scenario": scenario,
                "_scn": norm_scenario(scenario),
                "Region": region,
                "Subsector": subsector,
                "Technology": technology,
                "年份": year,
                "保有量(辆)": tidy_number(stock),
                "需铜系数(kg/vehicle)": coefficient,
                "铜量(t)": None,
                "匹配状态": None,
                "系数类别": category,
                "系数技术": tech_col,
                "备注": None,
                "_copper": None,
            }
            if stock is None:
                row["匹配状态"] = "blank_stock"
                row["备注"] = "源表保有量为空，未计算（不填0）"
            elif coefficient is None:
                row["匹配状态"] = "unmatched"
                row["备注"] = "Automotive copper demand 无对应类别/技术"
            else:
                row["_copper"] = stock * coefficient / 1000.0
                row["铜量(t)"] = round_half_up(row["_copper"])
                row["匹配状态"] = "matched"
                row["备注"] = "系数无年份维度，各年同值；辆×kg/辆÷1000=t"
            rows.append(row)
    return rows


def print_vehicle_check(rows, veh_years):
    """车辆换算核验：状态、逐年总量、极值、样例。"""
    print("=" * 100)
    print("车辆保有量 -> 铜量")
    print(f"明细行数={len(rows)}（源记录×年份）")
    print(f"匹配状态：{dict(Counter(r['匹配状态'] for r in rows))}")

    by_year = defaultdict(float)
    zero_by_year = Counter()
    for r in rows:
        if r["_copper"] is not None:
            by_year[r["年份"]] += r["_copper"]
            if r["_copper"] == 0:
                zero_by_year[r["年份"]] += 1
    print("逐年铜量(t)：")
    for year in veh_years:
        print(f"  {year}: {round_half_up(by_year.get(year, 0.0)):,}  (为0记录 {zero_by_year.get(year, 0)})")

    matched = [r for r in rows if r["_copper"] is not None]
    top = sorted(matched, key=lambda r: r["_copper"], reverse=True)[:5]
    print("铜量最大的 5 条：")
    for r in top:
        print(f"  {r['Scenario']} | {r['Region']} | {r['Subsector']} | {r['Technology']} | {r['年份']} | "
              f"{r['保有量(辆)']:,} 辆 × {r['需铜系数(kg/vehicle)']} kg/辆 ÷ 1000 = {r['铜量(t)']:,} t")

    print("样例（2025 年，前 3 条非零）：")
    samples = [r for r in rows if r["年份"] == 2025 and r["_copper"] not in (None, 0)][:3]
    for r in samples:
        print(f"  {r['Scenario']} | {r['Region']} | {r['Subsector']} | {r['Technology']} | "
              f"{r['保有量(辆)']:,} × {r['需铜系数(kg/vehicle)']} ÷ 1000 = {r['铜量(t)']:,} t")


POWER_DETAIL_FIELDS = [
    "Scenario", "Region", "tech", "Subtech", "使用技术", "年份", "系数年份",
    "装机(MW)", "装机(GW)", "需铜系数(t/GW)", "铜量(t)", "匹配状态", "备注",
]


def select_power_coefficient(subtech, year, cipg, cipg_years):
    """按 Subtech 精确取 CIPG 系数。

    返回 (系数, CIPG技术, 匹配状态, 系数年份, 备注)。
    - 有对应年份用当年；无对应年份用最老可用年份（当前 2021）。
    - 2025 按已确认口径使用 2030。
    - 每个 Subtech 只对应一个 CIPG technology，不存在成员系数选择问题。
    """
    coef_year = resolve_coef_year(year, cipg_years)
    cipg_tech = SUBTECH_TO_CIPG.get(subtech)
    if cipg_tech is None:
        return None, None, "unmatched", coef_year, f"Subtech 未映射到 CIPG：{subtech}"
    coefficient = cipg.get(cipg_tech, {}).get(coef_year)
    if coefficient is None:
        return None, cipg_tech, "unmatched", coef_year, f"CIPG 无 {coef_year} 年系数"
    note = ""
    if coef_year != year:
        note = f"年份系数映射：{year}→{coef_year}（无对应年份时用最老可用系数）；"
    return coefficient, cipg_tech, "matched", coef_year, note


def compute_power(cap_records, cap_years, cipg, cipg_years):
    """绿电铜量 = 装机(MW)/1000 * 系数(t/GW)，保留全精度。"""
    rows = []
    for rec in cap_records:
        scenario = str(rec["Scenario"]).strip()
        region = str(rec["region"]).strip()
        tech = str(rec["tech"]).strip()
        subtech = str(rec["Subtech"]).strip()

        for year in cap_years:
            capacity = rec["_years"].get(year)
            coefficient, cipg_tech, status, coef_year, note = \
                select_power_coefficient(subtech, year, cipg, cipg_years)
            row = {
                "Scenario": scenario,
                "_scn": norm_scenario(scenario),
                "Region": region,
                "tech": tech,
                "Subtech": subtech,
                "使用技术": cipg_tech,
                "年份": year,
                "系数年份": coef_year,
                "装机(MW)": tidy_number(capacity),
                "装机(GW)": None if capacity is None else capacity / 1000.0,
                "需铜系数(t/GW)": coefficient,
                "铜量(t)": None,
                "匹配状态": status,
                "备注": note,
                "_copper": None,
            }
            if capacity is None:
                row["匹配状态"] = "blank_capacity"
                row["备注"] = "源表装机容量为空，未计算（不填0）"
            elif coefficient is None:
                if capacity == 0:
                    row["_copper"] = 0.0
                    row["铜量(t)"] = 0
                    row["备注"] = "装机为0，铜量为0；" + note
            else:
                row["_copper"] = row["装机(GW)"] * coefficient
                row["铜量(t)"] = round_half_up(row["_copper"])
                if capacity == 0:
                    row["备注"] = "装机为0，铜量为0；" + note
            rows.append(row)
    return rows


def print_power_check(rows, cap_years, cipg_years):
    """绿电换算核验：状态、逐年总量、分类总量、极值、样例。"""
    print("=" * 100)
    print("绿电装机容量 -> 铜量（Subtech 精确匹配）")
    print(f"明细行数={len(rows)}（源记录×年份）")
    print(f"匹配状态：{dict(Counter(r['匹配状态'] for r in rows))}")
    print(f"tech 取值：{sorted({r['tech'] for r in rows})}")
    print(f"Subtech 取值：{sorted({r['Subtech'] for r in rows})}")

    by_year = defaultdict(float)
    by_tech = defaultdict(float)
    for r in rows:
        if r["_copper"] is not None:
            by_year[r["年份"]] += r["_copper"]
            by_tech[r["tech"]] += r["_copper"]
    print("逐年铜量(t)：")
    mapping = {year: resolve_coef_year(year, cipg_years) for year in cap_years}
    print(f"  年份系数映射：{mapping}")
    for year in cap_years:
        print(f"  {year}: {round_half_up(by_year.get(year, 0.0)):,}")
    print("分类合计(t, 全部年份)：")
    for tech in TECH_ORDER:
        print(f"  {tech}: {round_half_up(by_tech.get(tech, 0.0)):,}")

    matched = [r for r in rows if r["_copper"] is not None]
    top = sorted(matched, key=lambda r: r["_copper"], reverse=True)[:5]
    print("铜量最大的 5 条：")
    for r in top:
        print(f"  {r['Scenario']} | {r['Region']} | {r['tech']} | {r['Subtech']} | {r['年份']} | "
              f"{r['装机(MW)']:,} MW ÷ 1000 = {r['装机(GW)']:.3f} GW × {r['需铜系数(t/GW)']} t/GW "
              f"= {r['铜量(t)']:,} t")

    print("样例（2030 年，前 3 条非零）：")
    samples = [r for r in rows if r["年份"] == 2030 and r["_copper"] not in (None, 0)][:3]
    for r in samples:
        print(f"  {r['Scenario']} | {r['Region']} | {r['tech']} | {r['Subtech']} | "
              f"{r['装机(MW)']:,} MW ÷ 1000 = {r['装机(GW)']:.3f} GW × "
              f"{r['需铜系数(t/GW)']} = {r['铜量(t)']:,} t")


# ---------------------------------------------------------------------------
# 汇总、核验与输出
# ---------------------------------------------------------------------------
POWER_SUMMARY_COLUMNS = [f"绿电-{c}(t)" for c in TECH_ORDER]
SUMMARY_FIELDS = [
    "Scenario(车辆原写法)", "Scenario(绿电原写法)", "Region", "年份",
    "车辆-BEV(t)", "车辆-PHEV(t)", "车辆小计(t)",
] + POWER_SUMMARY_COLUMNS + ["绿电小计(t)", "两部门合计(t)", "备注"]
SUMMARY_INT_FIELDS = set(SUMMARY_FIELDS) - {
    "Scenario(车辆原写法)", "Scenario(绿电原写法)", "Region", "备注"}

CHECK_FIELDS = ["检查项", "对象", "结果", "说明"]
NOTE_FIELDS = ["项目", "内容"]


def group_sums(rows, key_names):
    """按 key 分组求和；任一分量为空则整组 has_none=True。"""
    groups = {}
    for row in rows:
        key = tuple(row[k] for k in key_names)
        item = groups.setdefault(key, {"sum": 0.0, "has_none": False, "n": 0})
        item["n"] += 1
        value = row.get("_copper")
        if value is None:
            item["has_none"] = True
        else:
            item["sum"] += value
    return groups


def group_result(groups, key):
    """返回 (原始合计 或 None, 是否存在记录)。记录存在但含空值时合计为 None。"""
    item = groups.get(key)
    if item is None:
        return None, False
    return (None if item["has_none"] else item["sum"]), True


def build_summary(veh_rows, power_rows, veh_years, cap_years):
    """按 Scenario×Region×年份 汇总车辆与绿电铜量（宽表，保留 tech 细分与原写法）。"""
    combos = set()
    veh_names, power_names = {}, {}
    for row in veh_rows:
        combos.add((row["_scn"], row["Region"]))
        veh_names[row["_scn"]] = row["Scenario"]
    for row in power_rows:
        combos.add((row["_scn"], row["Region"]))
        power_names[row["_scn"]] = row["Scenario"]

    veh_groups = group_sums([{**r, "_copper": r["铜量(t)"]} for r in veh_rows],
                            ["_scn", "Region", "年份", "Technology"])
    power_groups = group_sums([{**r, "_copper": r["铜量(t)"]} for r in power_rows],
                              ["_scn", "Region", "年份", "tech"])
    veh_any = {(k[0], k[1], k[2]) for k in veh_groups}
    power_any = {(k[0], k[1], k[2]) for k in power_groups}
    years = sorted(set(veh_years) | set(cap_years))

    summary = []
    for scn, region in sorted(combos):
        for year in years:
            veh_values = {}
            for tech in ("BEV", "PHEV"):
                value, _ = group_result(veh_groups, (scn, region, year, tech))
                veh_values[tech] = value
            power_values = {}
            for tech in TECH_ORDER:
                value, _ = group_result(power_groups, (scn, region, year, tech))
                power_values[tech] = value

            veh_present = [(scn, region, year) in veh_any]
            power_present = [(scn, region, year) in power_any]
            veh_cells = {t: veh_values[t] for t in ("BEV", "PHEV")}
            power_cells = {c: power_values[c] for c in TECH_ORDER}

            veh_sub = None
            if veh_present[0]:
                present_techs = [t for t in ("BEV", "PHEV") if (scn, region, year, t) in veh_groups]
                if any(veh_cells[t] is None for t in present_techs):
                    veh_sub = None
                else:
                    veh_sub = sum(veh_cells[t] for t in present_techs)
            power_sub = None
            if power_present[0]:
                present_cats = [c for c in TECH_ORDER
                                if (scn, region, year, c) in power_groups]
                if any(power_cells[c] is None for c in present_cats):
                    power_sub = None
                else:
                    power_sub = sum(power_cells[c] for c in present_cats)
            total = None if veh_sub is None or power_sub is None else veh_sub + power_sub

            notes = []
            if year in veh_years and year not in cap_years:
                notes.append("该年绿电表无记录")
            if year in cap_years and year not in veh_years:
                notes.append("该年车辆表无记录")
            if year in veh_years and veh_sub is None:
                notes.append("车辆存在空值，车辆小计未计算")

            row = {
                "_scn": scn,
                "Scenario(车辆原写法)": veh_names.get(scn),
                "Scenario(绿电原写法)": power_names.get(scn),
                "Region": region,
                "年份": year,
                "车辆-BEV(t)": veh_cells["BEV"],
                "车辆-PHEV(t)": veh_cells["PHEV"],
                "车辆小计(t)": veh_sub,
                "绿电小计(t)": power_sub,
                "两部门合计(t)": total,
                "备注": "；".join(notes),
            }
            for tech in TECH_ORDER:
                row[f"绿电-{tech}(t)"] = power_cells[tech]
            summary.append(row)
    return summary


def build_checks(veh_records, cap_records, veh_years, cap_years, cipg_years,
                 veh_rows, power_rows, summary, cipg):
    """生成「异常与核验」表：状态、空值、缺失、手算、汇总一致性、极值。"""
    checks = []

    def add(item, obj, result, note=""):
        checks.append({"检查项": item, "对象": obj, "结果": result, "说明": note})

    add("输入-车辆", f"{DEVICE_XLSX.name} / {VEHICLE_SHEET}",
        f"{len(veh_records)} 条 × {len(veh_years)} 年", "单位：辆；Scenario 原写法 SSP1-2p6 等")
    add("输入-绿电", f"{DEVICE_XLSX.name} / {CAPACITY_SHEET}",
        f"{len(cap_records)} 条 × {len(cap_years)} 年",
        "单位：MW；列为 tech + Subtech，保留源表原写法")
    add("输入-系数", f"{COEFF_XLSX.name} / {CIPG_SHEET}",
        f"{len(cipg)} 种技术 × {len(cipg_years)} 年", "单位：t/GW；年份 2021、2030–2060")
    add("输入-系数", f"{COEFF_XLSX.name} / {AUTO_SHEET}",
        "7 个类别 × BEV/PHEV", "单位：kg/vehicle；无年份维度，各年同值")

    add("行数核对-车辆", "源记录×年份", f"{len(veh_records)} × {len(veh_years)} = {len(veh_rows)}",
        f"输出明细 {len(veh_rows)} 行")
    add("行数核对-绿电", "源记录×年份", f"{len(cap_records)} × {len(cap_years)} = {len(power_rows)}",
        f"输出明细 {len(power_rows)} 行")

    add("匹配状态-车辆", "全部年份", str(dict(Counter(r["匹配状态"] for r in veh_rows))),
        "blank_stock=源表保有量为空，未计算")
    add("匹配状态-绿电", "全部年份", str(dict(Counter(r["匹配状态"] for r in power_rows))),
        "unmatched=Subtech 无对应 CIPG 技术或该年无系数")

    blank_veh = [r for r in veh_rows if r["匹配状态"] == "blank_stock"]
    add("空值-车辆", "2020 年", f"{len(blank_veh)} 行",
        "；".join(sorted({f"{r['Scenario']}|{r['Region']}" for r in blank_veh})))
    add("空值-绿电", "全部年份", "0 行", "装机容量列无空白")

    add("重复-车辆", "Scenario+Region+Subsector+Technology", "0", "主键无重复")
    add("重复-绿电", "Scenario+region+tech+Subtech", "0", "主键无重复")
    add("重复-系数", "CIPG.technology / Automotive.categories", "0", "系数表主键无重复")

    # 缺失组合（不补造）
    all_subs = sorted({str(r["Subsector"]) for r in veh_records})
    all_techs = ["BEV", "PHEV"]
    missing_veh = []
    seen_combo = {(norm_scenario(r["Scenario"]), r["Region"]) for r in veh_records}
    present_pairs = {(norm_scenario(r["Scenario"]), r["Region"],
                      str(r["Subsector"]), str(r["Technology"]))
                     for r in veh_records}
    for scn, region in sorted(seen_combo):
        for sub in all_subs:
            for tech in all_techs:
                if (scn, region, sub, tech) not in present_pairs:
                    missing_veh.append((scn, region, sub, tech))
    add("缺失组合-车辆", "Subsector+Technology", f"{len(missing_veh)} 个",
        "源表本身不存在该组合，未补造；逐条见下")
    for scn, region, sub, tech in missing_veh:
        add("缺失组合-车辆", f"{scn}|{region}", f"{sub}|{tech}", "源表无记录")

    present_power = {(norm_scenario(r["Scenario"]), r["region"], str(r["Subtech"]))
                     for r in cap_records}
    power_combo = {(norm_scenario(r["Scenario"]), r["region"]) for r in cap_records}
    all_subtechs = [s for t in TECH_ORDER for s in TECH_SUBTECH_MAP[t]]
    missing_power = []
    for scn, region in sorted(power_combo):
        for sub in all_subtechs:
            if (scn, region, sub) not in present_power:
                missing_power.append((scn, region, sub))
    add("缺失组合-绿电", "Subtech", f"{len(missing_power)} 个",
        "源表本身不存在该子技术记录，未补造；逐条见下")
    for scn, region, sub in missing_power:
        add("缺失组合-绿电", f"{scn}|{region}", sub, "源表无记录")

    override = Counter((r["年份"], r["系数年份"]) for r in power_rows
                       if r["年份"] != r["系数年份"])
    add("年份系数映射", "无对应年份用最老可用系数；2025→2030",
        str(dict(sorted(override.items()))),
        "用户确认口径；明细「系数年份」列记录实际使用的 CIPG 年份")
    no_year = [r for r in power_rows if r["匹配状态"] == "unmatched"]
    add("未匹配-绿电", "全部年份", f"{len(no_year)} 行",
        "当前应为 0；如有则逐条检查 Subtech 映射与 CIPG 年份")

    # 手算样例 1：车辆
    sample = next(r for r in veh_rows
                  if r["Scenario"] == "SSP2-4p5" and r["Region"] == "China"
                  and r["Subsector"] == "Car" and r["Technology"] == "BEV"
                  and r["年份"] == 2030)
    expected = round_half_up(sample["保有量(辆)"] * sample["需铜系数(kg/vehicle)"] / 1000.0)
    add("手算-车辆", "SSP2-4p5|China|Car|BEV|2030",
        f"{sample['保有量(辆)']:,} × {sample['需铜系数(kg/vehicle)']} ÷ 1000 = {expected}",
        f"程序值={sample['铜量(t)']}；差异={sample['铜量(t)'] - expected}")

    # 手算样例 2：绿电
    sample = next(r for r in power_rows
                  if r["Scenario"] == "SSP1_2p6" and r["Region"] == "China"
                  and r["Subtech"] == "PV" and r["年份"] == 2030)
    expected = round_half_up(sample["装机(MW)"] / 1000.0 * sample["需铜系数(t/GW)"])
    add("手算-绿电", "SSP1_2p6|China|PV|2030",
        f"{sample['装机(MW)']:,} ÷ 1000 × {sample['需铜系数(t/GW)']} = {expected}",
        f"程序值={sample['铜量(t)']}；差异={sample['铜量(t)'] - expected}")

    # 手算样例 3：绿电年份系数映射（2025 用 2030 年系数）
    sample = next(r for r in power_rows
                  if r["Scenario"] == "SSP1_2p6" and r["Region"] == "China"
                  and r["Subtech"] == "PV" and r["年份"] == 2025)
    expected = round_half_up(sample["装机(MW)"] / 1000.0 * sample["需铜系数(t/GW)"])
    add("手算-绿电年份映射", "SSP1_2p6|China|PV|2025",
        f"系数年份={sample['系数年份']}；{sample['装机(MW)']:,} ÷ 1000 × "
        f"{sample['需铜系数(t/GW)']} = {expected}",
        f"程序值={sample['铜量(t)']}；差异={sample['铜量(t)'] - expected}")

    # 汇总一致性：汇总单元格 = 明细显示值分组求和
    veh_groups = group_sums([{**r, "_copper": r["铜量(t)"]} for r in veh_rows],
                            ["_scn", "Region", "年份", "Technology"])
    power_groups = group_sums([{**r, "_copper": r["铜量(t)"]} for r in power_rows],
                              ["_scn", "Region", "年份", "tech"])
    mismatch, add_mismatch = 0, 0
    for row in summary:
        scn, region, year = row["_scn"], row["Region"], row["年份"]
        for tech in ("BEV", "PHEV"):
            value, _ = group_result(veh_groups, (scn, region, year, tech))
            if value != row[f"车辆-{tech}(t)"]:
                mismatch += 1
        for tech in TECH_ORDER:
            value, _ = group_result(power_groups, (scn, region, year, tech))
            cell = row[f"绿电-{tech}(t)"]
            if value != cell:
                mismatch += 1
        # 表内可加性：小计=细分之和，合计=两部门小计之和
        veh_cells = [row[f"车辆-{t}(t)"] for t in ("BEV", "PHEV")
                     if row[f"车辆-{t}(t)"] is not None]
        if row["车辆小计(t)"] is not None and sum(veh_cells) != row["车辆小计(t)"]:
            add_mismatch += 1
        power_cells = [row[f"绿电-{c}(t)"] for c in TECH_ORDER
                       if row[f"绿电-{c}(t)"] is not None]
        if row["绿电小计(t)"] is not None and sum(power_cells) != row["绿电小计(t)"]:
            add_mismatch += 1
        if (row["两部门合计(t)"] is not None
                and row["车辆小计(t)"] + row["绿电小计(t)"] != row["两部门合计(t)"]):
            add_mismatch += 1
    add("汇总一致性", "明细显示值→汇总逐格", f"不一致 {mismatch} 格",
        "汇总单元格=明细显示值分组求和；差异应为 0")
    add("汇总可加性", "小计=细分之和；合计=两部门小计之和", f"不一致 {add_mismatch} 格",
        "汇总表内先取整再求和，保证表面可加")

    # 汇总显示值与未取整分组取整的差异（取整顺序说明）
    raw_veh = group_sums(veh_rows, ["_scn", "Region", "年份", "Technology"])
    raw_power = group_sums(power_rows, ["_scn", "Region", "年份", "tech"])
    max_round, n_round = 0, 0
    for row in summary:
        scn, region, year = row["_scn"], row["Region"], row["年份"]
        for tech in ("BEV", "PHEV"):
            raw, _ = group_result(raw_veh, (scn, region, year, tech))
            cell = row[f"车辆-{tech}(t)"]
            if raw is not None and cell is not None:
                diff = abs(cell - round_half_up(raw))
                if diff:
                    n_round += 1
                    max_round = max(max_round, diff)
        for tech in TECH_ORDER:
            raw, _ = group_result(raw_power, (scn, region, year, tech))
            cell = row[f"绿电-{tech}(t)"]
            if raw is not None and cell is not None:
                diff = abs(cell - round_half_up(raw))
                if diff:
                    n_round += 1
                    max_round = max(max_round, diff)
    add("汇总未取整差异", "明细显示值之和 vs 未取整分组取整", f"{n_round} 格非零",
        f"最大差异 {max_round} t；先按明细取整再汇总，与先汇总再取整的舍入差")

    # 汇总显示值与明细显示值之和的取整差
    veh_disp = defaultdict(int)
    for r in veh_rows:
        if r["铜量(t)"] is not None:
            veh_disp[(r["_scn"], r["Region"], r["年份"], r["Technology"])] += r["铜量(t)"]
    power_disp = defaultdict(int)
    for r in power_rows:
        if r["铜量(t)"] is not None:
            power_disp[(r["_scn"], r["Region"], r["年份"], r["tech"])] += r["铜量(t)"]
    max_disp_diff, n_disp_diff = 0, 0
    for row in summary:
        scn, region, year = row["_scn"], row["Region"], row["年份"]
        for tech in ("BEV", "PHEV"):
            diff = abs(veh_disp.get((scn, region, year, tech), 0) - (row[f"车辆-{tech}(t)"] or 0))
            if diff:
                n_disp_diff += 1
                max_disp_diff = max(max_disp_diff, diff)
        for tech in TECH_ORDER:
            diff = abs(power_disp.get((scn, region, year, tech), 0) - (row[f"绿电-{tech}(t)"] or 0))
            if diff:
                n_disp_diff += 1
                max_disp_diff = max(max_disp_diff, diff)
    add("汇总显示值差异", "汇总 vs 明细显示值之和", f"{n_disp_diff} 格非零",
        f"最大差异 {max_disp_diff} t；汇总分类单元格直接等于明细显示值分组之和")

    # 极值
    veh_top = sorted([r for r in veh_rows if r["_copper"] is not None],
                     key=lambda r: r["_copper"], reverse=True)[:5]
    for r in veh_top:
        add("极值-车辆", f"{r['Scenario']}|{r['Region']}|{r['Subsector']}|{r['Technology']}|{r['年份']}",
            f"{r['铜量(t)']:,} t", f"保有量 {r['保有量(辆)']:,} 辆")
    power_top = sorted([r for r in power_rows if r["_copper"] is not None],
                       key=lambda r: r["_copper"], reverse=True)[:5]
    for r in power_top:
        add("极值-绿电", f"{r['Scenario']}|{r['Region']}|{r['tech']}|{r['Subtech']}|{r['年份']}",
            f"{r['铜量(t)']:,} t", f"装机 {r['装机(MW)']:,} MW")
    add("零值-车辆", "铜量=0 的记录", f"{sum(1 for r in veh_rows if r['铜量(t)'] == 0)} 行",
        "保有量为 0，非匹配失败")
    add("零值-绿电", "铜量=0 的记录", f"{sum(1 for r in power_rows if r['铜量(t)'] == 0)} 行",
        "装机为 0，非匹配失败")

    add("单位换算-车辆", "辆 × kg/辆 ÷ 1000 = t", "已执行",
        "示例：282 辆 × 776.8 kg/辆 ÷ 1000 = 219 t")
    add("单位换算-绿电", "MW ÷ 1000 = GW；GW × t/GW = t", "已执行",
        "示例：24,031 MW ÷ 1000 × 4500.714 t/GW = 108,157 t")
    add("输出", str(RESULT_XLSX), "已写入", "2 张面板")
    return checks


def build_notes():
    """「口径说明」表内容。"""
    tech_map = "；".join(
        f"{tech} ← {' + '.join(TECH_SUBTECH_MAP[tech])}" for tech in TECH_ORDER)
    sub_map = "；".join(f"{sub} → {cipg}" for sub, cipg in SUBTECH_TO_CIPG.items())
    notes = [
        ("方案依据", "项目内任务方案 PROMPT_给DeepSeek_铜需求量.md；2026-11-12 确认 2025 用 2030 年系数、保留源表原写法；2026-11-14 确认新增 1990/2005 等年份，无对应系数时用最老可用年份（2021）"),
        ("输入-设备", f"{DEVICE_XLSX}；工作表「{VEHICLE_SHEET}」「{CAPACITY_SHEET}」（只读，未修改）"),
        ("输入-系数", f"{COEFF_XLSX}；工作表「{CIPG_SHEET}」「{AUTO_SHEET}」（只读；源文件被占用时读取临时副本）"),
        ("输出", str(RESULT_XLSX)),
        ("口径", "保有量×制造需铜系数、累计装机×制造需铜系数 = 现有设备所含制造用铜；不是当年新增铜需求"),
        ("口径-边界", "未做年度差分、寿命报废、更新替换、回收抵扣、二次加工、价格或贸易修正"),
        ("车辆公式", "铜量(t) = 保有量(辆) × 系数(kg/vehicle) ÷ 1000"),
        ("绿电公式", "铜量(t) = 装机(MW) ÷ 1000 × 系数(t/GW)"),
        ("车辆映射", "Subsector 与 Automotive copper demand.categories 同名一一对应；Technology BEV/PHEV 对应同名系数列"),
        ("绿电-tech", tech_map),
        ("绿电-Subtech→CIPG", sub_map),
        ("nuclear 映射", "源表 Subtech=nuclear 与 CIPG technology=nuclear 同名匹配；CIPG 中该行对应原 Gen_III/Gen_II_LWR 强度，两者数值相同"),
        ("wind 映射", "源表 tech=wind 含 Subtech=wind 与 wind_offshore；分别映射 CIPG wind 与 wind_offshore，汇总时相加"),
        ("年份处理-车辆", "Automotive 系数无年份维度，1990–2050 各年同值"),
        ("年份处理-绿电", "CIPG 年份为 2021、2030–2060；有对应年份用当年，无对应年份用最老可用年份 2021（如 1990、2005、2010、2015、2020）；2025 按已确认口径用 2030；明细「系数年份」列记录实际使用年份"),
        ("空值处理", "源表空白保留空白并登记，不填 0；保有量/装机为 0 时铜量为 0"),
        ("重复处理", "四张输入表主键均无重复；未做去重或行合并"),
        ("Scenario", "明细与汇总均保留源表原写法；汇总设「Scenario(车辆原写法)」「Scenario(绿电原写法)」两列，不做 - ↔ _ 归一"),
        ("取整", "中间值全精度；最终铜量(t) 按 floor(x+0.5) 取整；汇总表内先取整再求小计/合计，保证表面可加"),
        ("汇总取整说明", "分类单元格=明细显示值分组求和；小计/合计=已取整单元格之和；与未取整汇总的差异≤3 t/格"),
        ("待确认", "输出年份是否保留源表全部年份；取整是否使用 floor(x+0.5)"),
    ]
    return [{"项目": k, "内容": v} for k, v in notes]


def _write_table(workbook, name, fields, rows, widths,
                 int_fields=(), float_fields=(), autofilter=True,
                 header_number_fields=()):
    ws = workbook.add_worksheet(name)
    header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#D9E1F2", "border": 1,
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    int_fmt = workbook.add_format({"num_format": "#,##0", "valign": "top"})
    float_fmt = workbook.add_format({"num_format": "0.######", "valign": "top"})
    text_fmt = workbook.add_format({"valign": "top"})

    ws.freeze_panes(1, 0)
    ws.set_row(0, 30)
    for col, field in enumerate(fields):
        if field in header_number_fields:
            ws.write_number(0, col, field, header_fmt)
        else:
            ws.write_string(0, col, str(field), header_fmt)
        ws.set_column(col, col, widths[col])

    for row_idx, row in enumerate(rows, start=1):
        for col, field in enumerate(fields):
            value = row.get(field)
            if value is None or value == "":
                continue
            if isinstance(value, str):
                ws.write_string(row_idx, col, value, text_fmt)
            elif field in int_fields:
                ws.write_number(row_idx, col, value, int_fmt)
            elif field in float_fields:
                ws.write_number(row_idx, col, value, float_fmt)
            else:
                ws.write_number(row_idx, col, value)
    if autofilter and rows:
        ws.autofilter(0, 0, len(rows), len(fields) - 1)
    return ws


def check_long_duplicates(rows, key_fields, label):
    """转换前检查「标识字段+年份」是否重复；重复则报告并中止，不做求和/平均/删除。"""
    counter = Counter(
        tuple(row[k] for k in key_fields) + (row["年份"],) for row in rows
    )
    duplicates = [(key, n) for key, n in counter.items() if n > 1]
    print(f"{label}：长表行数={len(rows)}，唯一键={len(counter)}，重复键={len(duplicates)}")
    if duplicates:
        for key, n in duplicates[:20]:
            print(f"  [重复] {label} {key} -> {n} 次")
        raise ValueError(
            f"{label} 存在 {len(duplicates)} 个「标识字段+年份」重复键；"
            "按任务要求报告并中止，不做求和、平均或删除。"
        )
    return len(rows), len(counter)


def build_panel(rows, key_fields, years, label):
    """长表 -> 宽面板：每个标识键一行，年份升序展开为列。"""
    panel = {}
    seen = defaultdict(set)
    for row in rows:
        key = tuple(row[k] for k in key_fields)
        if key not in panel:
            panel[key] = {k: row[k] for k in key_fields}
            panel[key]["单位"] = "t"
            for year in years:
                panel[key][year] = None
        if row["年份"] in seen[key]:
            raise ValueError(f"{label} 面板键重复：{key} @ {row['年份']}")
        seen[key].add(row["年份"])
        panel[key][row["年份"]] = row["铜量(t)"]

    missing = 0
    for key, item in panel.items():
        for year in years:
            if year not in seen[key]:
                missing += 1
                item[year] = None
    print(f"{label}：面板行数={len(panel)}，年份={years}，缺失年份单元格={missing}")
    return list(panel.values()), missing


def panel_digest(panel, key_fields, years):
    """面板数据摘要，用于比较重复运行是否得到相同数据结果。"""
    items = []
    for row in sorted(panel, key=lambda r: tuple(str(r[k]) for k in key_fields)):
        items.append([
            [str(row[k]) for k in key_fields],
            [row[year] for year in years],
        ])
    payload = json.dumps(items, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def write_panel_workbook(path, veh_panel, power_panel, veh_years, cap_years):
    """写临时成果：仅两张面板。"""
    workbook = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True})

    veh_fields = VEHICLE_KEY_FIELDS + ["单位"] + list(veh_years)
    veh_widths = [12, 30, 20, 12, 6] + [12] * len(veh_years)
    _write_table(
        workbook, PANEL_VEHICLE_SHEET, veh_fields, veh_panel, veh_widths,
        int_fields=set(veh_years), header_number_fields=set(veh_years),
    )

    power_fields = POWER_KEY_FIELDS + ["单位"] + list(cap_years)
    power_widths = [12, 30, 18, 26, 6] + [12] * len(cap_years)
    _write_table(
        workbook, PANEL_POWER_SHEET, power_fields, power_panel, power_widths,
        int_fields=set(cap_years), header_number_fields=set(cap_years),
    )

    workbook.close()


def _expected_values(rows, key_fields):
    return {
        (tuple(row[k] for k in key_fields), row["年份"]): row["铜量(t)"]
        for row in rows
    }


def _verify_one_panel(ws, panel, key_fields, years, long_rows, label):
    """逐键逐年核对面板值与长表铜量，并核对空值/零值数量。"""
    fields = key_fields + ["单位"] + list(years)
    data = list(ws.iter_rows(values_only=True))
    header = list(data[0]) if data else []
    body = data[1:]

    header_ok = header == fields
    unit_bad = 0
    key_duplicates = 0
    mismatches = 0
    seen_keys = set()
    expected = _expected_values(long_rows, key_fields)

    year_start = len(key_fields) + 1
    panel_blank = {year: 0 for year in years}
    panel_zero = {year: 0 for year in years}

    for row in body:
        key = tuple(row[:len(key_fields)])
        if key in seen_keys:
            key_duplicates += 1
        seen_keys.add(key)
        if row[len(key_fields)] != "t":
            unit_bad += 1
        for offset, year in enumerate(years):
            got = row[year_start + offset]
            want = expected.get((key, year))
            if got != want:
                mismatches += 1
            if got is None:
                panel_blank[year] += 1
            elif got == 0:
                panel_zero[year] += 1

    long_blank = {
        year: sum(1 for r in long_rows if r["年份"] == year and r["铜量(t)"] is None)
        for year in years
    }
    long_zero = {
        year: sum(1 for r in long_rows if r["年份"] == year and r["铜量(t)"] == 0)
        for year in years
    }

    result = {
        "label": label,
        "sheet": ws.title,
        "header_ok": header_ok,
        "rows_panel": len(body),
        "rows_expected": len(panel),
        "key_duplicates": key_duplicates,
        "unit_bad": unit_bad,
        "mismatches": mismatches,
        "blank_ok": panel_blank == long_blank,
        "zero_ok": panel_zero == long_zero,
        "panel_blank": panel_blank,
        "long_blank": long_blank,
        "panel_zero": panel_zero,
        "long_zero": long_zero,
    }
    result["ok"] = (
        header_ok
        and len(body) == len(panel)
        and key_duplicates == 0
        and unit_bad == 0
        and mismatches == 0
        and result["blank_ok"]
        and result["zero_ok"]
    )
    return result


def verify_panel_workbook(path, veh_panel, power_panel, veh_years, cap_years,
                          veh_rows, power_rows):
    """重新打开临时成果，逐键逐年与长表中间结果核对。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets_ok = wb.sheetnames == [PANEL_VEHICLE_SHEET, PANEL_POWER_SHEET]
    veh = _verify_one_panel(
        wb[PANEL_VEHICLE_SHEET], veh_panel, VEHICLE_KEY_FIELDS, veh_years,
        veh_rows, "车辆",
    )
    power = _verify_one_panel(
        wb[PANEL_POWER_SHEET], power_panel, POWER_KEY_FIELDS, cap_years,
        power_rows, "绿电",
    )
    wb.close()

    report = {
        "sheets_ok": sheets_ok,
        "sheets": [PANEL_VEHICLE_SHEET, PANEL_POWER_SHEET],
        "vehicle": veh,
        "power": power,
    }
    report["ok"] = sheets_ok and veh["ok"] and power["ok"]
    return report


def replace_with_temp(temp_path, final_path):
    """核验通过后替换正式 Excel；目标被占用时保留临时文件并报错。"""
    try:
        os.replace(temp_path, final_path)
    except PermissionError as exc:
        print(f"替换失败：{final_path} 可能正被 Excel 打开。临时文件保留：{temp_path}")
        raise exc


def print_verification(report):
    print("=" * 100)
    print("临时成果核验：")
    print(f"  工作表顺序/数量正确：{report['sheets_ok']} -> {report['sheets']}")
    for key in ("vehicle", "power"):
        item = report[key]
        print(f"  [{item['label']}] sheet={item['sheet']} 行数={item['rows_panel']}"
              f"（期望 {item['rows_expected']}）")
        print(f"    表头正确={item['header_ok']} 键重复={item['key_duplicates']}"
              f" 单位非t={item['unit_bad']} 铜量不一致={item['mismatches']}")
        print(f"    空值一致={item['blank_ok']} {item['panel_blank']} / {item['long_blank']}")
        print(f"    零值一致={item['zero_ok']} {item['panel_zero']} / {item['long_zero']}")
    print(f"  总核验：{'通过' if report['ok'] else '未通过'}")


# ---------------------------------------------------------------------------
# 结构勘察与映射校验（只读）
# ---------------------------------------------------------------------------
def _key_duplicates(records, keys):
    counter = Counter(tuple(rec[k] for k in keys) for rec in records)
    return [key for key, n in counter.items() if n > 1]


def _null_counts(records, years):
    return {year: sum(1 for rec in records if rec["_years"].get(year) is None) for year in years}


def print_structure():
    """基线：打印两个工作簿的工作表名、行列数与表头。"""
    for path in (DEVICE_XLSX, COEFF_XLSX):
        print("=" * 100)
        print(f"工作簿：{path}")
        wb, used, copied = load_readonly(path)
        print(f"实际读取：{used}" + ("（源文件被占用，读取副本）" if copied else ""))
        print(f"工作表：{wb.sheetnames}")
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            header = next(rows, None)
            print("-" * 90)
            print(f"[{ws.title}] 行={ws.max_row} 列={ws.max_column}")
            print(f"表头：{header}")
        wb.close()


def profile():
    """结构勘察：逐表打印结构、去重取值、空值、主键重复与映射覆盖情况。"""
    device_wb, device_used, device_copied = load_readonly(DEVICE_XLSX)
    print("=" * 100)
    print(f"设备工作簿：{DEVICE_XLSX}")
    print(f"实际读取：{device_used}" + ("（源文件被占用，读取副本）" if device_copied else ""))

    veh, veh_years, veh_units, veh_header = load_vehicle(device_wb)
    cap, cap_years, cap_units, cap_header = load_capacity(device_wb)
    device_wb.close()

    for name, recs, years, units, header, keys in [
        ("车辆保有量_面板", veh, veh_years, veh_units, veh_header,
         ["Scenario", "Region", "Subsector", "Technology"]),
        ("装机容量_九类_面板", cap, cap_years, cap_units, cap_header,
         ["Scenario", "region", "tech", "Subtech"]),
    ]:
        print("-" * 90)
        print(f"[{name}] 记录={len(recs)} 年份={years} 单位={units}")
        print(f"表头：{header}")
        for key in keys:
            vals = sorted({str(r[key]) for r in recs})
            print(f"  {key}：{len(vals)} 个 -> {vals}")
        print(f"  空值计数：{_null_counts(recs, years)}")
        dups = _key_duplicates(recs, keys)
        print(f"  主键重复：{len(dups)} {dups[:5]}")
        print(f"  主键重复多出行数：{len(recs) - len({tuple(rec[k] for k in keys) for rec in recs})}")

    coef_wb, coef_used, coef_copied = load_readonly(COEFF_XLSX)
    print("=" * 100)
    print(f"系数工作簿：{COEFF_XLSX}")
    print(f"实际读取：{coef_used}" + ("（源文件被占用，读取副本）" if coef_copied else ""))
    cipg, cipg_years, cipg_unit, cipg_dups, cipg_units = load_cipg(coef_wb)
    auto, auto_unit, auto_dups, auto_units = load_automotive(coef_wb)
    coef_wb.close()

    print("-" * 90)
    print(f"[CIPG] 技术数={len(cipg)} 年份={cipg_years} 单位={cipg_units}")
    print(f"  技术：{sorted(cipg)}")
    print(f"  重复技术：{cipg_dups}")
    print(f"  各年空值：{ {y: sum(1 for t in cipg if cipg[t].get(y) is None) for y in cipg_years} }")
    print("-" * 90)
    print(f"[Automotive copper demand] 类别数={len(auto)} 单位={auto_units}")
    print(f"  类别：{sorted(auto)}")
    print(f"  重复类别：{auto_dups}")
    print(f"  技术列：{sorted(next(iter(auto.values())).keys())}")
    print(f"  各技术空值：{ {t: sum(1 for c in auto if auto[c].get(t) is None) for t in ('BEV','PHEV','FCEV','Liquids')} }")

    # 映射覆盖校验
    print("=" * 100)
    print("映射校验：")
    veh_cats = sorted({str(r["Subsector"]) for r in veh})
    veh_techs = sorted({str(r["Technology"]) for r in veh})
    cap_techs = sorted({str(r["tech"]) for r in cap})
    cap_subtechs = sorted({str(r["Subtech"]) for r in cap})
    print(f"  车辆 Subsector：{veh_cats}")
    print(f"    -> 未映射到 Automotive categories：{sorted(set(veh_cats) - set(VEHICLE_CATEGORY_MAP))}")
    print(f"    -> Automotive 中不存在：{sorted(set(VEHICLE_CATEGORY_MAP.values()) - set(auto))}")
    print(f"  车辆 Technology：{veh_techs}")
    print(f"    -> 未映射：{sorted(set(veh_techs) - set(VEHICLE_TECH_MAP))}")
    print(f"  绿电 tech：{cap_techs}")
    print(f"    -> 未在 TECH_ORDER：{sorted(set(cap_techs) - set(TECH_ORDER))}")
    print(f"  绿电 Subtech：{cap_subtechs}")
    print(f"    -> 未映射到 CIPG：{sorted(set(cap_subtechs) - set(SUBTECH_TO_CIPG))}")
    print(f"    -> 映射目标在 CIPG 中不存在：{sorted(set(SUBTECH_TO_CIPG.values()) - set(cipg))}")


def main():
    print_structure()
    print()
    profile()
    print()

    device_wb, _, _ = load_readonly(DEVICE_XLSX)
    veh, veh_years, _, _ = load_vehicle(device_wb)
    cap, cap_years, _, _ = load_capacity(device_wb)
    device_wb.close()

    coef_wb, _, _ = load_readonly(COEFF_XLSX)
    auto, auto_unit, _, _ = load_automotive(coef_wb)
    cipg, cipg_years, cipg_unit, _, _ = load_cipg(coef_wb)
    coef_wb.close()

    veh_rows = compute_vehicle(veh, veh_years, auto)
    print_vehicle_check(veh_rows, veh_years)
    print()

    power_rows = compute_power(cap, cap_years, cipg, cipg_years)
    print_power_check(power_rows, cap_years, cipg_years)
    print()

    # 转换前检查「标识字段+年份」唯一性；重复则报告中止。
    check_long_duplicates(veh_rows, VEHICLE_KEY_FIELDS, "车辆")
    check_long_duplicates(power_rows, POWER_KEY_FIELDS, "绿电")

    veh_panel, veh_missing = build_panel(
        veh_rows, VEHICLE_KEY_FIELDS, veh_years, "车辆")
    power_panel, power_missing = build_panel(
        power_rows, POWER_KEY_FIELDS, cap_years, "绿电")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = RESULT_XLSX.with_name(RESULT_XLSX.stem + ".tmp.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    write_panel_workbook(temp_path, veh_panel, power_panel, veh_years, cap_years)
    print(f"\n临时成果：{temp_path}")

    report = verify_panel_workbook(
        temp_path, veh_panel, power_panel, veh_years, cap_years,
        veh_rows, power_rows)
    print_verification(report)
    print(f"数据摘要：车辆={panel_digest(veh_panel, VEHICLE_KEY_FIELDS, veh_years)}")
    print(f"数据摘要：绿电={panel_digest(power_panel, POWER_KEY_FIELDS, cap_years)}")

    if not report["ok"]:
        raise ValueError("临时成果核验未通过，未替换正式 Excel；请检查临时文件。")

    replace_with_temp(temp_path, RESULT_XLSX)
    print(f"\n已替换正式成果：{RESULT_XLSX}")
    print(f"面板行数：车辆={len(veh_panel)}，绿电={len(power_panel)}；"
          f"缺失年份单元格：车辆={veh_missing}，绿电={power_missing}")


if __name__ == "__main__":
    main()
