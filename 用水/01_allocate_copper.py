#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
01_allocate_copper.py
=====================
把"情景 × 年份 × 工艺路线"的铜需求，按**对应年份**的矿山铜产量权重分配到矿山。

业务口径（详见 README.md）：
  * 需求 D[s,t,r] 直接取自既有需求面板，不重新预测；
  * 权重只用 t 年的矿山铜产量（含铜金属量），不使用参考产能截断；
  * 参考产能仅用于展示"分配量 / 参考产能"比值（同为 t Cu/年）；
  * 缺失产量与零产量严格区分；缺失不插值、不外推、不向前/向后填充；
  * 同一矿山的多条路线用路线份额拆分，不重复使用未拆分总产量。

运行：python 01_allocate_copper.py
输出：intermediate/01_铜需求分配.csv、01_组级校验.csv、01_数据质量.csv
      以及最终 Excel 的第一张工作表（用水核算表由 02_calculate_water.py 生成）
"""

from __future__ import annotations

import hashlib
import math
import re
import shutil
import sys
import tempfile
import uuid
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# %% Chunk 1：配置与字段映射
# 功能：集中定义路径、工作表名、字段名、单位、路线判定规则、地区映射与匹配规则。
# 输入：无（常量配置）。
# 输出：供后续所有 chunk 使用的常量与格式定义。

BASE_DIR = Path(__file__).resolve().parent
INPUT_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx")
DEMAND_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\MFA\Output\冶炼路线与铜精矿需求_面板.xlsx")
OUTPUT_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\MFA\Output\铜需求分配与用水核算.xlsx")
INTERMEDIATE_DIR = BASE_DIR / "intermediate"
ALLOC_CSV = INTERMEDIATE_DIR / "01_铜需求分配.csv"
GROUP_CHECK_CSV = INTERMEDIATE_DIR / "01_组级校验.csv"
QUALITY_CSV = INTERMEDIATE_DIR / "01_数据质量.csv"

# --- 工作表 ---
SHEET_MINE = "Mine Production"
SHEET_DEMAND = "路线分配_面板"
SHEET_SUMMARY = "年度情景汇总"
SHEET_WATER = "矿山用水系数"

# --- Mine Production 字段（原表列名，含全角括号） ---
F_NAME = "矿山项目名称"
F_COUNTRY = "国家"
F_LAT = "纬度"
F_LON = "经度"
F_START = "投产年（实际或预计）"
F_END = "退役年（实际或预计）"
F_PRODUCT = "产品实际情况（标普原表）"
F_CAPACITY = "设计产品产能（吨铜/年，原表口径）"
YEAR_COL_PATTERN = re.compile(r"^(\d{4})年产量（吨铜）$")

# --- 需求面板字段 ---
F_DEM_SCEN = "需求情景"
F_DEM_REC = "回收情景"
F_DEM_TRADE = "贸易情景"
F_DEM_YEAR = "年份"
F_DEM_FIRE = "火法铜_tCu"
F_DEM_HYDRO = "湿法铜_tCu"
F_DEM_TOTAL = "原生铜需求_tCu"
F_DEM_CONC = "铜精矿实物量_t"

# --- 年份与容差 ---
YEAR_MIN = 2021            # 需求年份下限（与需求面板一致）
YEAR_MAX = 2050            # 需求年份上限
TOL_REL = 1e-9             # 守恒/权重容差（相对量）
USE_OPERATING_WINDOW = False  # 是否用投产/退役年硬筛；默认否（见 Chunk 3 说明）

# --- 工艺路线 ---
ROUTE_FIRE = "火法"
ROUTE_HYDRO = "湿法"
ROUTE_UNKNOWN = "未知"
ROUTE_ORDER = {ROUTE_FIRE: 0, ROUTE_HYDRO: 1, ROUTE_UNKNOWN: 2}

# --- 情景排序 ---
SCEN_ORDER = {"SSP1-2.6": 0, "SSP2-4.5": 1, "SSP5-8.5": 2}
REC_ORDER = {"基准": 0, "技术进步": 1, "政策推动": 2}
TRADE_ORDER = {"不开放": 0, "全球开放": 1}

# --- 国家 → 7 地区映射（与"矿山用水系数"表头完全一致；联合国地理分区，跨洲国家按下列约定） ---
REGION_MAP = {
    "大洋洲": ["Australia", "Papua New Guinea"],
    "非洲": ["Botswana", "Dem. Rep. Congo", "Mauritania", "Morocco",
             "Namibia", "South Africa", "Uganda", "Zambia"],
    "欧洲": ["Albania", "Bulgaria", "Finland", "North Macedonia", "Poland",
             "Portugal", "Romania", "Serbia", "Spain", "Sweden", "Russia"],
    "北美洲": ["Canada", "Mexico", "USA"],
    "中国": ["China"],
    "其他亚洲": ["Armenia", "Azerbaijan", "Georgia", "Türkiye", "Kazakhstan",
                 "Uzbekistan", "Kyrgyzstan", "Iran", "Israel", "Oman",
                 "Saudi Arabia", "Pakistan", "India", "Indonesia", "Laos",
                 "Mongolia", "Myanmar", "Philippines", "Vietnam"],
    "拉丁美洲": ["Argentina", "Bolivia", "Brazil", "Chile", "Colombia",
                 "Dominican Republic", "Ecuador", "Panama", "Peru"],
}
COUNTRY_TO_REGION = {c: r for r, cs in REGION_MAP.items() for c in cs}

# --- 输出列定义：(列名, 列宽, 数字格式键) ---
COL_SCEN_ID = "情景ID"
COL_YEAR = "年份"
COL_MINE_ID = "矿山ID"
COL_NAME = "矿山项目名称"
COL_COUNTRY = "国家"
COL_REGION = "地区"
COL_ROUTE = "工艺路线"
COL_W = "路线产量份额"
COL_P = "当年参考铜产量_tCu"
COL_A = "分配铜需求量_tCu"
COL_CAP = "参考产能_tCu每年"
COL_RATIO = "分配量除参考产能"
COL_D = "组内路线需求总量_tCu（组级·勿逐行求和）"
COL_PSUM = "组内参考产量合计_tCu（组级·勿逐行求和）"
COL_UNALLOC = "未分配需求_tCu（组级）"
COL_CONC = "组内铜精矿实物量_t（组级·勿逐行求和）"
COL_STATUS = "分配状态"
COL_FLAGS = "数据质量标记"
COL_SOURCE = "数据来源"
COL_NOTE = "备注"

SHEET1_COLUMNS = [
    (COL_SCEN_ID, 22, "text"),
    (F_DEM_SCEN, 12, "text"),
    (F_DEM_REC, 11, "text"),
    (F_DEM_TRADE, 11, "text"),
    (COL_YEAR, 7, "year"),
    (COL_MINE_ID, 9, "text"),
    (COL_NAME, 28, "text"),
    (COL_COUNTRY, 18, "text"),
    (COL_REGION, 10, "text"),
    (COL_ROUTE, 9, "text"),
    (COL_W, 13, "w"),
    (COL_P, 18, "tcu"),
    (COL_A, 18, "tcu"),
    (COL_CAP, 18, "tcu"),
    (COL_RATIO, 16, "ratio"),
    (COL_D, 26, "tcu_group"),
    (COL_PSUM, 26, "tcu_group"),
    (COL_UNALLOC, 20, "tcu_group"),
    (COL_STATUS, 14, "text"),
    (COL_FLAGS, 30, "text"),
    (COL_SOURCE, 26, "text"),
    (COL_NOTE, 26, "text"),
]
SHEET1_HEADERS = [c[0] for c in SHEET1_COLUMNS]
# 中间 CSV 在表 1 列基础上追加程序二所需的组级精矿实物量
CSV_COLUMNS = SHEET1_COLUMNS + [(COL_CONC, 26, "tcu_group")]
CSV_HEADERS = [c[0] for c in CSV_COLUMNS]


# %% Chunk 2：数据读取与宽表转长表
# 功能：稳健读取工作簿（文件被 Excel 占用时回退到临时副本），并把 41 个年份列转为长表。
# 输入：INPUT_XLSX、DEMAND_XLSX。
# 输出：矿山主数据（含矿山ID/地区/路线/标记）与矿山 × 年份长表。


def sha256_file(path: Path) -> str:
    """返回文件 sha256（用于证明输入未被改动）。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_excel_robust(path: Path, sheet_name: str) -> pd.DataFrame:
    """读取工作表；若源文件被 Excel 独占（PermissionError），复制临时副本读取，原文件不改动。"""
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / f"dsh_read_{uuid.uuid4().hex}.xlsx"
        shutil.copyfile(path, tmp)
        try:
            print(f"[提示] {path.name} 正被其他程序占用，已读取临时副本；原始文件未改动。")
            return pd.read_excel(tmp, sheet_name=sheet_name)
        finally:
            tmp.unlink(missing_ok=True)


def load_mines() -> tuple[pd.DataFrame, list[tuple[str, int]]]:
    """读取 Mine Production，标准化字段、判定工艺路线、映射地区、生成唯一矿山ID。"""
    mp = read_excel_robust(INPUT_XLSX, SHEET_MINE)
    mp.columns = [str(c).strip() for c in mp.columns]
    required = [F_NAME, F_COUNTRY, F_LAT, F_LON, F_START, F_END, F_PRODUCT, F_CAPACITY]
    missing = [c for c in required if c not in mp.columns]
    if missing:
        raise ValueError(f"Mine Production 缺少必需字段：{missing}")
    # 源行号先按读入顺序固定（与 Excel 行号一致：表头占第 1 行），再删全空行
    mp["源行号"] = np.arange(2, len(mp) + 2)
    mp = mp[~mp.drop(columns=["源行号"]).isna().all(axis=1)].reset_index(drop=True)

    # --- 年份列识别（2010–2050，必须连续） ---
    year_cols: list[tuple[str, int]] = []
    for col in mp.columns:
        m = YEAR_COL_PATTERN.match(str(col).strip())
        if m:
            year_cols.append((col, int(m.group(1))))
    year_cols.sort(key=lambda x: x[1])
    years_found = [y for _, y in year_cols]
    missing_years = [y for y in range(YEAR_MIN, YEAR_MAX + 1) if y not in years_found]
    if missing_years:
        raise ValueError(
            f"需求年份 {YEAR_MIN}–{YEAR_MAX} 在产量列中缺失 {missing_years}；"
            "按口径不得插值/外推，请补充原始数据后重跑。"
        )

    # --- 唯一标识（名称+国家存在 9 对重复记录，不能作主键） ---
    mp["矿山ID"] = [f"MP{r:04d}" for r in mp["源行号"]]

    # --- 工艺路线判定 R1–R4（当前"产品实际情况"标签） ---
    label = mp[F_PRODUCT].astype("string").str.strip()
    label_u = label.str.upper()
    name_u = mp[F_NAME].astype("string").str.strip().str.upper()
    is_conc = label_u.str.contains("CONCENTRATE", na=False)          # R1
    is_cath = label_u.str.contains("CATHODE", na=False)              # R2
    is_blank = label.isna() | (label == "")
    is_sxew_name = name_u.str.contains("SX-EW", na=False)
    route = np.select(
        [is_conc, is_cath, is_blank & is_sxew_name],                  # R3 兜底
        [ROUTE_FIRE, ROUTE_HYDRO, ROUTE_HYDRO],
        default=ROUTE_UNKNOWN,                                        # R4
    )
    mp["工艺路线"] = route
    mp["产品标签"] = label.fillna("")

    # --- 地区映射（必须 100% 覆盖） ---
    country = mp[F_COUNTRY].astype("string").str.strip()
    mp["地区"] = country.map(COUNTRY_TO_REGION)
    if mp["地区"].isna().any():
        bad = sorted(mp.loc[mp["地区"].isna(), F_COUNTRY].astype(str).unique())
        raise ValueError(f"以下国家未配置地区映射（不得自动套用其他地区）：{bad}")

    # --- 数值化 ---
    mp["参考产能_tCu每年"] = pd.to_numeric(mp[F_CAPACITY], errors="coerce")
    mp["投产年"] = pd.to_numeric(mp[F_START], errors="coerce")
    mp["退役年"] = pd.to_numeric(mp[F_END], errors="coerce")

    # --- 矿山级数据质量标记 ---
    key = mp[F_NAME].astype("string").str.strip() + "|" + country
    dup = key.duplicated(keep=False)
    cap_missing = mp["参考产能_tCu每年"].isna()
    r3_mask = mp["产品标签"].eq("") & mp[F_NAME].astype(str).str.upper().str.contains("SX-EW", na=False)
    base_flags = []
    for f_dup, f_cap, f_r3 in zip(dup, cap_missing, r3_mask):
        flags = []
        if f_dup:
            flags.append("疑似重复记录（同名同国）")
        if f_cap:
            flags.append("参考产能缺失")
        if f_r3:
            flags.append("标签空白，按名称 SX-EW 判定为湿法")
        base_flags.append(";".join(flags))
    mp["基础标记"] = base_flags

    out_cols = ["矿山ID", "源行号", F_NAME, F_COUNTRY, "地区", "工艺路线",
                "参考产能_tCu每年", "产品标签", "投产年", "退役年", "基础标记"]
    keep = out_cols + [c for c, _ in year_cols]
    out = mp[keep].rename(columns={F_NAME: COL_NAME, F_COUNTRY: COL_COUNTRY})
    return out, year_cols


def build_long(mines: pd.DataFrame, year_cols: list[tuple[str, int]]) -> pd.DataFrame:
    """宽表转长表：矿山 × 年份，保留缺失与真实零的区别，仅保留 P>0 的参与记录。"""
    id_cols = ["矿山ID", "源行号", COL_NAME, COL_COUNTRY, "地区", "工艺路线",
               "参考产能_tCu每年", "产品标签", "投产年", "退役年", "基础标记"]
    long = mines.melt(
        id_vars=id_cols,
        value_vars=[c for c, _ in year_cols],
        var_name="年份列",
        value_name="P_raw",
    )
    long["年份"] = long["年份列"].astype(str).str.extract(r"(\d{4})").astype(int)
    long = long[(long["年份"] >= YEAR_MIN) & (long["年份"] <= YEAR_MAX)].copy()
    long["P"] = pd.to_numeric(long["P_raw"], errors="coerce")

    # --- 运营期标记：只标记不筛选（源表投产/退役年缺失 52/312 行，且与产量冲突量 <0.7%） ---
    early = long["投产年"].notna() & (long["年份"] < long["投产年"])
    late = long["退役年"].notna() & (long["年份"] > long["退役年"])
    long["运营期标记"] = np.select(
        [early, late],
        ["运营期外(早于投产年)", "运营期外(晚于退役年)"],
        default="",
    )

    # --- 合并矿山级与年度级标记 ---
    long["数据质量标记"] = [
        ";".join([m for m in (a, b) if m])
        for a, b in zip(long["基础标记"].fillna(""), long["运营期标记"])
    ]

    # --- 缺失 ≠ 零：P 为 NaN 的记录不进入分母，P==0 明确排除；只保留 P>0 ---
    base = long[long["P"] > 0].copy()
    # 路线未知的矿山无法归入火法/湿法，显式剔除（产量占比在数据质量表中单独披露），
    # 避免内连接时被静默丢弃。
    base = base[base["工艺路线"].isin([ROUTE_FIRE, ROUTE_HYDRO])].copy()
    base = base.rename(columns={"P": COL_P})
    return base


# %% Chunk 3：字段、单位和年份检查
# 功能：核对行数、主键、年份覆盖、路线/地区分布、缺失与零产量、重复记录、运营期异常。
# 输入：mines、long、需求面板。
# 输出：数据质量记录（list[dict]）。


def collect_quality(mines: pd.DataFrame, year_cols: list[tuple[str, int]],
                    demand: pd.DataFrame) -> list[dict]:
    """生成数据质量记录，供 README 与 01_数据质量.csv 使用。"""
    q: list[dict] = []

    def add(item, value, note=""):
        q.append({"项目": item, "数值": value, "说明": note})

    # 矿山主数据
    add("矿山记录数（去全空行后）", len(mines), "Mine Production 数据行")
    add("矿山唯一ID数", mines["矿山ID"].nunique(), "矿山ID = MP+源行号")
    add("年份列数", len(year_cols), f"{year_cols[0][1]}–{year_cols[-1][1]}")
    add("国家数", mines[COL_COUNTRY].nunique(), "")
    add("地区数", mines["地区"].nunique(), "7 地区全覆盖")

    vc = mines["工艺路线"].value_counts()
    for r in [ROUTE_FIRE, ROUTE_HYDRO, ROUTE_UNKNOWN]:
        add(f"路线={r} 矿山数", int(vc.get(r, 0)), "")

    # 重复记录
    dup_key = mines[COL_NAME].astype("string").str.strip() + "|" + mines[COL_COUNTRY].astype(str)
    add("同名同国重复记录组数", int(dup_key[dup_key.duplicated(keep=False)].nunique()),
        "按口径保留为独立记录，矿山ID 区分")

    # 产能缺失
    add("参考产能缺失矿山数", int(mines["参考产能_tCu每年"].isna().sum()), "仅影响比值列，不影响分配")

    # 运营期字段缺失
    add("投产年缺失矿山数", int(mines["投产年"].isna().sum()), "")
    add("退役年缺失矿山数", int(mines["退役年"].isna().sum()), "")

    # 需求窗口内的产量统计（缺失/零/运营期）
    long_all = mines.melt(
        id_vars=["矿山ID", "工艺路线", "投产年", "退役年"],
        value_vars=[c for c, _ in year_cols],
        var_name="年份列", value_name="P_raw",
    )
    long_all["年份"] = long_all["年份列"].astype(str).str.extract(r"(\d{4})").astype(int)
    long_all = long_all[(long_all["年份"] >= YEAR_MIN) & (long_all["年份"] <= YEAR_MAX)]
    long_all["P"] = pd.to_numeric(long_all["P_raw"], errors="coerce")
    na_per_year = long_all.groupby("年份")["P"].apply(lambda s: int(s.isna().sum()))
    zero_per_year = long_all.groupby("年份")["P"].apply(lambda s: int((s == 0).sum()))
    add("需求窗口内年缺失产量矿山数（最小）", int(na_per_year.min()), "缺失不参与权重，不插值")
    add("需求窗口内年缺失产量矿山数（最大）", int(na_per_year.max()), "")
    add("需求窗口内年零产量矿山数（最小）", int(zero_per_year.min()), "真实零，明确排除")
    add("需求窗口内年零产量矿山数（最大）", int(zero_per_year.max()), "")

    active = long_all[long_all["P"] > 0]
    add("需求窗口内 active 矿山-年记录", int(len(active)), "")
    add("其中参与分配的记录（路线=火法/湿法）", int((active["工艺路线"].isin([ROUTE_FIRE, ROUTE_HYDRO])).sum()),
        "路线未知的记录不参与分配")
    add("2021 年矿山铜产量合计_tCu", round(float(long_all.loc[long_all['年份'] == 2021, 'P'].sum()), 3), "")
    add("2030 年矿山铜产量合计_tCu", round(float(long_all.loc[long_all['年份'] == 2030, 'P'].sum()), 3), "")
    add("2050 年矿山铜产量合计_tCu", round(float(long_all.loc[long_all['年份'] == 2050, 'P'].sum()), 3), "")

    # 运营期外产量占比
    early = active["投产年"].notna() & (active["年份"] < active["投产年"])
    late = active["退役年"].notna() & (active["年份"] > active["退役年"])
    tot = active["P"].sum()
    add("早于投产年的产量占比_%", round(float(active.loc[early, "P"].sum() / tot * 100), 4), "只标记不筛选")
    add("晚于退役年的产量占比_%", round(float(active.loc[late, "P"].sum() / tot * 100), 4), "只标记不筛选")
    add("投产/退役年是否用于硬筛", "否" if not USE_OPERATING_WINDOW else "是",
        "源表年份缺失多且与产量冲突量小，按产量准入")

    # 未知路线产量占比
    unk = active[active["工艺路线"] == ROUTE_UNKNOWN]["P"].sum()
    add("未知路线产量占比_%", round(float(unk / tot * 100), 4), "不进入任何路线分母")

    # 需求面板
    add("需求情景组合数", int(demand[[F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE]].drop_duplicates().shape[0]),
        "3 SSP × 3 回收 × 2 贸易")
    add("需求年份数", int(demand[F_DEM_YEAR].nunique()), f"{demand[F_DEM_YEAR].min()}–{demand[F_DEM_YEAR].max()}")
    add("需求组数（情景×年份×路线）", int(len(demand)), "")
    add("火法需求最小_tCu", round(float(demand[F_DEM_FIRE].min()), 3), "")
    add("火法需求最大_tCu", round(float(demand[F_DEM_FIRE].max()), 3), "")
    add("湿法需求最小_tCu", round(float(demand[F_DEM_HYDRO].min()), 3), "")
    add("湿法需求最大_tCu", round(float(demand[F_DEM_HYDRO].max()), 3), "")
    return q


# %% Chunk 4：年度矿山池和权重
# 功能：按 (年份, 路线) 汇总可参与矿山产量，计算组内份额，检查分母为零的组。
# 输入：长表 base。
# 输出：带权重与组内产量合计的 base；分组诊断信息。


def build_pools(base: pd.DataFrame) -> pd.DataFrame:
    """返回带权重、组内产量合计与分组规模的年度矿山池。"""
    base = base.copy()
    base["P_sum"] = base.groupby(["年份", "工艺路线"])[COL_P].transform("sum")
    base["路线产量份额"] = base[COL_P] / base["P_sum"]
    base["n_pool"] = base.groupby(["年份", "工艺路线"])["矿山ID"].transform("size")
    return base


# %% Chunk 5：铜需求分配
# 功能：读取需求面板，按 情景×年份×路线 把 D 乘以组内权重得到 A。
# 输入：base（含权重）、需求面板（含 火法/湿法 铜需求与铜精矿实物量）。
# 输出：每个情景一张明细块（含组级列），以及缺口行。


def load_demand() -> tuple[pd.DataFrame, list[dict]]:
    """读取需求面板并转长表；校验需求内部守恒与精矿实物量口径。"""
    dem = read_excel_robust(DEMAND_XLSX, SHEET_DEMAND)
    summ = read_excel_robust(DEMAND_XLSX, SHEET_SUMMARY)
    need = [F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE, F_DEM_YEAR, F_DEM_TOTAL, F_DEM_FIRE, F_DEM_HYDRO]
    missing = [c for c in need if c not in dem.columns]
    if missing:
        raise ValueError(f"路线分配_面板 缺少字段：{missing}")
    if F_DEM_CONC not in summ.columns:
        raise ValueError(f"年度情景汇总 缺少字段：{F_DEM_CONC}")

    dem = dem.copy()
    dem[F_DEM_YEAR] = dem[F_DEM_YEAR].astype(int)
    for c in [F_DEM_TOTAL, F_DEM_FIRE, F_DEM_HYDRO]:
        dem[c] = pd.to_numeric(dem[c], errors="coerce")
    if dem[[F_DEM_TOTAL, F_DEM_FIRE, F_DEM_HYDRO]].isna().any().any():
        raise ValueError("需求面板存在空值或非数值，请先核对上游结果。")

    # 需求内部闭合：火法 + 湿法 = 原生铜
    chk = (dem[F_DEM_FIRE] + dem[F_DEM_HYDRO] - dem[F_DEM_TOTAL]).abs()
    rel = float((chk / dem[F_DEM_TOTAL].clip(lower=1e-12)).max())
    if rel > 1e-8:
        raise ValueError(f"需求面板 火法+湿法≠原生铜，最大相对误差 {rel:.3e}")

    # 每个组合必须覆盖 2021–2050 全部年份
    years_expect = set(range(YEAR_MIN, YEAR_MAX + 1))
    for key, sub in dem.groupby([F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE]):
        got = set(sub[F_DEM_YEAR].astype(int))
        if got != years_expect:
            raise ValueError(f"情景组合 {key} 的年份不完整：缺 {sorted(years_expect - got)}")

    keys = [F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE, F_DEM_YEAR]
    summ2 = summ[keys + [F_DEM_CONC]].copy()
    summ2[F_DEM_CONC] = pd.to_numeric(summ2[F_DEM_CONC], errors="coerce")
    if summ2.duplicated(keys).any():
        raise ValueError("年度情景汇总 存在重复的 情景×年份 键")
    if (summ2[F_DEM_CONC] <= 0).any():
        raise ValueError("铜精矿实物量存在非正值，无法换算火法产物量。")

    dem = dem.merge(summ2, on=keys, how="left", validate="one_to_one")
    if dem[F_DEM_CONC].isna().any():
        raise ValueError("部分需求组缺少铜精矿实物量。")

    # 隐含品位 = 火法铜 / 铜精矿实物量，必须在 (0,1)
    grade = dem[F_DEM_FIRE] / dem[F_DEM_CONC]
    if ((grade <= 0) | (grade >= 1)).any():
        raise ValueError("火法铜/铜精矿实物量 不在 (0,1) 区间，精矿口径异常。")

    fire = dem[keys + [F_DEM_TOTAL, F_DEM_FIRE, F_DEM_CONC]].copy()
    fire["工艺路线"] = ROUTE_FIRE
    fire["D"] = fire[F_DEM_FIRE]
    hydro = dem[keys + [F_DEM_TOTAL, F_DEM_HYDRO, F_DEM_CONC]].copy()
    hydro["工艺路线"] = ROUTE_HYDRO
    hydro["D"] = hydro[F_DEM_HYDRO]
    long = pd.concat([fire, hydro], ignore_index=True)
    long["情景ID"] = (
        long[F_DEM_SCEN].astype(str) + "|" + long[F_DEM_REC].astype(str) + "|" + long[F_DEM_TRADE].astype(str)
    )
    long = long.rename(columns={
        F_DEM_CONC: COL_CONC,
        "D": COL_D,
    })

    # 情景排序
    long["_o1"] = long[F_DEM_SCEN].map(SCEN_ORDER).fillna(99)
    long["_o2"] = long[F_DEM_REC].map(REC_ORDER).fillna(99)
    long["_o3"] = long[F_DEM_TRADE].map(TRADE_ORDER).fillna(99)
    long = long.sort_values(["_o1", "_o2", "_o3", F_DEM_YEAR, "工艺路线"],
                            key=lambda s: s.map(ROUTE_ORDER) if s.name == "工艺路线" else s)
    long = long.drop(columns=["_o1", "_o2", "_o3"]).reset_index(drop=True)

    info = [{
        "项目": "需求组隐含精矿含铜率（最小/最大）",
        "数值": f"{grade.min():.6f} / {grade.max():.6f}",
        "说明": "= 火法铜_tCu / 铜精矿实物量_t，直接沿用既有面板",
    }]
    return long, info


def allocate_scenario(base: pd.DataFrame, dem_one: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """对一个情景做分配：A = D × w；返回明细块与组级校验表。"""
    d = dem_one[[F_DEM_YEAR, "工艺路线", COL_D, COL_CONC]].copy()
    block = base.merge(
        d, on=[F_DEM_YEAR, "工艺路线"], how="inner", validate="many_to_one",
        suffixes=("", "_dem"),
    )
    block[COL_A] = block[COL_D] * block["路线产量份额"]

    # 组级合计与未分配
    gsum = block.groupby(["年份", "工艺路线"])[COL_A].transform("sum")
    block[COL_UNALLOC] = block[COL_D] - gsum
    block[COL_STATUS] = np.where(
        block[COL_D] <= 0, "需求为零",
        np.where(block["路线产量份额"].isna(), "无有效矿山（缺口）", "正常分配"),
    )

    # 产出列
    block[COL_CAP] = block["参考产能_tCu每年"]
    ratio = np.where(block[COL_CAP] > 0, block[COL_A] / block[COL_CAP], np.nan)
    block[COL_RATIO] = ratio
    block[COL_SOURCE] = "Mine Production!R" + block["源行号"].astype(str)
    block[COL_NOTE] = "标签=" + block["产品标签"].fillna("").astype(str)

    # 缺口组：需求存在但年度矿山池为空
    present = set(block[[F_DEM_YEAR, "工艺路线"]].drop_duplicates().itertuples(index=False, name=None))
    gaps = d.drop_duplicates([F_DEM_YEAR, "工艺路线"]).copy()
    keep = [tuple(x) not in present for x in gaps[[F_DEM_YEAR, "工艺路线"]].itertuples(index=False, name=None)]
    gaps = gaps[keep]
    gap_cols = [COL_SCEN_ID, F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE, F_DEM_YEAR, COL_MINE_ID,
                COL_NAME, COL_COUNTRY, COL_REGION, COL_ROUTE, COL_W, COL_P, COL_A, COL_CAP,
                COL_RATIO, COL_D, COL_PSUM, COL_UNALLOC, COL_CONC, COL_STATUS, COL_FLAGS,
                COL_SOURCE, COL_NOTE, "工艺路线序"]
    if gaps.empty:
        gap_rows = pd.DataFrame(columns=gap_cols)
    else:
        gap_rows = pd.DataFrame({
            COL_SCEN_ID: dem_one[COL_SCEN_ID].iloc[0],
            F_DEM_SCEN: dem_one[F_DEM_SCEN].iloc[0],
            F_DEM_REC: dem_one[F_DEM_REC].iloc[0],
            F_DEM_TRADE: dem_one[F_DEM_TRADE].iloc[0],
            F_DEM_YEAR: gaps[F_DEM_YEAR].astype(int),
            COL_MINE_ID: "GAP000",
            COL_NAME: "无可分配矿山",
            COL_COUNTRY: "",
            COL_REGION: "",
            COL_ROUTE: gaps["工艺路线"],
            COL_W: np.nan,
            COL_P: 0.0,
            COL_A: gaps[COL_D].to_numpy(),
            COL_CAP: np.nan,
            COL_RATIO: np.nan,
            COL_D: gaps[COL_D].to_numpy(),
            COL_PSUM: 0.0,
            COL_UNALLOC: gaps[COL_D].to_numpy(),
            COL_CONC: gaps[COL_CONC].to_numpy(),
            COL_STATUS: "无有效矿山（缺口）",
            COL_FLAGS: "年度矿山池为空",
            COL_SOURCE: "",
            COL_NOTE: "缺口记录：需求无法分配",
            "工艺路线序": gaps["工艺路线"].map(ROUTE_ORDER).to_numpy(),
        })

    block[COL_PSUM] = block["P_sum"]
    block["工艺路线序"] = block["工艺路线"].map(ROUTE_ORDER)
    block[F_DEM_SCEN] = dem_one[F_DEM_SCEN].iloc[0]
    block[F_DEM_REC] = dem_one[F_DEM_REC].iloc[0]
    block[F_DEM_TRADE] = dem_one[F_DEM_TRADE].iloc[0]
    block[COL_SCEN_ID] = dem_one[COL_SCEN_ID].iloc[0]

    # 组级校验
    grp = block.groupby([F_DEM_YEAR, "工艺路线"], as_index=False).agg(
        sumA=(COL_A, "sum"),
        sumP=(COL_P, "sum"),
        n=(COL_MINE_ID, "size"),
        sumw=("路线产量份额", "sum"),
        minA=(COL_A, "min"),
        minw=("路线产量份额", "min"),
        maxw=("路线产量份额", "max"),
        sumD=(COL_D, "max"),
    )
    grp[COL_SCEN_ID] = dem_one[COL_SCEN_ID].iloc[0]
    grp["未分配"] = grp["sumD"] - grp["sumA"]
    grp["相对误差"] = grp["未分配"].abs() / grp["sumD"].clip(lower=1e-12)

    detail = pd.concat([block, gap_rows], ignore_index=True, sort=False)
    return detail, grp


# %% Chunk 6：校验
# 功能：逐组检查守恒、权重、非负、主键、路线匹配、年份覆盖与重复记录。
# 输入：各情景明细块与组级校验表。
# 输出：错误列表（为空即通过）。


def validate_all(block: pd.DataFrame, grp: pd.DataFrame, sid: str, errors: list[str]) -> None:
    """对单个情景块执行硬校验，错误追加到 errors。"""
    # 主键唯一
    key_dup = block.duplicated([COL_SCEN_ID, F_DEM_YEAR, COL_MINE_ID, "工艺路线"]).sum()
    if key_dup:
        errors.append(f"[{sid}] 主键重复 {int(key_dup)} 条")

    # 组守恒
    bad = grp[grp["相对误差"] > TOL_REL]
    if len(bad):
        errors.append(
            f"[{sid}] 守恒失败 {len(bad)} 组，最大相对误差 {bad['相对误差'].max():.3e}"
        )

    # 权重
    real = grp[grp["sumD"] > 0]
    badw = real[(real["sumw"] - 1).abs() > TOL_REL]
    if len(badw):
        errors.append(f"[{sid}] 组内权重和≠1 共 {len(badw)} 组")

    # 非负与值域
    if (block[COL_A] < 0).any():
        errors.append(f"[{sid}] 存在负的分配量")
    w = block["路线产量份额"].dropna()
    if (w < 0).any() or (w > 1 + 1e-12).any():
        errors.append(f"[{sid}] 权重越界 [0,1]")

    # 路线匹配
    bad_route = block[~block["工艺路线"].isin([ROUTE_FIRE, ROUTE_HYDRO])]
    if len(bad_route):
        errors.append(f"[{sid}] 存在非火法/湿法路线 {len(bad_route)} 条")

    # 年份覆盖
    if set(block[F_DEM_YEAR].unique()) != set(range(YEAR_MIN, YEAR_MAX + 1)):
        errors.append(f"[{sid}] 年份覆盖不完整")


# %% Chunk 7：保存中间结果和 Excel
# 功能：流式写 CSV 与最终 Excel 第一张表（XlsxWriter constant_memory，控制内存）。
# 输入：各情景明细块、组级校验、数据质量记录。
# 输出：01_铜需求分配.csv、01_组级校验.csv、01_数据质量.csv、最终 Excel（表 1 + 占位表 2）。


def make_formats(wb: xlsxwriter.Workbook) -> dict:
    """定义表头与数字格式。"""
    fmts = {
        "header": wb.add_format({
            "bold": True, "bg_color": "#DDEBF7", "border": 1,
            "text_wrap": True, "valign": "vcenter", "align": "center",
        }),
        "year": wb.add_format({"num_format": "0", "align": "center"}),
        "w": wb.add_format({"num_format": "0.000000"}),
        "tcu": wb.add_format({"num_format": "#,##0.000"}),
        "tcu_group": wb.add_format({"num_format": "#,##0.000"}),
        "ratio": wb.add_format({"num_format": "0.000000"}),
        "m3": wb.add_format({"num_format": "#,##0.000"}),
        "mm3": wb.add_format({"num_format": "0.000000"}),
        "text": None,
    }
    return fmts


class StreamSheet:
    """按行顺序写工作表（constant_memory 模式），支持冻结首行、筛选、列宽与列格式。"""

    def __init__(self, wb, name, columns, formats, total_rows):
        self.ws = wb.add_worksheet(name)
        self.headers = [c[0] for c in columns]
        self.fmt_keys = [c[2] for c in columns]
        self.formats = formats
        self.ws.freeze_panes(1, 0)
        for j, (_, width, _) in enumerate(columns):
            self.ws.set_column(j, j, width)
        self.ws.write_row(0, 0, self.headers, formats["header"])
        if total_rows:
            self.ws.autofilter(0, 0, int(total_rows), len(columns) - 1)
        self.row = 0

    def write_values(self, values) -> None:
        r = self.row + 1
        for j, v in enumerate(values):
            if v is None or v is pd.NA:
                continue
            if isinstance(v, np.generic):
                v = v.item()
            if isinstance(v, float) and math.isnan(v):
                continue
            fmt = self.formats.get(self.fmt_keys[j])
            if isinstance(v, str):
                self.ws.write_string(r, j, v[:32000], None)
            else:
                self.ws.write_number(r, j, v, fmt)
        self.row = r


def frame_to_rows(df: pd.DataFrame, columns: list[tuple[str, int, str]]):
    """按列定义顺序把 DataFrame 转成可写行。"""
    view = df[[c[0] for c in columns]]
    for rec in view.itertuples(index=False, name=None):
        yield rec


def write_excel(blocks, total_rows_sheet1: int) -> None:
    """生成最终 Excel 的第一张表（第二张表由程序二覆盖生成）。"""
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_XLSX), {"constant_memory": True})
    fmts = make_formats(wb)
    sh1 = StreamSheet(wb, "铜需求分配", SHEET1_COLUMNS, fmts, total_rows_sheet1)
    # 占位表 2：程序二会用完整两张表覆盖本工作簿
    ws2 = wb.add_worksheet("用水核算")
    ws2.write_row(0, 0, ["占位说明"], fmts["header"])
    ws2.write_string(1, 0, "本工作表由 02_calculate_water.py 生成并覆盖；请先运行程序二。")
    ws2.set_column(0, 0, 80)

    for block in blocks:
        for rec in frame_to_rows(block, SHEET1_COLUMNS):
            sh1.write_values(rec)
    if sh1.row != total_rows_sheet1:
        raise RuntimeError(f"表 1 行数不符：写入 {sh1.row}，预期 {total_rows_sheet1}")
    wb.close()


# %% Chunk 8：main()
# 功能：串联读取 → 检查 → 权重 → 分配 → 校验 → 保存，并打印摘要。
# 输入：无（路径来自 Chunk 1）。
# 输出：中间 CSV、最终 Excel 第一张表、控制台摘要。


def main() -> None:
    print("=" * 78)
    print("程序一：逐年 情景×路线 铜需求分配")
    print("=" * 78)

    for p in (INPUT_XLSX, DEMAND_XLSX):
        if not p.exists():
            raise FileNotFoundError(f"输入文件不存在：{p}")
    INTERMEDIATE_DIR.mkdir(parents=True, exist_ok=True)

    hash_in = sha256_file(INPUT_XLSX)
    hash_dem = sha256_file(DEMAND_XLSX)
    print(f"输入 数据.xlsx sha256 = {hash_in}")
    print(f"输入 需求面板 sha256 = {hash_dem}")

    # Chunk 2
    mines, year_cols = load_mines()
    print(f"矿山记录 {len(mines)}，年份列 {len(year_cols)}（{year_cols[0][1]}–{year_cols[-1][1]}）")
    base = build_long(mines, year_cols)
    base = build_pools(base)
    print(f"需求窗口内参与分配的矿山-年记录：{len(base)}")

    # Chunk 3
    demand_long, demand_info = load_demand()
    quality = collect_quality(mines, year_cols, demand_long)
    quality.extend(demand_info)

    # 情景清单
    scen = demand_long[[COL_SCEN_ID, F_DEM_SCEN, F_DEM_REC, F_DEM_TRADE]].drop_duplicates().reset_index(drop=True)
    print(f"情景组合 {len(scen)}，需求组 {len(demand_long)}")

    # 缺口组预计算（用于表 1 行数）
    base_pairs = set(base[[F_DEM_YEAR, "工艺路线"]].drop_duplicates().itertuples(index=False, name=None))
    dem_pairs = set(demand_long[[F_DEM_YEAR, "工艺路线"]].drop_duplicates().itertuples(index=False, name=None))
    gap_pairs = sorted(dem_pairs - base_pairs, key=lambda x: (x[0], ROUTE_ORDER.get(x[1], 9)))
    n_gap_total = len(gap_pairs) * len(scen)
    total_rows_sheet1 = len(base) * len(scen) + n_gap_total
    print(f"预期表 1 明细行数 = {len(base)} × {len(scen)} + 缺口 {n_gap_total} = {total_rows_sheet1}")

    # Chunk 5 + 6 + 7：逐情景分配、校验、写 CSV
    blocks = []
    grp_tables = []
    errors: list[str] = []
    first = True
    for _, srow in scen.iterrows():
        sid = srow[COL_SCEN_ID]
        one = demand_long[demand_long[COL_SCEN_ID] == sid]
        block, grp = allocate_scenario(base, one)
        validate_all(block, grp, sid, errors)
        if len(block) != len(base) + len(gap_pairs):
            errors.append(f"[{sid}] 明细行数异常：{len(block)} vs {len(base) + len(gap_pairs)}")

        block = block.sort_values(
            [COL_MINE_ID, "工艺路线序", F_DEM_YEAR]
        ).reset_index(drop=True)

        out = block[CSV_HEADERS]
        out.to_csv(ALLOC_CSV, mode="w" if first else "a", header=first,
                   index=False, encoding="utf-8-sig")
        first = False
        blocks.append(block)
        grp_tables.append(grp)
        print(f"  {sid}: 明细 {len(block)} 行，最大相对误差 {grp['相对误差'].max():.3e}")

    if errors:
        print("\n校验失败：")
        for e in errors[:20]:
            print("  -", e)
        raise RuntimeError(f"校验未通过，共 {len(errors)} 项；未生成正式结果。")

    # 汇总校验
    grp_all = pd.concat(grp_tables, ignore_index=True)
    grp_all = grp_all.rename(columns={
        F_DEM_YEAR: "年份", "工艺路线": "工艺路线", "sumD": "需求D_tCu", "sumA": "分配合计_tCu",
        "sumP": "矿山产量合计_tCu", "n": "矿山数", "sumw": "权重之和",
        "minw": "权重最小值", "maxw": "权重最大值", "minA": "分配量最小值",
        "未分配": "未分配_tCu", "相对误差": "相对误差", COL_SCEN_ID: "情景ID",
    })
    grp_all = grp_all[["情景ID", "年份", "工艺路线", "需求D_tCu", "分配合计_tCu", "未分配_tCu",
                       "相对误差", "矿山产量合计_tCu", "矿山数", "权重之和", "权重最小值", "权重最大值", "分配量最小值"]]
    grp_all.to_csv(GROUP_CHECK_CSV, index=False, encoding="utf-8-sig")

    # 校验结果并入数据质量
    quality.extend([
        {"项目": "组级守恒检查组数", "数值": int(len(grp_all)), "说明": "18 情景 × 30 年 × 2 路线"},
        {"项目": "组级守恒最大相对误差", "数值": f"{grp_all['相对误差'].max():.3e}", "说明": f"容差 {TOL_REL:g}"},
        {"项目": "权重之和偏离 1 的最大值", "数值": f"{(grp_all['权重之和'] - 1).abs().max():.3e}", "说明": ""},
        {"项目": "最小分配量_tCu", "数值": f"{grp_all['分配量最小值'].min():.6f}", "说明": "组内最小分配量，仅作非负检查"},
        {"项目": "表 1 明细行数", "数值": total_rows_sheet1, "说明": "矿山-年 × 18 情景 + 缺口"},
        {"项目": "输出 CSV 校验行数", "数值": sum(len(b) for b in blocks), "说明": ""},
        {"项目": "输入 数据.xlsx sha256（运行前）", "数值": hash_in, "说明": ""},
        {"项目": "输入 需求面板 sha256（运行前）", "数值": hash_dem, "说明": ""},
    ])
    pd.DataFrame(quality).to_csv(QUALITY_CSV, index=False, encoding="utf-8-sig")

    # 表 1 Excel
    write_excel(blocks, total_rows_sheet1)

    # 输入未被改动
    if sha256_file(INPUT_XLSX) != hash_in:
        raise RuntimeError("输入 数据.xlsx 在运行期间被修改，请复核。")

    print("\n完成：")
    print(f"  中间结果 {ALLOC_CSV}")
    print(f"  组级校验 {GROUP_CHECK_CSV}")
    print(f"  数据质量 {QUALITY_CSV}")
    print(f"  最终 Excel {OUTPUT_XLSX}（表 1 完成，表 2 待程序二）")
    print(f"  组守恒最大相对误差 {grp_all['相对误差'].max():.3e}")


if __name__ == "__main__":
    main()
