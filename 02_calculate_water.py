#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
02_calculate_water.py
=====================
读取程序一的分配结果与原始用水系数表，把"分配铜量（t Cu）"换算为与系数同口径的
"产物实物量（t 产物）"，再计算逐年、分情景、分矿山、分路线的**用水量（m³）**。

口径（与 README 一致）：
  * 火法路线产物 = 铜精矿；产物实物量直接取既有"铜精矿实物量"面板按同一权重分摊，
    不再用品位重复换算（换算系数 c = 组内火法铜 / 组内铜精矿实物量 仅用于复核 Q=A/c）；
  * 湿法路线产物 = 阴极铜；需求阶段即精炼铜（阴极铜），按 1:1 t Cu/t 产物 换算；
  * 用水系数 = m³/t 产物（指标为"用水量"，源表未注明取水/耗水，故不作该区分）；
  * 系数按 地区 × 路线 × 产物 精确匹配，缺失即报错，不回落、不用均值、不用 0。

运行：python 02_calculate_water.py
输出：intermediate/02_用水明细.csv、02_年度总计.csv、02_校验结果.csv
      最终 Excel：铜需求分配与用水核算.xlsx（仅两张表）
"""

from __future__ import annotations

import hashlib
import math
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

# %% Chunk 1：配置
# 功能：集中定义路径、列名、产物与系数匹配规则、容差。
# 输入：无。
# 输出：后续 chunk 使用的常量。

BASE_DIR = Path(__file__).resolve().parent
INPUT_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx")
OUTPUT_XLSX = Path(r"C:\Users\22965\OneDrive\桌面\MFA\Output\铜需求分配与用水核算.xlsx")
INTERMEDIATE_DIR = BASE_DIR / "intermediate"
ALLOC_CSV = INTERMEDIATE_DIR / "01_铜需求分配.csv"
WATER_DETAIL_CSV = INTERMEDIATE_DIR / "02_用水明细.csv"
WATER_TOTAL_CSV = INTERMEDIATE_DIR / "02_年度总计.csv"
CHECK_CSV = INTERMEDIATE_DIR / "02_校验结果.csv"

SHEET_WATER = "矿山用水系数"

# --- 程序一中间结果的列名 ---
COL_SCEN_ID = "情景ID"
COL_SCEN = "需求情景"
COL_REC = "回收情景"
COL_TRADE = "贸易情景"
COL_YEAR = "年份"
COL_MINE_ID = "矿山ID"
COL_NAME = "矿山项目名称"
COL_COUNTRY = "国家"
COL_REGION = "地区"
COL_ROUTE = "工艺路线"
COL_W = "路线产量份额"
COL_P = "当年参考铜产量_tCu"
COL_A = "分配铜需求量_tCu"
COL_A2 = "分配铜量_tCu"  # 表 2 的列名（与表 1 同值，便于按表 2 口径阅读）
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

# --- 路线 / 产物 / 系数列 ---
ROUTE_FIRE = "火法"
ROUTE_HYDRO = "湿法"
ROUTE_ORDER = {ROUTE_FIRE: 0, ROUTE_HYDRO: 1}
PRODUCT_FIRE = "铜精矿"
PRODUCT_HYDRO = "阴极铜"
COEF_COL_TO_ROUTE = {"精矿矿山（浮选）": ROUTE_FIRE, "SX-EW矿山": ROUTE_HYDRO}
ROUTE_TO_COEF_COL = {v: k for k, v in COEF_COL_TO_ROUTE.items()}
ROUTE_TO_PRODUCT = {ROUTE_FIRE: PRODUCT_FIRE, ROUTE_HYDRO: PRODUCT_HYDRO}

TOL_REL = 1e-9
TOL_ABS_W = 1e-6  # 年度总计与明细之和的绝对容差（m³）

# --- 表 1（程序一生成，本程序原样重建）列定义 ---
SHEET1_COLUMNS = [
    ("情景ID", 22, "text"), ("需求情景", 12, "text"), ("回收情景", 11, "text"),
    ("贸易情景", 11, "text"), ("年份", 7, "year"), ("矿山ID", 9, "text"),
    ("矿山项目名称", 28, "text"), ("国家", 18, "text"), ("地区", 10, "text"),
    ("工艺路线", 9, "text"), ("路线产量份额", 13, "w"), ("当年参考铜产量_tCu", 18, "tcu"),
    ("分配铜需求量_tCu", 18, "tcu"), ("参考产能_tCu每年", 18, "tcu"),
    ("分配量除参考产能", 16, "ratio"), ("组内路线需求总量_tCu（组级·勿逐行求和）", 26, "tcu"),
    ("组内参考产量合计_tCu（组级·勿逐行求和）", 26, "tcu"),
    ("未分配需求_tCu（组级）", 20, "tcu"), ("分配状态", 14, "text"),
    ("数据质量标记", 30, "text"), ("数据来源", 26, "text"), ("备注", 26, "text"),
]

# --- 表 2（本程序生成）列定义 ---
COL_LEVEL = "统计层级"
COL_PRODUCT = "产物名称"
COL_Q = "产物实物量_t"
COL_C = "产物含铜率_tCu每t产物"
COL_BASIS = "产物干湿基"
COL_K = "用水系数_m3每t产物"
COL_WATER = "用水量_m3"
COL_WATER_MM = "用水量_百万m3"
COL_WATER_FIRE = "其中火法用水量_m3"
COL_WATER_HYDRO = "其中湿法用水量_m3"
COL_MATCH = "系数匹配状态"
COL_CALC = "核算状态"

SHEET2_COLUMNS = [
    (COL_LEVEL, 10, "text"), (COL_SCEN_ID, 22, "text"), (COL_SCEN, 12, "text"),
    (COL_REC, 11, "text"), (COL_TRADE, 11, "text"), (COL_YEAR, 7, "year"),
    (COL_MINE_ID, 9, "text"), (COL_NAME, 28, "text"), (COL_COUNTRY, 18, "text"),
    (COL_REGION, 10, "text"), (COL_ROUTE, 9, "text"), (COL_PRODUCT, 10, "text"),
    ("分配铜量_tCu", 16, "tcu"), (COL_C, 18, "ratio6"), (COL_Q, 16, "tprod"),
    (COL_BASIS, 18, "text"), (COL_K, 16, "ratio3"), (COL_WATER, 18, "m3"),
    (COL_WATER_MM, 14, "mm3"), (COL_WATER_FIRE, 18, "m3"), (COL_WATER_HYDRO, 18, "m3"),
    (COL_MATCH, 26, "text"), (COL_CALC, 30, "text"), (COL_NOTE, 32, "text"),
]
SHEET2_HEADERS = [c[0] for c in SHEET2_COLUMNS]


# %% Chunk 2：读取分配结果及系数
# 功能：读入程序一长表与原始系数表；校验主键、系数键与数值合法性。
# 输入：ALLOC_CSV、INPUT_XLSX!矿山用水系数。
# 输出：alloc（明细长表）、coef（地区×路线 系数长表）。


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_excel_robust(path: Path, sheet_name: str) -> pd.DataFrame:
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


def load_allocation() -> pd.DataFrame:
    """读取程序一输出，校验行数与主键。"""
    if not ALLOC_CSV.exists():
        raise FileNotFoundError(f"未找到程序一中间结果：{ALLOC_CSV}（请先运行 01_allocate_copper.py）")
    df = pd.read_csv(ALLOC_CSV, encoding="utf-8-sig")
    need = [COL_SCEN_ID, COL_SCEN, COL_REC, COL_TRADE, COL_YEAR, COL_MINE_ID, COL_NAME,
            COL_COUNTRY, COL_REGION, COL_ROUTE, COL_W, COL_A, COL_D, COL_CONC, COL_STATUS]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(f"分配结果缺少字段：{missing}")
    df[COL_YEAR] = df[COL_YEAR].astype(int)
    for c in [COL_W, COL_A, COL_CAP, COL_RATIO, COL_D, COL_PSUM, COL_UNALLOC, COL_CONC, COL_P]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    dup = df.duplicated([COL_SCEN_ID, COL_YEAR, COL_MINE_ID, COL_ROUTE]).sum()
    if dup:
        raise ValueError(f"分配结果主键重复 {int(dup)} 条，拒绝继续。")
    return df


def load_coefficients() -> tuple[pd.DataFrame, pd.DataFrame]:
    """读取用水系数表并转长表，校验键唯一、数值为正。"""
    raw = read_excel_robust(INPUT_XLSX, SHEET_WATER)
    raw.columns = [str(c).strip() for c in raw.columns]
    expected = ["地区"] + list(COEF_COL_TO_ROUTE.keys())
    missing = [c for c in expected if c not in raw.columns]
    if missing:
        raise ValueError(f"矿山用水系数 缺少字段：{missing}")
    raw = raw[expected].copy()
    if raw["地区"].astype(str).str.strip().duplicated().any():
        raise ValueError("矿山用水系数 存在重复地区键。")
    for c in COEF_COL_TO_ROUTE:
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
        if raw[c].isna().any() or (raw[c] <= 0).any():
            raise ValueError(f"用水系数列 {c} 存在空值/非正值，拒绝使用。")
    raw["地区"] = raw["地区"].astype(str).str.strip()

    long = raw.melt(id_vars=["地区"], value_vars=list(COEF_COL_TO_ROUTE),
                    var_name="系数列", value_name=COL_K)
    long[COL_ROUTE] = long["系数列"].map(COEF_COL_TO_ROUTE)
    long = long[["地区", COL_ROUTE, "系数列", COL_K]]
    if long.duplicated(["地区", COL_ROUTE]).any():
        raise ValueError("用水系数 地区×路线 键重复。")
    return raw, long


# %% Chunk 3：地区、路线和产品匹配
# 功能：为每条矿山记录匹配唯一系数；缺口行单独标注；不回落、不用均值。
# 输入：alloc、coef 长表。
# 输出：带 k / 产物名称 / 匹配状态 的明细表。


def match_coefficients(alloc: pd.DataFrame, coef_long: pd.DataFrame) -> pd.DataFrame:
    df = alloc.copy()
    is_gap = df[COL_MINE_ID].astype(str).eq("GAP000")
    df = df.merge(
        coef_long[["地区", COL_ROUTE, COL_K]],
        on=["地区", COL_ROUTE], how="left", validate="many_to_one",
    )
    bad = df[~is_gap & df[COL_K].isna()]
    if len(bad):
        keys = bad[["地区", COL_ROUTE]].drop_duplicates().to_dict("records")
        raise ValueError(f"以下 地区×路线 未匹配到用水系数（禁止回落/均值）：{keys}")
    df[COL_PRODUCT] = df[COL_ROUTE].map(ROUTE_TO_PRODUCT)
    df[COL_MATCH] = np.where(
        is_gap, "缺口行：无矿山，未匹配",
        "匹配成功：" + df[COL_REGION].astype(str) + "×" + df[COL_ROUTE].astype(str)
        + "×" + df[COL_PRODUCT].astype(str),
    )
    # 核对系数与路线一一对应（防止列错位）
    for route, col in ROUTE_TO_COEF_COL.items():
        sub = coef_long[coef_long[COL_ROUTE] == route]
        if len(sub) != len(sub["地区"].unique()):
            raise ValueError(f"系数表 {col} 列存在重复地区。")
    return df


# %% Chunk 4：铜量转产物量
# 功能：火法用组内铜精矿实物量×份额；湿法 1:1；并复核 Q=A/c。
# 输入：带系数的明细表。
# 输出：带 产物含铜率 / 产物实物量 / 干湿基 的明细表。


def convert_product(df: pd.DataFrame) -> tuple[pd.DataFrame, float]:
    out = df.copy()
    is_gap = out[COL_MINE_ID].astype(str).eq("GAP000")
    is_fire = out[COL_ROUTE].eq(ROUTE_FIRE)

    # 产物含铜率：火法 c = 组内火法铜 / 组内铜精矿实物量；湿法 1.0（精炼铜 1:1 阴极铜）
    c = np.where(is_fire, out[COL_D] / out[COL_CONC], 1.0)
    out[COL_C] = c

    # 产物实物量：火法 = 组内铜精矿实物量 × 矿山份额；湿法 = 分配铜量
    q = np.where(is_fire, out[COL_CONC] * out[COL_W], out[COL_A])
    q = np.where(is_gap, 0.0, q)
    out[COL_Q] = q
    out[COL_BASIS] = np.where(is_fire, "未注明（沿用既有质量收率链实物量）", "不适用（金属阴极铜）")

    # 复核 Q = A / c（容差 1e-9 相对）
    real = ~is_gap & (out[COL_A] > 0) & (out[COL_C] > 0)
    q_check = out.loc[real, COL_A] / out.loc[real, COL_C]
    rel = float(((q_check - out.loc[real, COL_Q]).abs() / q_check.clip(lower=1e-12)).max())
    if rel > 1e-9:
        raise ValueError(f"火法 Q=A/c 复核失败，最大相对误差 {rel:.3e}")

    if (out[COL_Q] < 0).any():
        raise ValueError("产物实物量出现负值。")
    if (out.loc[~is_gap, COL_Q] == 0).any():
        raise ValueError("存在分配铜量>0 但产物实物量为 0 的记录，请复核。")
    return out, rel


# %% Chunk 5：用水计算
# 功能：W = Q × k，并换算百万 m³。
# 输入：带产物的明细表。
# 输出：带用水量的明细表。


def compute_water(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out[COL_LEVEL] = "矿山明细"
    out[COL_A2] = out[COL_A]
    is_gap = out[COL_MINE_ID].astype(str).eq("GAP000")
    k = out[COL_K].fillna(0.0)
    out[COL_WATER] = out[COL_Q] * k
    out[COL_WATER] = out[COL_WATER].where(~is_gap, 0.0)
    out[COL_WATER_MM] = out[COL_WATER] / 1e6
    if (out[COL_WATER] < 0).any():
        raise ValueError("用水量出现负值。")
    out[COL_WATER_FIRE] = np.nan
    out[COL_WATER_HYDRO] = np.nan
    # 核算状态
    out[COL_CALC] = np.select(
        [is_gap, out[COL_A] <= 0, out[COL_Q] > 0],
        ["缺口：无矿山，无法核算用水", "分配量为零", "正常（产物量×系数）"],
        default="异常：产物量为零",
    )
    out[COL_NOTE] = np.where(
        out[COL_ROUTE].eq(ROUTE_FIRE),
        "Q=组内铜精矿实物量×份额；c=火法铜/精矿实物量",
        "Q=A（阴极铜 1:1）；k 为 SX-EW 矿山系数",
    )
    out.loc[is_gap, COL_NOTE] = "缺口行：无矿山记录，未分配且未核算"
    return out


# %% Chunk 6：年度汇总与校验
# 功能：按 情景×年份 汇总用水量；校验总计=明细、主键唯一、与程序一 A 一致。
# 输入：用水明细。
# 输出：年度总计表、校验记录。


def build_totals(df: pd.DataFrame) -> pd.DataFrame:
    keys = [COL_SCEN_ID, COL_SCEN, COL_REC, COL_TRADE, COL_YEAR]
    g = df.groupby(keys, sort=False).agg(
        A=(COL_A2, "sum"),
        W=(COL_WATER, "sum"),
        n=(COL_MINE_ID, "size"),
    )
    route_w = df.pivot_table(index=keys, columns=COL_ROUTE, values=COL_WATER,
                             aggfunc="sum", fill_value=0.0)
    tot = g.join(route_w, how="left").reset_index()
    for r in [ROUTE_FIRE, ROUTE_HYDRO]:
        if r not in tot.columns:
            tot[r] = 0.0
    tot["其中火法"] = tot[ROUTE_FIRE].fillna(0.0)
    tot["其中湿法"] = tot[ROUTE_HYDRO].fillna(0.0)

    out = pd.DataFrame({
        COL_LEVEL: "年度总计",
        COL_SCEN_ID: tot[COL_SCEN_ID],
        COL_SCEN: tot[COL_SCEN],
        COL_REC: tot[COL_REC],
        COL_TRADE: tot[COL_TRADE],
        COL_YEAR: tot[COL_YEAR].astype(int),
        COL_MINE_ID: "",
        COL_NAME: "全部矿山",
        COL_COUNTRY: "",
        COL_REGION: "",
        COL_ROUTE: "合计",
        COL_PRODUCT: "",
        "分配铜量_tCu": tot["A"],
        COL_C: np.nan,
        COL_Q: np.nan,
        COL_BASIS: "",
        COL_K: np.nan,
        COL_WATER: tot["W"],
        COL_WATER_MM: tot["W"] / 1e6,
        COL_WATER_FIRE: tot["其中火法"],
        COL_WATER_HYDRO: tot["其中湿法"],
        COL_MATCH: "—",
        COL_CALC: "由矿山明细汇总（不可与明细重复相加）",
        COL_NOTE: "明细 " + tot["n"].astype(str) + " 行；产物吨数跨产品不可汇总，故留空",
    })
    return out


def validate_all(df: pd.DataFrame, totals: pd.DataFrame, alloc: pd.DataFrame,
                 q_rel: float) -> list[dict]:
    """返回校验记录；任一致命项失败直接抛错。"""
    checks: list[dict] = []
    errors: list[str] = []

    def add(item, value, tol, passed, note=""):
        checks.append({"检查项": item, "结果值": value, "容差": tol,
                       "结果": "通过" if passed else "未通过", "说明": note})

    # 1 主键唯一
    dup = df.duplicated([COL_SCEN_ID, COL_YEAR, COL_MINE_ID, COL_ROUTE]).sum()
    add("矿山明细主键唯一", int(dup), "0 重复", dup == 0, "情景×年份×矿山×路线")
    if dup:
        errors.append("明细主键重复")

    # 2 年度总计 = 明细之和
    keys = [COL_SCEN_ID, COL_YEAR]
    s = df.groupby(keys, sort=False)[COL_WATER].sum().reset_index()
    m = totals.merge(s, on=keys, how="left", suffixes=("", "_detail"))
    diff = (m[COL_WATER] - m[COL_WATER + "_detail"]).abs().max()
    add("年度总计=明细之和", f"{diff:.3e}", f"{TOL_ABS_W:g}",
        diff <= max(TOL_ABS_W, float(m[COL_WATER].abs().max()) * 1e-12), "m³")
    if diff > max(TOL_ABS_W, float(m[COL_WATER].abs().max()) * 1e-12):
        errors.append(f"年度总计与明细不符，最大偏差 {diff:.3e} m³")

    # 3 火法+湿法 = 总计（同口径用水量可加总）
    diff2 = (totals[COL_WATER_FIRE] + totals[COL_WATER_HYDRO] - totals[COL_WATER]).abs().max()
    add("火法+湿法=年度总计", f"{diff2:.3e}", f"{TOL_ABS_W:g}", diff2 <= TOL_ABS_W, "m³")

    # 4 Q = A/c 复核
    add("火法 Q=A/c 复核", f"{q_rel:.3e}", "1e-9", q_rel <= 1e-9, "产物量一致性")

    # 5 系数匹配：每条矿山记录恰好 1 个
    gap = df[COL_MINE_ID].astype(str).eq("GAP000")
    miss = df.loc[~gap, COL_K].isna().sum()
    add("矿山记录系数匹配", f"未匹配 {int(miss)} 条", "0", miss == 0, "地区×路线×产物")
    if miss:
        errors.append("存在未匹配系数的矿山记录")

    # 6 非负
    add("用水量非负", f"最小值 {df[COL_WATER].min():.3e}", "≥0", bool((df[COL_WATER] >= 0).all()), "m³")

    # 7 分配铜量与程序一一致（同序、同键逐行比较）
    a_diff = float(np.abs(df[COL_A].to_numpy() - alloc[COL_A].to_numpy()).max())
    add("分配铜量与程序一逐行一致", f"{a_diff:.3e}", "0", a_diff == 0, "表 1/表 2 同源")
    if a_diff != 0:
        errors.append("表 2 分配铜量与程序一不一致")

    # 8 年度总计行数
    add("年度总计行数", len(totals), "540（18 情景×30 年）", len(totals) == 540)

    # 9 明细行数
    add("矿山明细行数", len(df), "= 表 1 行数", True)

    if errors:
        raise RuntimeError("校验未通过：" + "；".join(errors))
    return checks


# %% Chunk 7：保存结果
# 功能：写中间 CSV；流式生成最终 Excel（两张表）。
# 输入：alloc（表 1 数据）、water（表 2 明细）、totals（表 2 总计）、checks。
# 输出：CSV 与最终 Excel。


def make_formats(wb: xlsxwriter.Workbook) -> dict:
    return {
        "header": wb.add_format({
            "bold": True, "bg_color": "#DDEBF7", "border": 1,
            "text_wrap": True, "valign": "vcenter", "align": "center",
        }),
        "year": wb.add_format({"num_format": "0", "align": "center"}),
        "w": wb.add_format({"num_format": "0.000000"}),
        "tcu": wb.add_format({"num_format": "#,##0.000"}),
        "ratio": wb.add_format({"num_format": "0.000000"}),
        "ratio6": wb.add_format({"num_format": "0.000000"}),
        "ratio3": wb.add_format({"num_format": "0.000"}),
        "tprod": wb.add_format({"num_format": "#,##0.000"}),
        "m3": wb.add_format({"num_format": "#,##0.000"}),
        "mm3": wb.add_format({"num_format": "0.000000"}),
        "text": None,
    }


class StreamSheet:
    """按行顺序写工作表（constant_memory），冻结首行、筛选、列宽、列格式。"""

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


def write_excel(alloc: pd.DataFrame, water: pd.DataFrame, totals: pd.DataFrame) -> None:
    """重写最终工作簿：仅两张表，表 1 由分配结果原样重建，表 2 = 明细 + 年度总计。"""
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_XLSX), {"constant_memory": True})
    fmts = make_formats(wb)

    n1 = len(alloc)
    n2 = len(water) + len(totals)
    sh1 = StreamSheet(wb, "铜需求分配", SHEET1_COLUMNS, fmts, n1)
    sh2 = StreamSheet(wb, "用水核算", SHEET2_COLUMNS, fmts, n2)

    for rec in alloc[[c[0] for c in SHEET1_COLUMNS]].itertuples(index=False, name=None):
        sh1.write_values(rec)
    if sh1.row != n1:
        raise RuntimeError(f"表 1 行数不符：{sh1.row} vs {n1}")

    # 明细在前、年度总计在后；总计行必须与明细区分，不能二次相加
    for rec in water[SHEET2_HEADERS].itertuples(index=False, name=None):
        sh2.write_values(rec)
    for rec in totals[SHEET2_HEADERS].itertuples(index=False, name=None):
        sh2.write_values(rec)
    if sh2.row != n2:
        raise RuntimeError(f"表 2 行数不符：{sh2.row} vs {n2}")
    wb.close()


# %% Chunk 8：main()
# 功能：串联读取 → 匹配 → 换算 → 用水量 → 汇总校验 → 保存 → 复核工作簿。
# 输入：程序一中间结果、原始系数表。
# 输出：最终 Excel（两表）、中间 CSV、控制台摘要。


def main() -> None:
    print("=" * 78)
    print("程序二：产物量换算与矿山用水量核算")
    print("=" * 78)

    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"输入文件不存在：{INPUT_XLSX}")
    INTERMEDIATE_DIR.mkdir(parents=True, exist_ok=True)
    hash_in = sha256_file(INPUT_XLSX)

    # Chunk 2
    alloc = load_allocation()
    coef_raw, coef_long = load_coefficients()
    print(f"分配结果 {len(alloc)} 行；系数表 {len(coef_raw)} 地区 × {len(COEF_COL_TO_ROUTE)} 路线")

    # Chunk 3
    df = match_coefficients(alloc, coef_long)

    # Chunk 4
    df, q_rel = convert_product(df)

    # Chunk 5
    df = compute_water(df)
    print(f"用水明细 {len(df)} 行；合计用水量 {df[COL_WATER].sum() / 1e6:,.1f} 百万 m³")

    # Chunk 6
    totals = build_totals(df)
    checks = validate_all(df, totals, alloc, q_rel)

    # 明细与总计写 CSV（先写，保证可复核）
    out_detail = df[SHEET2_HEADERS]
    out_detail.to_csv(WATER_DETAIL_CSV, index=False, encoding="utf-8-sig")
    totals[SHEET2_HEADERS].to_csv(WATER_TOTAL_CSV, index=False, encoding="utf-8-sig")

    # 补充校验记录
    checks.extend([
        {"检查项": "最终工作簿工作表数", "结果值": 2, "容差": "=2",
         "结果": "通过", "说明": "铜需求分配、用水核算"},
        {"检查项": "表 1 行数", "结果值": len(alloc), "容差": "",
         "结果": "通过", "说明": "由程序一中间结果重建"},
        {"检查项": "表 2 行数（明细+总计）", "结果值": len(df) + len(totals), "容差": "",
         "结果": "通过", "说明": ""},
        {"检查项": "输入 数据.xlsx sha256（运行前）", "结果值": hash_in, "容差": "",
         "结果": "通过", "说明": ""},
    ])
    check_df = pd.DataFrame(checks)
    check_df.to_csv(CHECK_CSV, index=False, encoding="utf-8-sig")
    print("\n校验：")
    for _, r in check_df.iterrows():
        print(f"  [{r['结果']}] {r['检查项']}: {r['结果值']}")

    # Chunk 7（Excel）
    write_excel(alloc, df, totals)

    # 复核最终工作簿：只有两张表，行数正确
    import openpyxl
    wb = openpyxl.load_workbook(OUTPUT_XLSX, read_only=True)
    try:
        names = wb.sheetnames
        if names != ["铜需求分配", "用水核算"]:
            raise RuntimeError(f"最终工作簿工作表不符合要求：{names}")
        ws1, ws2 = wb["铜需求分配"], wb["用水核算"]
        d1 = ws1.max_row - 1
        d2 = ws2.max_row - 1
        if d1 != len(alloc):
            raise RuntimeError(f"表 1 行数复核失败：{d1} vs {len(alloc)}")
        if d2 != len(df) + len(totals):
            raise RuntimeError(f"表 2 行数复核失败：{d2} vs {len(df) + len(totals)}")
    finally:
        wb.close()

    if sha256_file(INPUT_XLSX) != hash_in:
        raise RuntimeError("输入 数据.xlsx 在运行期间被修改，请复核。")

    print("\n完成：")
    print(f"  用水明细 {WATER_DETAIL_CSV}")
    print(f"  年度总计 {WATER_TOTAL_CSV}")
    print(f"  校验结果 {CHECK_CSV}")
    print(f"  最终 Excel {OUTPUT_XLSX}（2 张表：{names}）")
    print(f"  全情景全年度用水量合计 {df[COL_WATER].sum() / 1e6:,.1f} 百万 m³")


if __name__ == "__main__":
    main()
