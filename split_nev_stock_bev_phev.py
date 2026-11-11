# -*- coding: utf-8 -*-
"""
GCAM 新能源汽车保有量 —— 按区域存量结构拆分为 BEV / PHEV
=========================================================
适用数据（本项目口径，不可推广为所有 GCAM 数据的通用定义）：
    车辆保有量_面板 中 Technology 全部标记为 BEV 的数值，本次视为
    “BEV + PHEV 合计保有量”，再按区域存量结构拆成 BEV 与 PHEV 两条记录。

输入 1：GCAM_保有量与九类装机容量_2010-2050.xlsx  → 工作表“车辆保有量_面板”
输入 2：GCAM32_BEV_PHEV存量结构_Logit预测_最终结果_V2.xlsx → 工作表“10_GCAM32最终比例”
输出  ：同目录下 *_BEV_PHEV拆分.xlsx（保留输入 1 的全部原始工作表，追加 6 张结果表）

运行：python split_nev_stock_bev_phev.py
依赖：pandas、openpyxl（见运行说明中的安装命令）

作者说明：本脚本只读取两份输入文件，绝不写回；输出文件若已存在则自动加时间戳。
"""

import argparse
import datetime as dt
import hashlib
import math
import numbers
import os
import shutil
import sys

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================================
# 配置区（如需换文件，只改这里）
# ============================================================================

INPUT_STOCK = r"C:\Users\22965\OneDrive\桌面\MFA\Output\GCAM_保有量与九类装机容量_2010-2050.xlsx"
INPUT_SHARE = (
    r"C:\Users\22965\OneDrive\桌面\MFA\预备工作\output"
    r"\新能源汽车BEV和PHEV占比情况\GCAM32_BEV_PHEV存量结构_Logit预测_最终结果_V2.xlsx"
)
VEHICLE_SHEET = "车辆保有量_面板"      # 输入 1 中待拆分的车辆面板
SHARE_SHEET = "10_GCAM32最终比例"      # 输入 2 中唯一允许读取的工作表

OUTPUT_BASE = (
    r"C:\Users\22965\OneDrive\桌面\MFA\Output"
    r"\GCAM_保有量与九类装机容量_2010-2050_BEV_PHEV拆分.xlsx"
)
TIMESTAMP_FMT = "%Y%m%d_%H%M%S"        # 输出已存在时追加的时间戳格式

# 车辆面板的标识列（拆分后原样保留）
VEHICLE_KEY_COLS = ["Scenario", "Region", "Sector", "Subsector", "Technology", "单位"]

# 车辆面板的年份列（顺序即输出面板的年份顺序）
PANEL_YEARS = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
HISTORY_YEARS = [2010, 2015, 2020, 2025]                       # 三情景共用历史比例
FUTURE_YEARS = [y for y in PANEL_YEARS if y not in HISTORY_YEARS]  # 必须用“情景|年份”

# 情景映射：车辆表长字符串前缀 → 比例表列名前缀（标准化情景）
SCENARIO_PREFIX = [
    ("SSP1-2p6_", "SSP1-2.6"),
    ("SSP2-4p5_", "SSP2-4.5"),
    ("SSP5_8p5_", "SSP5-8.5"),
]

# 地区别名：车辆表名称 → 比例表“GCAM区域”名称（只做显式别名，不做模糊匹配）
REGION_ALIAS = {
    "Central America and Caribbean": "Central America and the Caribbean",
}

# 比例表列名
SHARE_REGION_COL = "GCAM区域"
SHARE_CN_COL = "中文名称"

# 新增工作表名
SHEET_PANEL = "车辆保有量_BEV_PHEV_面板"
SHEET_LONG = "车辆保有量_BEV_PHEV_长表"
SHEET_CHECK = "拆分核查"
SHEET_GROUP = "汇总校验_情景地区年份"
SHEET_ANOMALY = "异常清单"
SHEET_NOTES = "处理说明"

# 校验容差（浮点误差，绝不放宽到可能掩盖真实错误的程度）
TOL_ABS = 1e-6        # |BEV+PHEV - N| 的绝对容差
TOL_REL = 1e-12       # |BEV+PHEV - N| 的相对容差
RATIO_TOL = 1e-12     # |p + (1-p) - 1|
NUMFMT_STOCK = "0.0000"   # 保有量显示格式（保留小数显示，不改变实际值）
NUMFMT_RATIO = "0.0000"

# ---- 取整设置 ----
# ROUND_OUTPUT = True  ：结果四舍五入为整数
# ROUND_OUTPUT = False ：保留浮点全精度（本次要求：不做整数取整，保留小数精度）
ROUND_OUTPUT = False
# ROUND_MODE 仅在 ROUND_OUTPUT=True 时生效：
#   "complement"   BEV 四舍五入，PHEV = N − BEV → BEV+PHEV 严格等于原始总量（守恒优先，默认）
#   "independent"  BEV、PHEV 各自独立四舍五入 → 合计可能出现 ±1 差额
ROUND_MODE = "complement"

# 处理状态
ST_OK = "已拆分"
ST_PENDING = "待补充比例"
ST_ZERO_MISS = "零总量_比例缺失"
ST_N_INVALID = "异常_原始总量缺失或非数值"
ST_N_NEG = "异常_原始总量为负"
ST_P_INVALID = "异常_比例非数值"
ST_P_RANGE = "异常_比例超出[0,1]"
ST_REGION = "异常_地区无法匹配"
ST_SCEN = "异常_情景无法识别"
ST_DUPKEY = "异常_比例匹配键重复"
ST_COL_MISS = "异常_比例列缺失"

BLOCKING_STATUS = {ST_PENDING, ST_N_INVALID, ST_N_NEG, ST_P_INVALID, ST_P_RANGE,
                   ST_REGION, ST_SCEN, ST_DUPKEY, ST_COL_MISS}

# ============================================================================
# 通用小工具
# ============================================================================


def norm_label(value):
    """把表头统一成 int 年份或去空格字符串，兼容 Excel 年份为整数或字符串。"""
    if isinstance(value, bool):
        return str(value).strip()
    if isinstance(value, numbers.Real):
        f = float(value)
        if math.isnan(f):
            return ""
        return int(f) if f.is_integer() else value
    s = str(value).strip()
    if s == "":
        return s
    try:
        f = float(s)
    except ValueError:
        return s
    return int(f) if f.is_integer() else s


def to_number(value):
    """返回 (数值, 是否有效)。缺失 → (None, True)；非数值 → (None, False)。"""
    if value is None:
        return None, True
    if isinstance(value, bool):
        return None, False
    if isinstance(value, numbers.Real):
        f = float(value)
        return (None, True) if math.isnan(f) else (f, True)
    if isinstance(value, str):
        s = value.strip().replace(",", "")
        if s == "":
            return None, True
        try:
            f = float(s)
        except ValueError:
            return None, False
        return (None, True) if math.isnan(f) else (f, True)
    return None, False


def round_half_up(x):
    """四舍五入（.5 进位），避免 Python 内置 round() 的银行家舍入。"""
    f = float(x)
    if f < 0:
        return -int(math.floor(-f + 0.5))
    return int(math.floor(f + 0.5))


def conservation_tol(n):
    """守恒校验容差：浮点模式按数值量级，取整模式按取整方式给定。"""
    if not ROUND_OUTPUT:
        return max(TOL_ABS, TOL_REL * abs(n))
    return 0.0 if ROUND_MODE == "complement" else 1.0


def map_scenario(raw):
    """按配置前缀映射标准化情景；无法识别返回 None。"""
    s = "" if raw is None else str(raw)
    for prefix, std in SCENARIO_PREFIX:
        if s.startswith(prefix):
            return std
    return None


def resolve_output_path(base):
    """输出已存在则加时间戳，避免覆盖。"""
    if not os.path.exists(base):
        return base
    stem, ext = os.path.splitext(base)
    ts = dt.datetime.now().strftime(TIMESTAMP_FMT)
    cand = f"{stem}_{ts}{ext}"
    k = 2
    while os.path.exists(cand):
        cand = f"{stem}_{ts}_{k}{ext}"
        k += 1
    return cand


def sha256(path, chunk=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def fmt_num(x):
    """控制台/说明用的数字格式，避免把浮点尾巴写进文字。"""
    if x is None:
        return "空值"
    x = float(x)
    if x == int(x) and abs(x) < 1e15:
        return f"{int(x):,}"
    return f"{x:,.6f}"


# ============================================================================
# 读取输入
# ============================================================================


def read_vehicle_panel(path):
    df = pd.read_excel(path, sheet_name=VEHICLE_SHEET, engine="openpyxl")
    raw_cols = list(df.columns)
    df.columns = [norm_label(c) for c in raw_cols]

    missing_key = [c for c in VEHICLE_KEY_COLS if c not in df.columns]
    if missing_key:
        raise SystemExit(f"[终止] 车辆面板缺少字段：{missing_key}；实际列：{raw_cols}")

    missing_years = [y for y in PANEL_YEARS if y not in df.columns]
    if missing_years:
        raise SystemExit(f"[终止] 车辆面板缺少年份列：{missing_years}；实际列：{raw_cols}")

    # 标识列统一成字符串（去首尾空格），年份列保持原值以便逐格校验
    for c in VEHICLE_KEY_COLS:
        df[c] = df[c].apply(lambda v: "" if v is None else str(v).strip())
    return df


def read_share_table(path):
    """按“表头行 + 原始单元格”方式读取比例表，保留“缺失 / 非数值”的差别。"""
    raw = pd.read_excel(path, sheet_name=SHARE_SHEET, header=None, engine="openpyxl")
    if raw.shape[0] < 2:
        raise SystemExit(f"[终止] 比例表“{SHARE_SHEET}”为空。")

    header = [norm_label(v) for v in raw.iloc[0].tolist()]
    body = raw.iloc[1:].reset_index(drop=True)

    # 重复列名（同一键出现两次 → 比例匹配键重复，必须报出）
    dup_cols = sorted({h for h in header if header.count(h) > 1}, key=str)

    if SHARE_REGION_COL not in header:
        raise SystemExit(
            f"[终止] 比例表首行缺少“{SHARE_REGION_COL}”，实际表头：{header[:6]}…"
        )
    region_idx = header.index(SHARE_REGION_COL)
    cn_idx = header.index(SHARE_CN_COL) if SHARE_CN_COL in header else None

    # 只认识三个情景 × 2010–2060；历史年份列 = 整数列
    ratio_cols = [h for h in header if isinstance(h, int)]
    future_cols = [h for h in header if isinstance(h, str) and "|" in h]

    body.columns = header
    body = body.reset_index(drop=True)
    body[SHARE_REGION_COL] = body[SHARE_REGION_COL].apply(
        lambda v: "" if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip()
    )
    body = body[body[SHARE_REGION_COL] != ""]

    regions = body[SHARE_REGION_COL].tolist()
    dup_regions = sorted({r for r in regions if regions.count(r) > 1})
    region_order = list(dict.fromkeys(regions))

    keep = [SHARE_REGION_COL] + ratio_cols + future_cols
    table = body[keep]

    # 展开成 {(region, key): raw_value}
    values = {}
    for _, row in table.iterrows():
        reg = row[SHARE_REGION_COL]
        for col in ratio_cols + future_cols:
            values[(reg, col)] = row[col]

    return {
        "raw": raw,
        "header": header,
        "dup_cols": dup_cols,
        "dup_regions": dup_regions,
        "region_order": region_order,
        "regions": set(region_order),
        "history_cols": ratio_cols,
        "future_cols": future_cols,
        "values": values,
        "cn_idx": cn_idx,
    }


# ============================================================================
# 拆分主逻辑
# ============================================================================


def split_records(df, share):
    """返回 (cells, anomalies)。cells 为“每条源记录 × 每个年份”的明细。"""
    cells, anomalies = [], []
    share_regions = share["regions"]
    values = share["values"]
    dup_regions = set(share["dup_regions"])
    history_set = set(share["history_cols"])

    # 源记录重复检查（不删除、不求和，只报告）
    dup_key_cols = ["Scenario", "Region", "Sector", "Subsector", "Technology"]
    dup_mask = df.duplicated(subset=dup_key_cols, keep=False)
    for i in df.index[dup_mask]:
        anomalies.append({
            "type": "源记录重复",
            "severity": "警告",
            "rid": f"R{i + 1:04d}",
            "scenario": df.at[i, "Scenario"],
            "std": map_scenario(df.at[i, "Scenario"]) or "",
            "region": df.at[i, "Region"],
            "region_key": REGION_ALIAS.get(df.at[i, "Region"], df.at[i, "Region"]),
            "sector": df.at[i, "Sector"],
            "subsector": df.at[i, "Subsector"],
            "tech": df.at[i, "Technology"],
            "year": "全部",
            "n": None,
            "src_col": "",
            "p": None,
            "note": f"与其它行在 {'/'.join(dup_key_cols)} 上完全相同；未删除、未求和，请人工确认后处理",
        })

    for i in df.index:
        rid = f"R{i + 1:04d}"
        row = df.loc[i]
        scen_raw = row["Scenario"]
        std = map_scenario(scen_raw)
        reg_raw = row["Region"]
        reg_key = REGION_ALIAS.get(reg_raw, reg_raw)

        rec_err = None      # 记录级错误：整条记录所有年份同一个状态
        if std is None:
            rec_err = ST_SCEN
            anomalies.append({
                "type": "情景无法识别", "severity": "错误", "rid": rid,
                "scenario": scen_raw, "std": "", "region": reg_raw, "region_key": reg_key,
                "sector": row["Sector"], "subsector": row["Subsector"], "tech": row["Technology"],
                "year": "全部", "n": None, "src_col": "",
                "p": None,
                "note": f"Scenario 不以 {[p for p, _ in SCENARIO_PREFIX]} 任一前缀开头，无法确定情景",
            })
        elif reg_key not in share_regions:
            rec_err = ST_REGION
            anomalies.append({
                "type": "地区无法匹配", "severity": "错误", "rid": rid,
                "scenario": scen_raw, "std": std, "region": reg_raw, "region_key": reg_key,
                "sector": row["Sector"], "subsector": row["Subsector"], "tech": row["Technology"],
                "year": "全部", "n": None, "src_col": "",
                "p": None,
                "note": "比例表“GCAM区域”中无此名称（仅允许显式别名，不做模糊匹配/代理地区）",
            })
        elif reg_key in dup_regions:
            rec_err = ST_DUPKEY
            anomalies.append({
                "type": "比例匹配键重复", "severity": "错误", "rid": rid,
                "scenario": scen_raw, "std": std, "region": reg_raw, "region_key": reg_key,
                "sector": row["Sector"], "subsector": row["Subsector"], "tech": row["Technology"],
                "year": "全部", "n": None, "src_col": "",
                "p": None,
                "note": f"比例表区域“{reg_key}”出现多行，比例匹配键 (地区, 比例列) 不唯一",
            })

        for year in PANEL_YEARS:
            n_raw = row[year]
            n, n_ok = to_number(n_raw)

            if year in history_set:
                key = year
                src_col = str(year)
            else:
                key = f"{std}|{year}" if std else None
                src_col = f"{std}|{year}" if std else f"（情景未识别）|{year}"

            cell = {
                "rid": rid, "scenario": scen_raw, "std": std or "",
                "region": reg_raw, "region_key": reg_key,
                "sector": row["Sector"], "subsector": row["Subsector"],
                "tech": row["Technology"], "unit": row["单位"],
                "year": year, "src_col": src_col, "n": n, "n_raw": n_raw,
                "p": None, "phev_p": None, "bev": None, "phev": None,
                "bev_exact": None, "phev_exact": None,
                "total": None, "diff": None, "status": "", "note": "",
            }

            # 1) 记录级错误优先
            if rec_err is not None:
                cell["status"] = rec_err
                cell["note"] = "记录级匹配失败，未拆分"
                cells.append(cell)
                continue

            # 2) 原始总量缺失 / 非数值 / 负数
            if not n_ok:
                cell["status"] = ST_N_INVALID
                cell["note"] = f"原始总量不是有效数值（原值={n_raw!r}），保留为异常，不当作 0"
                anomalies.append(_anom_from_cell(cell, "原始总量缺失或非数值", "错误", cell["note"]))
                cells.append(cell)
                continue
            if n is None:
                cell["status"] = ST_N_INVALID
                cell["note"] = "原始总量为空，保留为异常，不当作 0"
                anomalies.append(_anom_from_cell(cell, "原始总量缺失或非数值", "错误", cell["note"]))
                cells.append(cell)
                continue
            if n < 0:
                cell["status"] = ST_N_NEG
                cell["note"] = f"原始总量为负（{n}）"
                anomalies.append(_anom_from_cell(cell, "原始总量为负", "错误", cell["note"]))
                cells.append(cell)
                continue

            # 3) 比例匹配
            if key is None or (key not in history_set and key not in set(share["future_cols"])):
                cell["status"] = ST_COL_MISS
                cell["note"] = f"比例表无列“{src_col}”"
                anomalies.append(_anom_from_cell(cell, "比例列缺失", "错误", cell["note"]))
                cells.append(cell)
                continue

            p_raw = values.get((reg_key, key))
            if (reg_key, key) not in values:
                cell["status"] = ST_COL_MISS
                cell["note"] = f"比例表“{reg_key}”行无列“{src_col}”"
                anomalies.append(_anom_from_cell(cell, "比例列缺失", "错误", cell["note"]))
                cells.append(cell)
                continue

            p, p_ok = to_number(p_raw)
            if not p_ok:
                cell["status"] = ST_P_INVALID
                cell["note"] = f"比例不是有效数值（原值={p_raw!r}）"
                anomalies.append(_anom_from_cell(cell, "比例非数值", "错误", cell["note"]))
                cells.append(cell)
                continue

            if p is None:
                # 比例缺失：区分“总量为 0”与“总量 > 0”
                if n == 0:
                    cell["status"] = ST_ZERO_MISS
                    cell["bev"], cell["phev"] = 0.0, 0.0
                    cell["total"], cell["diff"] = 0.0, 0.0
                    cell["note"] = "原始总量为零，比例缺失不影响拆分"
                    anomalies.append(_anom_from_cell(
                        cell, "比例缺失_总量为零", "提示", cell["note"]))
                else:
                    cell["status"] = ST_PENDING
                    cell["note"] = "比例缺失且原始总量大于 0：BEV/PHEV 保留空值，待补充比例"
                    anomalies.append(_anom_from_cell(
                        cell, "比例缺失_待补充", "待补充", cell["note"]))
                cells.append(cell)
                continue

            if not (0.0 <= p <= 1.0):
                cell["status"] = ST_P_RANGE
                cell["note"] = f"BEV 占比 {p} 超出 [0,1]"
                anomalies.append(_anom_from_cell(cell, "比例超出[0,1]", "错误", cell["note"]))
                cells.append(cell)
                continue

            # 4) 正常拆分（按配置决定是否四舍五入取整）
            bev_exact = n * p
            phev_exact = n * (1.0 - p)
            if ROUND_OUTPUT:
                bev = float(round_half_up(bev_exact))
                if ROUND_MODE == "complement" and float(n).is_integer():
                    # 守恒优先：PHEV 取补数，保证 BEV+PHEV 严格等于原始总量
                    phev = float(int(n) - int(bev))
                else:
                    phev = float(round_half_up(phev_exact))
            else:
                bev, phev = bev_exact, phev_exact
            total = bev + phev
            diff = total - n
            cell.update({
                "p": p, "phev_p": 1.0 - p, "bev": bev, "phev": phev,
                "bev_exact": bev_exact, "phev_exact": phev_exact,
                "total": total, "diff": diff, "status": ST_OK, "note": "",
            })
            if abs(diff) > conservation_tol(n):
                cell["status"] = "异常_拆分不守恒"
                cell["note"] = f"BEV+PHEV 与原始总量相差 {diff!r}，超出容差"
                anomalies.append(_anom_from_cell(cell, "拆分不守恒", "错误", cell["note"]))
            elif ROUND_OUTPUT and abs(diff) > 0:
                cell["note"] = (f"取整后 BEV+PHEV 与原始总量相差 {diff:+.0f}，"
                                f"属±1 取整误差（ROUND_MODE={ROUND_MODE}）")
            if abs((p + (1.0 - p)) - 1.0) > RATIO_TOL:
                cell["note"] = (cell["note"] + "；") if cell["note"] else ""
                cell["note"] += "BEV占比+PHEV占比 偏离 1"
            cells.append(cell)

    return cells, anomalies


def _anom_from_cell(cell, anom_type, severity, note):
    return {
        "type": anom_type, "severity": severity, "rid": cell["rid"],
        "scenario": cell["scenario"], "std": cell["std"],
        "region": cell["region"], "region_key": cell["region_key"],
        "sector": cell["sector"], "subsector": cell["subsector"],
        "tech": cell["tech"], "year": cell["year"], "n": cell["n"],
        "src_col": cell["src_col"], "p": cell["p"], "note": note,
    }


# ============================================================================
# 结果表构造
# ============================================================================


def build_panel(cells, df):
    """每条源记录 → BEV / PHEV 两行，字段顺序与原车辆面板一致。"""
    by_rec = {}
    for c in cells:
        by_rec.setdefault(c["rid"], {})[c["year"]] = c

    header = VEHICLE_KEY_COLS + PANEL_YEARS
    rows = []
    for i in df.index:
        rid = f"R{i + 1:04d}"
        row = df.loc[i]
        years = by_rec.get(rid, {})
        for tech in ("BEV", "PHEV"):
            vals = []
            for year in PANEL_YEARS:
                c = years.get(year)
                if c is None:
                    vals.append(None)
                elif c["status"] == ST_OK:
                    vals.append(c["bev"] if tech == "BEV" else c["phev"])
                elif c["status"] == ST_ZERO_MISS:
                    vals.append(0.0)
                else:
                    vals.append(None)   # 待补充 / 异常：保留空值，绝不填 0
            rows.append([
                row["Scenario"], row["Region"], row["Sector"], row["Subsector"],
                tech, row["单位"],
            ] + vals)
    return header, rows


def build_long(cells):
    header = ["源记录ID", "Scenario", "标准情景", "Region", "Sector", "Subsector",
              "Technology", "单位", "Year", "Stock", "处理状态", "备注"]
    rows = []
    for c in cells:
        for tech in ("BEV", "PHEV"):
            if c["status"] == ST_OK:
                stock = c["bev"] if tech == "BEV" else c["phev"]
            elif c["status"] == ST_ZERO_MISS:
                stock = 0.0
            else:
                stock = None        # 待补充 / 异常保持空值
            rows.append([
                c["rid"], c["scenario"], c["std"], c["region"], c["sector"],
                c["subsector"], tech, c["unit"], c["year"], stock,
                c["status"], c["note"],
            ])
    return header, rows


def build_check(cells):
    header = ["源记录ID", "原始情景", "标准情景", "原始地区", "匹配地区", "Sector",
              "Subsector", "年份", "原始合计保有量", "比例来源列", "BEV占比", "PHEV占比",
              "BEV保有量", "PHEV保有量", "拆分合计", "差额", "处理状态", "备注"]
    rows = []
    for c in cells:
        rows.append([
            c["rid"], c["scenario"], c["std"], c["region"], c["region_key"],
            c["sector"], c["subsector"], c["year"], c["n"], c["src_col"],
            c["p"], c["phev_p"], c["bev"], c["phev"], c["total"], c["diff"],
            c["status"], c["note"],
        ])
    return header, rows


def build_group_check(cells):
    """按 标准化情景 × 原始地区 × 年份 汇总守恒；含待补充/异常的分组标记“未完成”。"""
    header = ["标准情景", "原始地区", "年份", "原始合计保有量", "拆分合计",
              "差额", "记录数", "已完成记录数", "待补充记录数", "异常记录数", "校验状态"]
    buckets = {}
    for c in cells:
        key = (c["std"] or "（未识别）", c["region"], c["year"])
        b = buckets.setdefault(key, {"n": 0.0, "total": 0.0, "rows": 0,
                                     "ok": 0, "pending": 0, "err": 0, "bad": 0})
        b["rows"] += 1
        if c["n"] is not None:
            b["n"] += c["n"]
        st = c["status"]
        if st in (ST_OK, ST_ZERO_MISS):
            b["ok"] += 1
            b["total"] += c["total"]
        elif st == ST_PENDING:
            b["pending"] += 1
            b["bad"] += 1
        else:
            b["err"] += 1
            b["bad"] += 1

    rows = []
    for (std, region, year) in sorted(buckets, key=lambda k: (k[0], k[1], k[2])):
        b = buckets[(std, region, year)]
        diff = b["total"] - b["n"]
        if b["bad"] > 0:
            state = "未完成"
            diff_cell = None        # 不完整分组不报差额，避免求和跳过空值造成“虚通过”
            total_cell = None
        else:
            state = "通过" if abs(diff) <= max(TOL_ABS, TOL_REL * abs(b["n"])) else "不守恒"
            diff_cell = diff
            total_cell = b["total"]
        rows.append([std, region, year, b["n"], total_cell, diff_cell,
                     b["rows"], b["ok"], b["pending"], b["err"], state])
    return header, rows


def build_anomaly_sheet(anomalies):
    header = ["序号", "异常类型", "严重程度", "源记录ID", "原始情景", "标准情景", "原始地区",
              "匹配地区", "Sector", "Subsector", "Technology", "年份",
              "原始合计保有量", "比例来源列", "比例值", "说明"]
    order = {"错误": 0, "待补充": 1, "警告": 2, "提示": 3}
    rows = []
    for k, a in enumerate(sorted(anomalies, key=lambda x: (order.get(x["severity"], 9),
                                                          x["rid"] or "", str(x["year"]))), 1):
        rows.append([k, a["type"], a["severity"], a["rid"], a["scenario"], a["std"],
                     a["region"], a["region_key"], a["sector"], a["subsector"], a["tech"],
                     a["year"], a["n"], a["src_col"], a["p"], a["note"]])
    return header, rows


# ============================================================================
# 工作簿写入
# ============================================================================

HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_ALIGN = Alignment(horizontal="center", vertical="center")


def write_sheet(wb, name, header, rows, num_cols=(), numfmt=None, widths=None,
                header_fill=HDR_FILL):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(header)
    for col_idx in range(1, len(header) + 1):
        cell = ws.cell(1, col_idx)
        cell.font = HDR_FONT
        cell.fill = header_fill
        cell.alignment = HDR_ALIGN
    for r in rows:
        ws.append(r)
    if num_cols and numfmt:
        for col_idx in num_cols:
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row_idx, col_idx)
                if cell.value is not None:
                    cell.number_format = numfmt
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(ws.max_row, 1)}"
    return ws


def compare_original_sheets(src_path, out_path, sheet_names):
    """逐格核对原始工作表在输出文件中的一致性（值相同，数值按容差比较）。"""
    src = openpyxl.load_workbook(src_path, read_only=True, data_only=False)
    out = openpyxl.load_workbook(out_path, read_only=True, data_only=False)
    details = []
    ok_all = True
    try:
        for name in sheet_names:
            if name not in src.sheetnames or name not in out.sheetnames:
                details.append(f"{name}: 缺失")
                ok_all = False
                continue
            a, b = src[name], out[name]
            diffs, cells = 0, 0
            for ra, rb in zip(a.iter_rows(values_only=True), b.iter_rows(values_only=True)):
                for va, vb in zip(ra, rb):
                    cells += 1
                    if not _same_value(va, vb):
                        diffs += 1
            details.append(f"{name}: {cells} 个单元格比对，差异 {diffs} 个")
            ok_all = ok_all and diffs == 0
    finally:
        src.close()
        out.close()
    return ok_all, details


def _same_value(a, b):
    if a is None and b is None:
        return True
    if isinstance(a, numbers.Real) and isinstance(b, numbers.Real):
        if isinstance(a, bool) or isinstance(b, bool):
            return a == b
        fa, fb = float(a), float(b)
        if math.isnan(fa) and math.isnan(fb):
            return True
        return fa == fb or math.isclose(fa, fb, rel_tol=TOL_REL, abs_tol=1e-9)
    return str(a) == str(b)


# ============================================================================
# 主流程
# ============================================================================


def main(argv=None):
    ap = argparse.ArgumentParser(description="GCAM 新能源汽车保有量 BEV/PHEV 拆分")
    ap.add_argument("--stock", default=INPUT_STOCK, help="输入 1（含车辆保有量_面板）")
    ap.add_argument("--share", default=INPUT_SHARE, help="输入 2（BEV/PHEV 存量结构）")
    ap.add_argument("--out", default=OUTPUT_BASE, help="输出文件基础路径")
    args = ap.parse_args(argv)

    print("=" * 78)
    print("GCAM 新能源汽车保有量 → BEV / PHEV 拆分")
    print("=" * 78)
    for p in (args.stock, args.share):
        if not os.path.exists(p):
            raise SystemExit(f"[终止] 找不到输入文件：{p}")

    # ---------- 读取 ----------
    print(f"[1/7] 读取车辆面板：{os.path.basename(args.stock)}")
    df = read_vehicle_panel(args.stock)
    print(f"      工作表“{VEHICLE_SHEET}”记录数 = {len(df)}，年份列 = {PANEL_YEARS}")
    print("      情景取值：")
    for s in df["Scenario"].unique():
        print(f"        {s}  →  {map_scenario(s)}")

    print(f"[2/7] 读取比例表：{os.path.basename(args.share)} / “{SHARE_SHEET}”")
    share = read_share_table(args.share)
    print(f"      比例表区域数 = {len(share['regions'])}；历史列 = "
          f"{min(share['history_cols'])}–{max(share['history_cols'])}；"
          f"情景列 {len(share['future_cols'])} 个")
    if share["dup_cols"]:
        print(f"      [警告] 重复列名：{share['dup_cols']}")
    if share["dup_regions"]:
        print(f"      [警告] 重复区域行：{share['dup_regions']}")

    # ---------- 拆分 ----------
    print("[3/7] 逐记录 × 逐年拆分 ...")
    cells, anomalies = split_records(df, share)

    # 表头级异常也进清单
    for c in share["dup_cols"]:
        anomalies.append({"type": "比例表重复列名", "severity": "错误", "rid": "",
                          "scenario": "", "std": "", "region": "", "region_key": "",
                          "sector": "", "subsector": "", "tech": "", "year": c,
                          "n": None, "src_col": str(c), "p": None,
                          "note": f"比例表表头“{c}”出现多次，比例匹配键不唯一"})

    # ---------- 统计 ----------
    n_rec = len(df)
    n_year = len(PANEL_YEARS)
    counts = {st: 0 for st in [ST_OK, ST_ZERO_MISS, ST_PENDING, ST_N_INVALID, ST_N_NEG,
                               ST_P_INVALID, ST_P_RANGE, ST_REGION, ST_SCEN,
                               ST_DUPKEY, ST_COL_MISS, "异常_拆分不守恒"]}
    for c in cells:
        counts[c["status"]] = counts.get(c["status"], 0) + 1

    ok_cells = counts[ST_OK]
    zero_miss = counts[ST_ZERO_MISS]
    pending = counts[ST_PENDING]
    other_err = sum(v for k, v in counts.items()
                    if k not in (ST_OK, ST_ZERO_MISS, ST_PENDING))

    # 地区 / 情景覆盖率（按记录，不按单元格）
    rec_region_ok = sum(1 for i in df.index
                        if REGION_ALIAS.get(df.at[i, "Region"], df.at[i, "Region"]) in share["regions"])
    rec_scen_ok = sum(1 for i in df.index if map_scenario(df.at[i, "Scenario"]) is not None)
    ratio_cells_total = len(cells)
    ratio_cells_found = sum(
        1 for c in cells
        if (c["region_key"], c["year"] if c["year"] in set(share["history_cols"])
            else f"{c['std']}|{c['year']}") in share["values"]
        and to_number(share["values"][(c["region_key"], c["year"] if c["year"] in set(share["history_cols"])
                                       else f"{c['std']}|{c['year']}")])[0] is not None)

    # ---------- 预期条数（动态） ----------
    exp_panel = n_rec * 2
    exp_long = n_rec * 2 * n_year
    exp_check = n_rec * n_year

    # ---------- 输出路径 ----------
    out_path = resolve_output_path(args.out)
    print(f"[4/7] 输出文件：{out_path}")
    shutil.copy2(args.stock, out_path)      # 先原样复制，保留全部原始工作表

    # ---------- 构造结果 ----------
    print("[5/7] 构造结果表 ...")
    panel_hdr, panel_rows = build_panel(cells, df)
    long_hdr, long_rows = build_long(cells)
    check_hdr, check_rows = build_check(cells)
    group_hdr, group_rows = build_group_check(cells)
    anom_hdr, anom_rows = build_anomaly_sheet(anomalies)

    wb = openpyxl.load_workbook(out_path)
    write_sheet(wb, SHEET_PANEL, panel_hdr, panel_rows,
                num_cols=range(len(VEHICLE_KEY_COLS) + 1, len(panel_hdr) + 1),
                numfmt=NUMFMT_STOCK,
                widths=[46, 34, 20, 22, 12, 6] + [14] * n_year)
    write_sheet(wb, SHEET_LONG, long_hdr, long_rows,
                num_cols=[10], numfmt=NUMFMT_STOCK,
                widths=[11, 46, 12, 34, 20, 22, 12, 6, 8, 16, 20, 60])
    write_sheet(wb, SHEET_CHECK, check_hdr, check_rows,
                num_cols=[9, 11, 12, 13, 14, 15, 16], numfmt=NUMFMT_STOCK,
                widths=[11, 46, 12, 34, 34, 20, 22, 8, 16, 18, 10, 10, 16, 16, 16, 12, 20, 60])
    write_sheet(wb, SHEET_GROUP, group_hdr, group_rows,
                num_cols=[4, 5, 6], numfmt=NUMFMT_STOCK,
                widths=[12, 34, 8, 18, 18, 12, 10, 12, 14, 12, 10])
    write_sheet(wb, SHEET_ANOMALY, anom_hdr, anom_rows,
                num_cols=[13, 15], numfmt=NUMFMT_STOCK,
                widths=[6, 24, 10, 11, 46, 12, 34, 34, 20, 22, 12, 8, 16, 18, 10, 70])

    # ---------- 校验 ----------
    print("[6/7] 校验 ...")
    # 单元格级守恒
    bad_tol = [c for c in cells
               if c["status"] == ST_OK and abs(c["diff"]) > conservation_tol(c["n"])]
    bad_ratio = [c for c in cells
                 if c["status"] == ST_OK and abs((c["p"] + c["phev_p"]) - 1.0) > RATIO_TOL]
    # 取整相关统计
    ok_cell_list = [c for c in cells if c["status"] == ST_OK]
    max_round_err = max(
        (max(abs(c["bev"] - c["bev_exact"]), abs(c["phev"] - c["phev_exact"]))
         for c in ok_cell_list), default=0.0)
    if ROUND_OUTPUT:
        indep_diff = sum(
            1 for c in ok_cell_list
            if round_half_up(c["phev_exact"]) != int(c["phev"])
            or round_half_up(c["bev_exact"]) != int(c["bev"]))
    else:
        indep_diff = 0
    non_int = sum(
        1 for c in ok_cell_list
        for v in (c["bev"], c["phev"]) if abs(float(v) - round(float(v))) > 1e-9)
    # 分组级
    groups_total = len(group_rows)
    groups_pass = sum(1 for r in group_rows if r[-1] == "通过")
    groups_fail = sum(1 for r in group_rows if r[-1] == "不守恒")
    groups_unfinished = sum(1 for r in group_rows if r[-1] == "未完成")

    # 面板/长表/核查 行数
    got_panel = len(panel_rows)
    got_long = len(long_rows)
    got_check = len(check_rows)

    # 待补充记录定位（用于说明）
    pending_cells = [c for c in cells if c["status"] == ST_PENDING]
    pending_keys = sorted({(c["region"], c["year"]) for c in pending_cells})

    complete = (pending == 0 and other_err == 0)

    # ---------- 处理说明 ----------
    notes = []
    add = notes.append
    add(("处理说明", ""))
    add(("一、任务与口径", ""))
    add(("说明", "本脚本将 GCAM 新能源汽车保有量按地区/情景/年份的 BEV 存量占比拆分为 BEV 与 PHEV 两条记录。"))
    add(("数据口径（重要）", "本次输入的“车辆保有量_面板”中 Technology 全部标记为 BEV；按本项目口径，"
                        "这些数值实际是 BEV 与 PHEV 的合计保有量，据此拆分。该口径仅适用于本次输入数据，"
                        "不得推广为所有 GCAM 数据的通用定义。"))
    add(("来源标签说明", "原始工作表中的 BEV 标签沿用源文件含义（合计口径下的单一标签）；"
                     "后续分析请使用本文件新增的拆分结果表，避免与原始表重复统计。"))
    add(("", ""))
    add(("二、输入文件", ""))
    add(("输入 1", args.stock))
    add(("输入 1 SHA256", sha256(args.stock)))
    add(("读取工作表", f"{VEHICLE_SHEET}（原始工作表全部保留，未做任何修改）"))
    add(("输入 2", args.share))
    add(("输入 2 SHA256", sha256(args.share)))
    add(("读取工作表", f"{SHARE_SHEET}（仅读此表，未读取其它工作表）"))
    add(("", ""))
    add(("三、计算公式", ""))
    add(("BEV保有量", "BEV = N × p"))
    add(("PHEV保有量", "PHEV = N × (1 − p)"))
    add(("N", "该条原始记录的当年合计保有量"))
    add(("p", "该地区、该情景（历史年份为共用）、该年份的 BEV 占比（0–1 小数，不再除以 100）"))
    if ROUND_OUTPUT:
        add(("精度/取整", "结果四舍五入取整为整数（.5 进位，非银行家舍入）；表内数字均为整数，不再保留小数。"))
        if ROUND_MODE == "complement":
            add(("取整规则", "BEV = 四舍五入(N×p)；PHEV = N − BEV（取补数）。"
                          "因此 BEV+PHEV 严格等于原始合计保有量，逐个单元格守恒差额为 0；"
                          "个别单元格的 PHEV 会与 N×(1−p) 的独立四舍五入相差 1（这是守恒优先的必然取舍）。"))
        else:
            add(("取整规则", "BEV = 四舍五入(N×p)，PHEV = 四舍五入(N×(1−p))，两者独立取整；"
                          "个别单元格合计可能与原始总量相差 ±1，已在备注列标注。"))
    else:
        add(("精度", "保留浮点全精度，不做整数取整；表内仅设置显示小数位，不改变存储值。"))
    add(("拆分行", "每条原始记录拆成 BEV、PHEV 两条，其余标识字段原样保留；车型不合并、不跨车型汇总。"))
    add(("", ""))
    add(("四、情景映射", ""))
    for prefix, std in SCENARIO_PREFIX:
        add((f"前缀 {prefix}", f"标准化为 {std}（原始 Scenario 保持不变，仅用于匹配与核查）"))
    add(("未识别情景", "报入“异常清单”，不猜测、不代理。"))
    add(("", ""))
    add(("五、地区匹配与别名", ""))
    add(("匹配方式", "用车辆表 Region 匹配比例表“GCAM区域”，不依赖行顺序，不使用中文名称。"))
    for k, v in REGION_ALIAS.items():
        add((f"别名 {k}", f"→ {v}（输出仍保留车辆表原始地区名称）"))
    add(("无法匹配", "列入异常清单，不做模糊匹配、不寻找代理地区。"))
    add(("", ""))
    add(("六、年份匹配", ""))
    add(("历史年份", "2010、2015、2020、2025 使用比例表同年历史比例，三个情景共用。"))
    add(("未来年份", "2030、2035、2040、2045、2050 使用“标准情景|年份”列（如 SSP2-4.5|2040）。"))
    add(("不使用", "2055、2060 列；不新增年度插值；不重新拟合 Logit 模型。"))
    add(("表头兼容", "年份表头为整数或字符串均可识别（本文件实际为字符串）。"))
    add(("", ""))
    add(("七、跨车型统一比例假设", ""))
    add(("假设内容", "同一地区、同一情景、同一年份的 BEV 占比，统一应用于表内全部车型："
                 "Car、Mini Car、Large Car and Truck、Bus、Light truck、Medium truck、Heavy truck。"))
    add(("性质说明", "这是“统一使用区域存量结构”的建模假设；比例表本身并未提供分车型占比，"
                 "因此结果不体现车型间 BEV 渗透差异。"))
    add(("", ""))
    add(("八、缺失值与异常处理规则", ""))
    add(("N=0 且比例缺失", "BEV、PHEV 均记 0，状态“零总量_比例缺失”，异常清单注明“原始总量为零，比例缺失不影响拆分”。"))
    add(("N>0 且比例缺失", "BEV、PHEV 保留空值，状态“待补充比例”；异常清单保留原始总量与完整匹配信息；"
                       "不当作 0、不默认归入某一技术、不用邻近年份或其他地区填补。"))
    add(("N 缺失/非数值", "保留为异常，不当作 0。"))
    add(("N 为负", "报异常，不拆分。"))
    add(("比例非数值/超范围", "报异常，不静默修正。"))
    add(("地区/情景无法匹配", "报异常。"))
    add(("比例匹配键重复", "比例表同一区域多行或同名列重复时报异常，不擅自选取其中一行。"))
    add(("源记录重复", "不删除、不求和，逐条列入异常清单待人工确认。"))
    add(("部分无法拆分时", f"仍导出已完成结果与异常清单；本次结论见下节。"))
    add(("", ""))
    add(("九、校验结果", ""))
    add(("原始记录数", f"{n_rec}（动态读取，非硬编码）"))
    add(("年份数", f"{n_year}"))
    add(("拆分面板预期/实际", f"{exp_panel} / {got_panel} 行  ——  {'一致' if exp_panel == got_panel else '不一致！'}"))
    add(("长表预期/实际", f"{exp_long} / {got_long} 行  ——  {'一致' if exp_long == got_long else '不一致！'}"))
    add(("拆分核查预期/实际", f"{exp_check} / {got_check} 行  ——  {'一致' if exp_check == got_check else '不一致！'}"))
    add(("成功拆分单元格", f"{ok_cells} / {ratio_cells_total}"))
    add(("零总量且比例缺失", f"{zero_miss} 个单元格（BEV=PHEV=0，不影响合计）"))
    add(("待补充比例", f"{pending} 个单元格，涉及：" +
                    ("；".join(f"{r} {y}" for r, y in pending_keys) if pending_keys else "无")))
    add(("其他异常", f"{other_err} 个单元格"))
    add(("地区匹配覆盖率", f"{rec_region_ok}/{n_rec} = {rec_region_ok / n_rec:.2%}"))
    add(("情景匹配覆盖率", f"{rec_scen_ok}/{n_rec} = {rec_scen_ok / n_rec:.2%}"))
    add(("比例取值覆盖率", f"{ratio_cells_found}/{ratio_cells_total} = {ratio_cells_found / ratio_cells_total:.2%}"))
    tol_desc = ("0（取整-补数模式，要求严格相等）" if (ROUND_OUTPUT and ROUND_MODE == "complement")
                else "±1（取整-独立模式）" if ROUND_OUTPUT
                else f"max({TOL_ABS}, {TOL_REL}×|N|)")
    add(("单元格守恒校验", f"已拆分单元格中 BEV+PHEV≠N 的个数 = {len(bad_tol)}（容差 {tol_desc}）"))
    if ROUND_OUTPUT:
        add(("取整方式", f"ROUND_OUTPUT=True, ROUND_MODE={ROUND_MODE}"))
        add(("取整最大偏差", f"取整值与精确值的最大绝对偏差 = {max_round_err:.4f}（理论上限 0.5）"))
        add(("非整数残留", f"已拆分单元格中仍带小数的值个数 = {non_int}（应为 0）"))
        add(("与独立四舍五入的差异",
             f"{indep_diff} 个单元格的 PHEV 与“各自独立四舍五入”结果相差 1"
             if indep_diff else "无（两数取整结果与独立四舍五入完全一致）"))
    add(("占比和校验", f"BEV占比+PHEV占比 偏离 1（> {RATIO_TOL}）的个数 = {len(bad_ratio)}"))
    add(("分组守恒（情景×地区×年份）",
         f"共 {groups_total} 组：通过 {groups_pass}、不守恒 {groups_fail}、未完成 {groups_unfinished}"))
    add(("原始工作表一致性", "见下一节逐表比对结果"))
    add(("", ""))
    add(("十、总体结论", ""))
    if complete and groups_fail == 0:
        add(("结论", "全部记录已完成拆分；单元格与分组守恒校验通过。"))
    elif pending > 0:
        add(("结论", f"本次处理尚未全部完成：{pending} 个单元格因比例缺失（N>0）标记“待补充比例”，"
                    f"BEV/PHEV 保留空值；{groups_unfinished} 个分组因此标记“未完成”，"
                    f"不能宣称整体守恒校验通过。其余记录已完成拆分。"))
    else:
        add(("结论", f"存在 {other_err} 个其它异常单元格，处理未全部完成，请先处理异常清单。"))
    add(("", ""))
    add(("十一、输出表清单", ""))
    add((SHEET_PANEL, f"{got_panel} 行（{n_rec} 条源记录 × BEV/PHEV），字段顺序同原车辆面板"))
    add((SHEET_LONG, f"{got_long} 行，长表；待补充数据保留空值与状态"))
    add((SHEET_CHECK, f"{got_check} 行，逐记录逐年核查；未完成拆分的记录“拆分合计/差额”为空"))
    add((SHEET_GROUP, f"{groups_total} 行，情景×地区×年份 汇总守恒；含待补充的分组标“未完成”"))
    add((SHEET_ANOMALY, f"{len(anom_rows)} 行，全部缺失比例、匹配失败、无效数值、重复键等问题"))
    add(("原始工作表", "、".join(wb.sheetnames[:4]) + " 等（原样保留）"))

    # 一致性比对（写盘之后再读一次，确保是最终文件）
    wb.save(out_path)
    src_wb = openpyxl.load_workbook(args.stock, read_only=True)
    original_sheets = list(src_wb.sheetnames)
    src_wb.close()
    same_ok, same_details = compare_original_sheets(args.stock, out_path, original_sheets)
    add(("十二、原始工作表一致性逐表核对", "全部一致" if same_ok else "存在差异，详见下行"))
    for d in same_details:
        add(("", d))

    # 把“一致性”结论回填后重写说明表（说明表最后写，保证内容含校验结果）
    if SHEET_NOTES in wb.sheetnames:
        del wb[SHEET_NOTES]
    ws = wb.create_sheet(SHEET_NOTES)
    ws.append(["项目", "内容"])
    for col_idx in (1, 2):
        cell = ws.cell(1, col_idx)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = HDR_ALIGN
    for r in notes:
        ws.append(list(r))
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 130
    for row in ws.iter_rows(min_row=2, max_col=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
        row[0].alignment = Alignment(vertical="top")
    ws.freeze_panes = "A2"
    wb.save(out_path)

    # ---------- 控制台汇总 ----------
    print("[7/7] 完成。")
    print("-" * 78)
    print(f"输出文件        : {out_path}")
    print(f"原始记录数      : {n_rec}")
    print(f"拆分面板        : {got_panel} 行（预期 {exp_panel}）")
    print(f"长表            : {got_long} 行（预期 {exp_long}）")
    print(f"拆分核查        : {got_check} 行（预期 {exp_check}）")
    print(f"成功拆分单元格  : {ok_cells}")
    print(f"零总量且缺比例  : {zero_miss}")
    print(f"待补充比例      : {pending}")
    print(f"其它异常单元格  : {other_err}")
    print(f"分组守恒        : 通过 {groups_pass} / 不守恒 {groups_fail} / 未完成 {groups_unfinished}")
    if ROUND_OUTPUT:
        print(f"取整方式        : 四舍五入取整，ROUND_MODE={ROUND_MODE}"
              f"（最大取整偏差 {max_round_err:.4f}，非整数残留 {non_int}）")
    else:
        print("取整方式        : 不取整（保留浮点全精度）")
    print(f"原始表一致性    : {'通过' if same_ok else '未通过'}")
    print(f"处理状态        : {'全部完成' if complete and groups_fail == 0 else '尚未全部完成（见异常清单）'}")
    print("-" * 78)
    if pending or other_err:
        print("提示：存在未完成拆分的记录，请查看输出文件中的“异常清单”与“处理说明”。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
