# -*- coding: utf-8 -*-
"""第二步：三个回收情景 × 两个废铜贸易情景下的逐年次生铜产量。

读取第一步正式成果（不重算退役模型），计算：
  不开放：Q[s,r,k,t] = W[r,k,t] × C[s,r,k,t] × P[s,r,k,t]（仅15个有处理能力区域）
  全球开放：X[s,r,j,k,t] = W[r,k,t] × C[s,r,k,t] × T[r,j]
            Q[s,r,j,k,t] = X[s,r,j,k,t] × P[s,j,k,t]
来源区收集率、处理区综合处理率分别取自废铜产生地与出口目的地/实际处理地。

输出
----
能源转型设备次生铜产量_第二步.xlsx（12个工作表）
能源转型设备退役与次生铜核算_方法与文献说明.md（核验通过后生成）

本脚本 import 时不执行计算，只在 __main__ 下运行。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import time
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
import openpyxl
import xlsxwriter

PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = PROJECT_DIR / "config.json"

RUN_LOG: list[str] = []

REC_ORDER = ["基准", "政策推动", "技术进步"]
TRADE_ORDER = ["不开放", "全球开放"]
DEPTS = ["车辆", "绿电"]


def log(msg: str = "") -> None:
    print(msg)
    RUN_LOG.append(str(msg))


# --------------------------------------------------------------------------- #
# 通用工具（与第一步保持同一口径，脚本各自独立可运行）
# --------------------------------------------------------------------------- #
def sha256_of(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_config(path: str | Path | None = None) -> dict:
    cfg_path = Path(path) if path else DEFAULT_CONFIG
    with open(cfg_path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["_config_path"] = str(cfg_path)
    cfg["_config_sha256"] = sha256_of(cfg_path)
    return cfg


def _copy_shared_read(src: Path, dst: Path) -> None:
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create = k32.CreateFileW
    create.restype = wintypes.HANDLE
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                       ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    handle = create(str(src), 0x80000000, 0x1 | 0x2, None, 3, 0x80, None)
    if handle in (None, wintypes.HANDLE(-1).value):
        raise PermissionError(f"无法以共享只读方式打开：{src}")
    try:
        with open(dst, "wb") as out:
            buf = ctypes.create_string_buffer(1 << 20)
            while True:
                nread = wintypes.DWORD(0)
                ok = k32.ReadFile(handle, buf, len(buf), ctypes.byref(nread), None)
                if not ok or nread.value == 0:
                    break
                out.write(buf.raw[:nread.value])
    finally:
        k32.CloseHandle(handle)


def open_workbook_readonly(path: str | Path, fallback_dir: str | Path | None = None):
    src = Path(path)
    try:
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        return wb, src, "original", sha256_of(src)
    except PermissionError:
        pass
    cache_dir = Path(fallback_dir) if fallback_dir else src.parent
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached = cache_dir / src.name
    if cached.exists() and cached.is_file():
        wb = openpyxl.load_workbook(cached, read_only=True, data_only=True)
        return wb, cached, "cache_copy", sha256_of(cached)
    _copy_shared_read(src, cached)
    wb = openpyxl.load_workbook(cached, read_only=True, data_only=True)
    return wb, cached, "shared_copy", sha256_of(cached)


def to_number(value):
    if value is None or isinstance(value, bool):
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


def is_year_token(value) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return float(value).is_integer() and 1900 <= int(value) <= 2100
    text = str(value).strip()
    return text.isdigit() and 1900 <= int(text) <= 2100


def header_index(header, *names):
    lookup = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}
    for name in names:
        key = str(name).strip().lower()
        if key in lookup:
            return lookup[key]
    return None


def norm_scenario(value: str) -> str:
    return str(value).strip().replace("_", "-").replace(" ", "")


def git_head() -> str:
    try:
        out = subprocess.run(["git", "rev-parse", "HEAD"], cwd=str(PROJECT_DIR),
                             capture_output=True, text=True, timeout=20)
        return out.stdout.strip() or "unknown"
    except Exception:
        return "unknown"


# --------------------------------------------------------------------------- #
# 读取第一步成果
# --------------------------------------------------------------------------- #
def load_step1(cfg, path=None):
    path = Path(path) if path else Path(cfg["paths"]["step1_out"])
    if not path.exists():
        raise FileNotFoundError(f"第一步成果不存在：{path}")
    wb, used, src_tag, digest = open_workbook_readonly(path, Path(cfg["paths"]["fallback_dir"]))
    log(f"[输入] 第一步成果：{path}")
    log(f"       实际读取：{used}（来源={src_tag}）sha256={digest}")

    # 先核验第一步核验结果
    ws = wb["核验结果"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    i_item = header_index(header, "检查项目")
    i_ok = header_index(header, "是否通过")
    failed = [r[i_item] for r in rows[1:] if r and r[i_ok] != "通过"]
    if failed:
        raise RuntimeError(f"第一步核验未全部通过，拒绝进入第二步：{failed}")
    first_checks = []
    i_n = header_index(header, "检查数量")
    i_err = header_index(header, "最大误差")
    i_note = header_index(header, "说明")
    for r in rows[1:]:
        if not r or r[i_item] is None:
            continue
        first_checks.append({"检查项目": r[i_item], "检查数量": r[i_n],
                             "最大误差": r[i_err], "是否通过": r[i_ok], "说明": r[i_note]})

    def read_panel(sheet, nkeys):
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        header = next(it)
        years = [int(float(h)) for h in header if is_year_token(h)]
        data = {}
        for row in it:
            if row is None or row[0] is None:
                continue
            key = tuple(str(row[i]).strip() for i in range(nkeys))
            data[key] = {y: to_number(row[nkeys + 1 + i]) for i, y in enumerate(years)}
        return data, years, header

    veh, v_years, _ = read_panel("车辆退役铜_面板", 4)
    pow_, p_years, _ = read_panel("绿电退役铜_面板", 4)
    # 水电规模权重（参数审计表；第一步已核验）
    hydro_weights = {}
    if "水电规模权重_面板" in wb.sheetnames:
        ws = wb["水电规模权重_面板"]
        it = ws.iter_rows(values_only=True)
        header = next(it)
        wyears = [int(float(h)) for h in header if is_year_token(h)]
        for row in it:
            if row is None or row[0] is None:
                continue
            kind = str(row[0]).strip()
            scale = str(row[1]).strip()
            hydro_weights[(kind, scale)] = {
                y: to_number(row[3 + i]) for i, y in enumerate(wyears)}
    wb.close()
    return {
        "veh": veh, "pow": pow_, "v_years": v_years, "p_years": p_years,
        "path": str(path), "sha256": digest, "source": src_tag,
        "first_checks": first_checks, "hydro_weights": hydro_weights,
    }


def load_baseline_hydro(cfg):
    """读取修改前基线成果（若配置存在），用于记录 hydro 口径变化的差异。"""
    rel = cfg["paths"].get("step1_baseline")
    if not rel:
        return None
    path = Path(rel)
    if not path.is_absolute():
        path = PROJECT_DIR / path
    if not path.exists():
        return None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["绿电退役铜_面板"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    years = [int(float(x)) for x in header if is_year_token(x)]
    idx21 = [i for i, y in enumerate(years) if y >= 2021]
    i50 = years.index(2050) if 2050 in years else None
    out = {}
    for row in it:
        if row is None or row[0] is None:
            continue
        subtech = str(row[3]).strip()
        if subtech != "hydro":
            continue
        scn = norm_scenario(row[0]) if not str(row[0]).startswith("SSP") else str(row[0]).strip()
        key = str(row[0]).strip()
        cum = sum((row[5 + i] or 0.0) for i in idx21)
        v50 = (row[5 + i50] or 0.0) if i50 is not None else None
        item = out.setdefault(key, {"cum": 0.0, "v2050": 0.0})
        item["cum"] += cum
        item["v2050"] += (v50 or 0.0)
    wb.close()
    return {"path": str(path), "sha256": sha256_of(path), "by_scenario": out,
            "years": years}


# --------------------------------------------------------------------------- #
# 读取回收率与贸易矩阵
# --------------------------------------------------------------------------- #
def load_recycling(cfg):
    wb, used, src_tag, digest = open_workbook_readonly(
        cfg["paths"]["recycling_xlsx"], Path(cfg["paths"]["fallback_dir"]))
    ws = wb[cfg["sheets"]["recycling_rate"]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    hdr_idx = None
    for idx, row in enumerate(rows):
        if row and str(row[0]).strip() == "数据范围":
            hdr_idx = idx
            break
    if hdr_idx is None:
        raise ValueError("逐年回收率：未找到表头行（列A='数据范围'）")
    header = rows[hdr_idx]
    log(f"[输入] 回收率表头（Excel第{hdr_idx + 1}行）：{header[:10]}")
    col = cfg["recycling"]["param_columns"]
    need = ["data_range", "region", "device", "waste", "year", "collection",
            "processing", "combined_base", "combined_policy", "combined_tech"]
    for key in need:
        if col[key] >= len(header):
            raise ValueError(f"回收率表列数不足：{key}")

    data_range_filter = cfg["recycling"]["data_range_filter"]
    excluded = set(cfg["recycling"]["excluded_data_ranges"])
    rates = {}
    skipped_ranges = Counter()
    for row in rows[hdr_idx + 1:]:
        if row is None or row[col["data_range"]] is None:
            continue
        rng = str(row[col["data_range"]]).strip()
        if rng != data_range_filter:
            skipped_ranges[rng] += 1
            continue
        region = str(row[col["region"]]).strip()
        device = str(row[col["device"]]).strip()
        year = int(float(row[col["year"]]))
        rates[(region, device, year)] = {
            "C": to_number(row[col["collection"]]),
            "P": to_number(row[col["processing"]]),
            "H": to_number(row[col["combined_base"]]),
            "I": to_number(row[col["combined_policy"]]),
            "J": to_number(row[col["combined_tech"]]),
            "waste": str(row[col["waste"]]).strip(),
        }
    return {"rates": rates, "skipped": skipped_ranges, "path": cfg["paths"]["recycling_xlsx"],
            "sha256": digest, "source": src_tag, "n_rows": len(rates)}


def load_trade(cfg):
    wb, used, src_tag, digest = open_workbook_readonly(
        cfg["paths"]["trade_xlsx"], Path(cfg["paths"]["fallback_dir"]))
    ws = wb[cfg["sheets"]["trade_matrix"]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = [str(x).strip() for x in rows[0]]
    dest = header[1:]
    src = []
    M = np.zeros((len(dest), len(dest)), dtype=float)
    for i, row in enumerate(rows[1:]):
        src.append(str(row[0]).strip())
        for j in range(len(dest)):
            M[i, j] = to_number(row[j + 1]) or 0.0
    if src != dest:
        raise ValueError("贸易矩阵行列区域名顺序不一致")
    return {"matrix": M, "regions": dest, "path": cfg["paths"]["trade_xlsx"],
            "sha256": digest, "source": src_tag}


# --------------------------------------------------------------------------- #
# 计算
# --------------------------------------------------------------------------- #
def build_r3(cfg, rates, step1):
    """把回收率整理为 numpy 数组：C[rec][(region, device)] -> array(年份)。"""
    years = list(range(cfg["years"]["step2_start"], cfg["years"]["step2_end"] + 1))
    devices = cfg["recycling"]["device_categories"]
    regions = cfg["_panel_regions"]
    C = {rec: {} for rec in REC_ORDER}
    P = {rec: {} for rec in REC_ORDER}
    combined = {rec: {} for rec in REC_ORDER}
    c0, p0 = {}, {}
    missing = []
    for region in regions:
        canon = cfg["region_aliases"].get(region, {}).get("canonical", region)
        for device in devices:
            base = rates.get((canon, device, cfg["years"]["recycling_calibration_year"]))
            if base is None:
                missing.append((canon, device))
                continue
            c_base = base["C"]
            p_base = base["P"]
            c0[(region, device)] = c_base
            p0[(region, device)] = p_base
            arr_c = np.zeros(len(years))
            arr_p = np.zeros(len(years))
            arr_h = {"基准": np.zeros(len(years)), "政策推动": np.zeros(len(years)),
                     "技术进步": np.zeros(len(years))}
            for i, y in enumerate(years):
                row = rates.get((canon, device, y))
                if row is None:
                    raise ValueError(f"回收率缺年份：{canon}/{device}/{y}")
                arr_c[i] = row["C"]
                arr_p[i] = row["P"]
                arr_h["基准"][i] = row["H"]
                arr_h["政策推动"][i] = row["I"]
                arr_h["技术进步"][i] = row["J"]
            # 三种回收情景的收集率/处理率
            C["基准"][(region, device)] = np.full(len(years), c_base)
            P["基准"][(region, device)] = np.full(len(years), p_base)
            C["政策推动"][(region, device)] = arr_c.copy()
            P["政策推动"][(region, device)] = np.full(len(years), p_base)
            C["技术进步"][(region, device)] = np.full(len(years), c_base)
            P["技术进步"][(region, device)] = arr_p.copy()
            for rec in REC_ORDER:
                combined[rec][(region, device)] = C[rec][(region, device)] * P[rec][(region, device)]
    return {"years": years, "C": C, "P": P, "combined": combined,
            "c0": c0, "p0": p0, "missing": missing,
            "reported": {"基准": {"C": None, "P": None},
                         "政策推动": {"C": "F列逐年", "P": "P0"},
                         "技术进步": {"C": "C0", "P": "G列逐年"}}}


def build_device_weights(cfg, step1, years_index):
    """把第一步退役量整理为按设备类别的 W[scn][device] -> (32, 年份)。"""
    devices = cfg["recycling"]["device_categories"]
    sub_map = cfg["recycling"]["subtech_to_device"]
    scn_list = cfg["scenarios"]["canonical"]
    regions = cfg["_panel_regions"]
    r_pos = {r: i for i, r in enumerate(regions)}
    n_t = len(years_index)

    W = {}          # (scn, dept, detail) -> (32, n_t) 且 detail 为原技术组合
    W_dev = {}      # (scn, device) -> (32, n_t)
    coverage = defaultdict(int)
    for scn in scn_list:
        for device in devices:
            W_dev[(scn, device)] = np.zeros((len(regions), n_t))

    def add(dept, detail, device, scn, region, value_by_year):
        key = (scn, dept, detail)
        mat = W.setdefault(key, np.zeros((len(regions), n_t)))
        pos = r_pos[region]
        for i, y in enumerate(years_index):
            v = value_by_year.get(y)
            if v is None:
                continue
            mat[pos, i] += v
            W_dev[(scn, device)][pos, i] += v
        coverage[(scn, device)] += 1

    for key, series in step1["veh"].items():
        scn_raw, region, subsector, tech = key
        scn = norm_scenario(scn_raw)
        add("车辆", f"{subsector}|{tech}", "新能源汽车", scn, region, series)

    for key, series in step1["pow"].items():
        scn_raw, region, tech, subtech = key
        scn = norm_scenario(scn_raw)
        device = sub_map.get(subtech)
        if device is None:
            raise KeyError(f"Subtech 未映射到设备类别：{subtech}")
        add("绿电", subtech, device, scn, region, series)

    return W, W_dev, coverage


def compute_flows(cfg, step1, r3, trade_data, panel_regions):
    """主计算：返回各面板数据。"""
    years = list(range(cfg["years"]["step2_start"], cfg["years"]["step2_end"] + 1))
    n_t = len(years)
    devices = cfg["recycling"]["device_categories"]
    scn_list = cfg["scenarios"]["canonical"]
    matrix = trade_data["matrix"]
    m_regions = trade_data["regions"]
    m_pos = {r: i for i, r in enumerate(m_regions)}

    proc_canon = cfg["trade"]["processing_regions"]
    export_canon = cfg["trade"]["export_regions"]
    proc_pos = [m_pos[r] for r in proc_canon]
    T = matrix[:, proc_pos]                     # (32, 15)

    # 区域名：GCAM 原名 <-> 规范名
    canon_of = {}
    for region in panel_regions:
        canon = cfg["region_aliases"].get(region, {}).get("canonical", region)
        if canon not in m_pos:
            raise KeyError(f"区域无法映射到贸易矩阵：{region} -> {canon}")
        canon_of[region] = canon
    orig_of_canon = {v: k for k, v in canon_of.items()}

    W, W_dev, coverage = build_device_weights(cfg, step1, years)
    # 重新按规范名排列（与贸易矩阵顺序一致）；W 矩阵行序为 panel_regions 顺序
    panel_pos = {r: i for i, r in enumerate(panel_regions)}

    def to_matrix_order(mat):
        out = np.zeros((len(m_regions), n_t))
        for orig, canon in canon_of.items():
            out[m_pos[canon]] = mat[panel_pos[orig]]
        return out

    Wd_m = {(scn, dept, detail): to_matrix_order(mat) for (scn, dept, detail), mat in W.items()}
    Wdev_m = {(scn, device): to_matrix_order(W_dev[(scn, device)])
              for scn in scn_list for device in devices}

    details = {}
    # 规范顺序
    subs = cfg["devices"]["vehicle_subsectors"]
    techs = cfg["devices"]["vehicle_technologies"]
    details["车辆"] = [f"{s}|{t}" for s in subs for t in techs]
    details["绿电"] = cfg["devices"]["power_subtech_order"]

    detail_device = {}
    for d in details["车辆"]:
        detail_device[("车辆", d)] = "新能源汽车"
    for d in details["绿电"]:
        detail_device[("绿电", d)] = cfg["recycling"]["subtech_to_device"][d]

    proc_panels = {}   # (scn, rec, trade, dept, detail) -> (15, n_t)
    src_panels = {}    # (scn, rec, trade, dept, detail) -> (32, n_t)
    proc_dev = {}      # (scn, rec, trade, device) -> (15, n_t)
    src_dev = {}
    bilat_x = {}       # (scn, rec, trade, device) -> (32,15,n_t)
    bilat_q = {}
    totals = {}        # (scn, rec, trade, device) -> (n_t,) 处理地产量合计
    collected = {}     # (scn, rec, trade) -> (n_t,) 已收集量
    imported = {}      # (scn, rec, trade) -> dict

    C = r3["C"]
    P = r3["P"]

    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                Cmat = np.zeros((len(m_regions), n_t))
                Pmat = np.zeros((len(m_regions), n_t))
                for device in devices:
                    for orig, canon in canon_of.items():
                        Cmat[m_pos[canon]] = C[rs][(orig, device)]
                        Pmat[m_pos[canon]] = P[rs][(orig, device)]
                    # 设备层面双边
                    Wm = Wdev_m[(scn, device)]
                    if ts == "不开放":
                        Xin = Wm[proc_pos] * Cmat[proc_pos]          # 入厂含铜（本地收集）
                        Q = Xin * Pmat[proc_pos]                    # 次生铜产量
                        Xfull = np.zeros((len(m_regions), 15, n_t))
                        Qfull = np.zeros_like(Xfull)
                        Xfull[proc_pos, np.arange(15), :] = Xin
                        Qfull[proc_pos, np.arange(15), :] = Q
                    else:
                        Xfull = (Wm * Cmat)[:, None, :] * T[:, :, None]
                        Qfull = Xfull * Pmat[None, proc_pos, :]
                    bilat_x[(scn, rs, ts, device)] = Xfull
                    bilat_q[(scn, rs, ts, device)] = Qfull
                    proc_dev[(scn, rs, ts, device)] = Qfull.sum(axis=0)
                    src_dev[(scn, rs, ts, device)] = Qfull.sum(axis=1)
                    totals[(scn, rs, ts, device)] = proc_dev[(scn, rs, ts, device)].sum(axis=0)
                # 细节层面
                for dept in DEPTS:
                    for d in details[dept]:
                        device = detail_device[(dept, d)]
                        Wd = Wd_m.get((scn, dept, d))
                        if Wd is None:
                            Wd = np.zeros((len(m_regions), n_t))
                        Cmat_d = np.zeros((len(m_regions), n_t))
                        Pmat_d = np.zeros((len(m_regions), n_t))
                        for orig, canon in canon_of.items():
                            Cmat_d[m_pos[canon]] = C[rs][(orig, device)]
                            Pmat_d[m_pos[canon]] = P[rs][(orig, device)]
                        if ts == "不开放":
                            Xin_d = Wd[proc_pos] * Cmat_d[proc_pos]
                            Qd_local = Xin_d * Pmat_d[proc_pos]
                            proc_panels[(scn, rs, ts, dept, d)] = Qd_local
                            src = np.zeros((len(m_regions), n_t))
                            src[proc_pos] = Qd_local
                            src_panels[(scn, rs, ts, dept, d)] = src
                        else:
                            Xd = (Wd * Cmat_d)[:, None, :] * T[:, :, None]
                            Qd = Xd * Pmat_d[None, proc_pos, :]
                            proc_panels[(scn, rs, ts, dept, d)] = Qd.sum(axis=0)
                            src_panels[(scn, rs, ts, dept, d)] = Qd.sum(axis=1)

    return {
        "years": years, "n_t": n_t, "devices": devices, "details": details,
        "detail_device": detail_device, "proc_panels": proc_panels, "src_panels": src_panels,
        "proc_dev": proc_dev, "src_dev": src_dev, "totals": totals,
        "bilat_x": bilat_x, "bilat_q": bilat_q,
        "proc_canon": proc_canon, "export_canon": export_canon,
        "proc_pos": proc_pos, "canon_of": canon_of, "orig_of_canon": orig_of_canon,
        "m_regions": m_regions, "matrix": matrix, "T": T,
        "Wdev_m": Wdev_m, "Wd_m": Wd_m, "coverage": coverage,
        "panel_regions": panel_regions,
    }


# --------------------------------------------------------------------------- #
# 核验
# --------------------------------------------------------------------------- #
def verify_all(cfg, rec, trade_data, flow, r3, step1):
    checks = []

    def add(item, n, err, ok, note):
        checks.append({"检查项目": item, "检查数量": n, "最大误差": err,
                       "是否通过": "通过" if ok else "未通过", "说明": note})

    devices = flow["devices"]
    years = flow["years"]
    scn_list = cfg["scenarios"]["canonical"]
    M = trade_data["matrix"]
    n = M.shape[0]
    m_pos_of = {r: i for i, r in enumerate(flow["m_regions"])}

    # 1 矩阵非负、行和
    row_sum = M.sum(axis=1)
    err = float(np.max(np.abs(row_sum - 1.0)))
    add("贸易矩阵非负且各来源区行和为1", n, err,
        bool(np.all(M >= 0) and err <= 1e-9),
        f"最小值={M.min():.3e}；行和范围=[{row_sum.min():.10f}, {row_sum.max():.10f}]")

    # 2 路线数量
    nz = int(np.count_nonzero(M > 0))
    diag = int(sum(1 for i in range(n) if M[i, i] > 0))
    add("贸易矩阵非零路线数（15本地+255跨区）", nz, None, nz == 270 and diag == 15,
        f"非零={nz}，对角线={diag}，跨区={nz - diag}")

    # 3 已收集量 = 分配至处理区总量
    max_rel = 0.0
    n_chk = 0
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                for t in range(flow["n_t"]):
                    collected = 0.0
                    allocated = 0.0
                    for device in devices:
                        X = flow["bilat_x"][(scn, rs, ts, device)][:, :, t]
                        allocated += float(X.sum())
                        Wm = flow["Wdev_m"][(scn, device)][:, t]
                        col = np.zeros(n)
                        for orig, canon in flow["canon_of"].items():
                            col[m_pos_of[canon]] = r3["C"][rs][(orig, device)][t]
                        if ts == "不开放":
                            collected += float((Wm * col)[flow["proc_pos"]].sum())
                        else:
                            collected += float((Wm * col).sum())
                    n_chk += 1
                    denom = max(1e-9, abs(collected))
                    max_rel = max(max_rel, abs(allocated - collected) / denom)
    add("已收集量=分配至处理区总量", n_chk, max_rel, max_rel <= 1e-9,
        "逐能源情景×回收情景×贸易情景×年份检查（求和维度不同）")

    # 4 跨区进口 = 出口（离对角元素两种聚合路径）
    max_rel2 = 0.0
    n_chk2 = 0
    for scn in scn_list:
        for rs in REC_ORDER:
            for t in range(flow["n_t"]):
                imports = 0.0
                exports = 0.0
                for device in devices:
                    X = flow["bilat_x"][(scn, rs, "全球开放", device)][:, :, t]
                    col_sums = X.sum(axis=0).copy()      # 按处理区
                    row_sums = X.sum(axis=1).copy()      # 按来源区
                    for k, pos in enumerate(flow["proc_pos"]):
                        col_sums[k] -= X[pos, k]         # 扣除本地路线
                        row_sums[pos] -= X[pos, k]
                    imports += float(col_sums.sum())
                    exports += float(row_sums.sum())
                n_chk2 += 1
                max_rel2 = max(max_rel2, abs(imports - exports) / max(1e-9, abs(imports)))
    add("全球跨区进口总量=出口总量", n_chk2, max_rel2, max_rel2 <= 1e-9,
        "对同一组跨区双边流分别按处理区汇总（进口）与按来源区汇总（出口），扣除15条本地路线")

    # 5 Q ≤ X
    max_diff = 0.0
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                for device in devices:
                    X = flow["bilat_x"][(scn, rs, ts, device)]
                    Q = flow["bilat_q"][(scn, rs, ts, device)]
                    max_diff = max(max_diff, float(np.max(Q - X)))
    add("各路线次生铜产量不超过入厂含铜量", len(scn_list) * 3 * 2 * len(devices), max_diff,
        max_diff <= 1e-9, "Q=X×P 且 P≤1")

    # 6 来源归属合计 = 处理地生产合计
    max_rel3 = 0.0
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                tot_proc = sum(float(flow["totals"][(scn, rs, ts, d)].sum()) for d in devices)
                tot_src = sum(float(flow["src_dev"][(scn, rs, ts, d)].sum()) for d in devices)
                max_rel3 = max(max_rel3, abs(tot_proc - tot_src) / max(1e-9, tot_proc))
    add("来源归属回收铜合计=处理地次生铜合计", len(scn_list) * 3 * 2, max_rel3,
        max_rel3 <= 1e-9, "两条聚合路径的数值一致性")

    # 7 不开放未纳入17区
    export_set = set(cfg["trade"]["export_regions"])
    proc_set = set(cfg["trade"]["processing_regions"])
    bad = 0
    n_rows_na = 0
    for (scn, rs, ts, dept, d), arr in flow["proc_panels"].items():
        if ts != "不开放":
            continue
        n_rows_na += 1
        if arr.shape[0] != 15:
            bad += 1
    # 来源地面板：不开放时非处理区的行必须全为0
    for (scn, rs, ts, dept, d), arr in flow["src_panels"].items():
        if ts != "不开放":
            continue
        for pos, canon in enumerate(flow["m_regions"]):
            if canon in proc_set:
                continue
            if arr[pos].sum() != 0:
                bad += 1
    add("不开放情景仅含15个有处理能力区域", n_rows_na, None, bad == 0,
        f"不开放处理地面板键数={n_rows_na}；17个全出口区（{len(export_set)}个）"
        "不作为处理区生成行，也不得在来源地面板出现非零值")

    # 8 参数范围
    bad_param = 0
    max_out = 0.0
    for rs in REC_ORDER:
        for key, arr in r3["C"][rs].items():
            max_out = max(max_out, float(np.max(arr - 1.0)), float(-np.min(arr)))
            if np.any(arr < -1e-12) or np.any(arr > 1 + 1e-12):
                bad_param += 1
        for key, arr in r3["P"][rs].items():
            max_out = max(max_out, float(np.max(arr - 1.0)), float(-np.min(arr)))
            if np.any(arr < -1e-12) or np.any(arr > 1 + 1e-12):
                bad_param += 1
    add("收集率与综合处理率在[0,1]", len(r3["C"]["基准"]) * 6, max_out, bad_param == 0,
        f"越界序列数={bad_param}")

    # 9 2021 基准一致
    max_d = 0.0
    for key in r3["C"]["基准"]:
        for rec_name in REC_ORDER:
            max_d = max(max_d, float(abs(r3["C"][rec_name][key][0] - r3["c0"][key])))
            max_d = max(max_d, float(abs(r3["P"][rec_name][key][0] - r3["p0"][key])))
    add("三种回收情景2021年使用相同基准参数", len(r3["C"]["基准"]), max_d, max_d <= 1e-12,
        "校准年 C=C0、P=P0，与基准情景一致")

    # 10 综合回收率组合校验
    max_c = 0.0
    n_c = 0
    for region in flow["panel_regions"]:
        canon = cfg["region_aliases"].get(region, {}).get("canonical", region)
        for device in devices:
            for rec_name, col in (("基准", "H"), ("政策推动", "I"), ("技术进步", "J")):
                arr = r3["combined"][rec_name][(region, device)]
                exp = np.array([rec["rates"][(canon, device, y)][col] for y in years])
                max_c = max(max_c, float(np.max(np.abs(arr - exp))))
                n_c += 1
    add("综合回收率=收集率×综合处理率（对照原表H/I/J列）", n_c, max_c, max_c <= 1e-12,
        "基准 H=C0×P0；政策推动 I=C(t)×P0；技术进步 J=C0×P(t)")

    # 11 设备类别投入产出守恒
    max_rel4 = 0.0
    n_chk4 = 0
    for scn in scn_list:
        for rs in REC_ORDER:
            for device in devices:
                X = flow["bilat_x"][(scn, rs, "全球开放", device)]
                Q = flow["bilat_q"][(scn, rs, "全球开放", device)]
                # 用目的地处理率还原：ΣQ = Σ_j (Σ_r X)×P_j
                Pj = np.zeros((len(flow["proc_pos"]), flow["n_t"]))
                for k, canon in enumerate(flow["proc_canon"]):
                    orig = flow["orig_of_canon"][canon]
                    Pj[k] = r3["P"][rs][(orig, device)]
                X_dest = X.sum(axis=0)
                expected = (X_dest * Pj).sum()
                got = Q.sum()
                n_chk4 += 1
                max_rel4 = max(max_rel4, abs(expected - got) / max(1e-9, abs(got)))
    add("各设备类别处理投入与产出守恒", n_chk4, max_rel4, max_rel4 <= 1e-9,
        "处理地层面 Q=入厂量×该处理地综合处理率")

    # 12 细节汇总 = 设备汇总
    max_rel5 = 0.0
    n_chk5 = 0
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                for device in devices:
                    tot_dev = flow["proc_dev"][(scn, rs, ts, device)]
                    tot_det = np.zeros_like(tot_dev)
                    for dept in DEPTS:
                        for d in flow["details"][dept]:
                            if flow["detail_device"][(dept, d)] != device:
                                continue
                            tot_det += flow["proc_panels"][(scn, rs, ts, dept, d)]
                    n_chk5 += 1
                    denom = max(1e-9, float(tot_dev.sum()))
                    max_rel5 = max(max_rel5, float(np.abs(tot_det - tot_dev).sum()) / denom)
    add("技术明细汇总=设备类别汇总", n_chk5, max_rel5, max_rel5 <= 1e-9,
        "处理地面板按设备类别加总应与明细加总一致")

    # 12b 水电单一子技术、无规模明细行、无重复计量
    hydro_keys = [k for k in step1["pow"] if k[3] == "hydro"]
    scale_rows = [k for k in step1["pow"] if str(k[3]).startswith("hydro_")]
    hydro_dev_total = sum(float(flow["src_dev"][(scn, rs, ts, "水电")].sum())
                          for scn in scn_list for rs in REC_ORDER for ts in TRADE_ORDER)
    hydro_detail_total = 0.0
    n_hydro_detail = 0
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                arr = flow["src_panels"].get((scn, rs, ts, "绿电", "hydro"))
                if arr is not None:
                    hydro_detail_total += float(arr.sum())
                    n_hydro_detail += 1
    rel_hydro = abs(hydro_dev_total - hydro_detail_total) / max(1e-9, abs(hydro_dev_total))
    add("水电仅以单一hydro进入汇总且无重复计量", len(hydro_keys), rel_hydro,
        len(scale_rows) == 0 and rel_hydro <= 1e-12,
        f"第一步hydro序列={len(hydro_keys)}；规模明细行={len(scale_rows)}；"
        f"来源归属面板hydro明细键={n_hydro_detail}（3情景×3回收×2贸易）；"
        f"设备类别合计与hydro明细合计相对差={rel_hydro:.2e}")

    # 12c 第二步输入 = 第一步退役铜（逐键逐年）
    sub_map = cfg["recycling"]["subtech_to_device"]
    max_rel_in = 0.0
    n_in = 0
    yset = set(years)
    for scn in scn_list:
        for device in devices:
            exp = 0.0
            for key, series in step1["pow"].items():
                if key[0] != scn or sub_map.get(key[3]) != device:
                    continue
                exp += sum((series.get(y) or 0.0) for y in yset)
            for key, series in step1["veh"].items():
                if key[0] != scn or device != "新能源汽车":
                    continue
                exp += sum((series.get(y) or 0.0) for y in yset)
            got = float(flow["Wdev_m"][(scn, device)].sum())
            n_in += 1
            max_rel_in = max(max_rel_in, abs(got - exp) / max(1e-9, abs(exp)))
    add("第二步输入=第一步退役铜（逐键逐年）", n_in, max_rel_in, max_rel_in <= 1e-9,
        "按设备类别对第一步明细逐年求和，与第二步使用的 W 矩阵比较（2021—2050）")

    # 13 六组合完整性
    expected = len(scn_list) * len(REC_ORDER) * len(TRADE_ORDER)
    have = {(k[0], k[1], k[2]) for k in flow["proc_dev"]}
    add("六情景组合完整性", len(have), None, len(have) == expected,
        "能源路径3×回收3×贸易2=18个组合；实际=" + str(len(have)))

    # 14 各能源情景独立
    diffs = []
    for rs in REC_ORDER:
        for ts in TRADE_ORDER:
            a = sum(float(flow["totals"][(scn_list[0], rs, ts, d)].sum()) for d in devices)
            b = sum(float(flow["totals"][(scn_list[1], rs, ts, d)].sum()) for d in devices)
            c = sum(float(flow["totals"][(scn_list[2], rs, ts, d)].sum()) for d in devices)
            diffs.append((a, b, c))
    add("各能源情景独立计算", len(diffs), None,
        all(len(set(round(x, 6) for x in t)) == 3 for t in diffs),
        "同回收/贸易条件下三条能源路径的合计产量互不相同（未串用）")

    # 15 开放 ≥ 不开放（同一回收情景）
    n_neg = 0
    max_drop = 0.0
    for scn in scn_list:
        for rs in REC_ORDER:
            open_tot = sum(float(flow["totals"][(scn, rs, "全球开放", d)].sum()) for d in devices)
            closed_tot = sum(float(flow["totals"][(scn, rs, "不开放", d)].sum()) for d in devices)
            diff = open_tot - closed_tot
            if diff < -1e-6:
                n_neg += 1
                max_drop = min(max_drop, diff)
    add("开放情景全球产量不低于不开放（不要求相等）", len(scn_list) * 3, max_drop, n_neg == 0,
        "开放仅增加17区已收集铜的境外处理路径，不改变来源区退役量与收集量")

    # 16 SSP5 结构性零
    ssp5_ccs = float(sum(flow["totals"][("SSP5-8.5", rs, ts, "CCS相关设备")].sum()
                         for rs in REC_ORDER for ts in TRADE_ORDER))
    add("SSP5-8.5 无CCS技术的结构性零", 1, ssp5_ccs, True,
        "源表 power generation 的 SSP5-8.5 仅含8种原技术、不含六种CCS，"
        "故该路径CCS相关设备退役与次生铜为结构性0（非缺失值填充）")

    return checks


# --------------------------------------------------------------------------- #
# 输出 Excel
# --------------------------------------------------------------------------- #
def write_wide(workbook, name, index_fields, years, rows_iter, widths,
               freeze_cols=None, pct_fields=()):
    ws = workbook.add_worksheet(name)
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
    num_fmt = workbook.add_format({"num_format": "#,##0.00", "valign": "top"})
    pct_fmt = workbook.add_format({"num_format": "0.00%", "valign": "top"})
    text_fmt = workbook.add_format({"valign": "top"})

    ws.freeze_panes(1, freeze_cols if freeze_cols is not None else len(index_fields))
    ws.set_row(0, 28)
    for c, f in enumerate(index_fields):
        ws.write_string(0, c, str(f), header_fmt)
        ws.set_column(c, c, widths[c] if c < len(widths) else 16)
    for i, y in enumerate(years):
        c = len(index_fields) + i
        ws.write_number(0, c, y, header_fmt)
        ws.set_column(c, c, 12)

    nrow = 0
    for keys, values in rows_iter:
        nrow += 1
        use_pct = any(k in pct_fields for k in keys)
        for c, v in enumerate(keys):
            if v is None:
                continue
            if isinstance(v, str):
                ws.write_string(nrow, c, v, text_fmt)
            else:
                ws.write_number(nrow, c, float(v), text_fmt)
        for i, v in enumerate(values):
            if v is None:
                continue
            ws.write_number(nrow, len(keys) + i, float(v), pct_fmt if use_pct else num_fmt)
    if nrow:
        ws.autofilter(0, 0, nrow, len(index_fields) + len(years) - 1)
    return ws, nrow


def build_workbook(cfg, rec, trade_data, flow, r3, step1, checks):
    out_path = Path(cfg["paths"]["step2_out"])
    tmp_dir = Path(tempfile.mkdtemp(prefix="step2_"))
    tmp = tmp_dir / out_path.name
    wb = xlsxwriter.Workbook(str(tmp), {"nan_inf_to_errors": True})
    years = flow["years"]
    devices = flow["devices"]
    scn_list = cfg["scenarios"]["canonical"]
    proc_orig = [flow["orig_of_canon"][c] for c in flow["proc_canon"]]
    bilat_count = {"全球开放": 0, "不开放": 0}

    # 1) 车辆次生铜_处理地
    def veh_proc_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    for j, region in enumerate(proc_orig):
                        for d in flow["details"]["车辆"]:
                            subsector, tech = d.split("|")
                            arr = flow["proc_panels"][(scn, rs, ts, "车辆", d)][j]
                            yield [scn, rs, ts, region, subsector, tech, "t Cu"], arr
    _, n1 = write_wide(wb, "车辆次生铜_处理地",
                       ["Scenario", "回收情景", "贸易情景", "Region", "Subsector", "Technology", "单位"],
                       years, veh_proc_rows(), [11, 10, 10, 24, 20, 11, 7])

    # 2) 绿电次生铜_处理地
    def pow_proc_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    for j, region in enumerate(proc_orig):
                        for d in flow["details"]["绿电"]:
                            arr = flow["proc_panels"][(scn, rs, ts, "绿电", d)][j]
                            yield [scn, rs, ts, region, cfg["devices"]["subtech_to_tech"][d], d,
                                   "t Cu"], arr
    _, n2 = write_wide(wb, "绿电次生铜_处理地",
                       ["Scenario", "回收情景", "贸易情景", "Region", "tech", "Subtech", "单位"],
                       years, pow_proc_rows(), [11, 10, 10, 24, 18, 24, 7])

    # 3) 车辆回收铜_来源地
    def veh_src_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    regions = proc_orig if ts == "不开放" else flow["panel_regions"]
                    for region in regions:
                        pos = flow["m_regions"].index(flow["canon_of"][region])
                        for d in flow["details"]["车辆"]:
                            subsector, tech = d.split("|")
                            arr = flow["src_panels"][(scn, rs, ts, "车辆", d)][pos]
                            yield [scn, rs, ts, region, subsector, tech, "t Cu"], arr
    _, n3 = write_wide(wb, "车辆回收铜_来源地",
                       ["Scenario", "回收情景", "贸易情景", "Region", "Subsector", "Technology", "单位"],
                       years, veh_src_rows(), [11, 10, 10, 26, 20, 11, 7])

    # 4) 绿电回收铜_来源地
    def pow_src_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    regions = proc_orig if ts == "不开放" else flow["panel_regions"]
                    for region in regions:
                        pos = flow["m_regions"].index(flow["canon_of"][region])
                        for d in flow["details"]["绿电"]:
                            arr = flow["src_panels"][(scn, rs, ts, "绿电", d)][pos]
                            yield [scn, rs, ts, region, cfg["devices"]["subtech_to_tech"][d], d,
                                   "t Cu"], arr
    _, n4 = write_wide(wb, "绿电回收铜_来源地",
                       ["Scenario", "回收情景", "贸易情景", "Region", "tech", "Subtech", "单位"],
                       years, pow_src_rows(), [11, 10, 10, 26, 18, 24, 7])

    # 5) 全球六组合_面板
    def global_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    total = np.zeros(flow["n_t"])
                    for device in devices:
                        arr = flow["totals"][(scn, rs, ts, device)]
                        total += arr
                        yield [scn, rs, ts, device, "t Cu"], arr
                    yield [scn, rs, ts, "合计", "t Cu"], total
    ws5, n5 = write_wide(wb, "全球六组合_面板",
                         ["Scenario", "回收情景", "贸易情景", "设备类别", "单位"],
                         years, global_rows(), [11, 10, 10, 18, 7], freeze_cols=5)

    # 6) 情景差值_面板
    def diff_rows():
        for scn in scn_list:
            for ts in TRADE_ORDER:
                for label, rec_a, rec_b in (("政策推动-基准", "政策推动", "基准"),
                                            ("技术进步-基准", "技术进步", "基准")):
                    for device in devices + ["合计"]:
                        a = (sum((flow["totals"][(scn, rec_a, ts, d)] for d in devices),
                                 np.zeros(flow["n_t"])) if device == "合计"
                             else flow["totals"][(scn, rec_a, ts, device)])
                        b = (sum((flow["totals"][(scn, rec_b, ts, d)] for d in devices),
                                 np.zeros(flow["n_t"])) if device == "合计"
                             else flow["totals"][(scn, rec_b, ts, device)])
                        yield [scn, label, label, ts, device, "t Cu"], a - b
            for rs in REC_ORDER:
                for device in devices + ["合计"]:
                    o = (sum((flow["totals"][(scn, rs, "全球开放", d)] for d in devices),
                             np.zeros(flow["n_t"])) if device == "合计"
                         else flow["totals"][(scn, rs, "全球开放", device)])
                    c = (sum((flow["totals"][(scn, rs, "不开放", d)] for d in devices),
                             np.zeros(flow["n_t"])) if device == "合计"
                         else flow["totals"][(scn, rs, "不开放", device)])
                    yield [scn, "开放-不开放", rs, "全球开放-不开放", device, "t Cu"], o - c
    ws6, n6 = write_wide(wb, "情景差值_面板",
                         ["Scenario", "对比类型", "回收情景", "贸易情景", "设备类别", "单位"],
                         years, diff_rows(), [11, 16, 16, 18, 18, 7], freeze_cols=6)

    # 7) 双边废铜_面板
    def bilat_rows():
        for scn in scn_list:
            for rs in REC_ORDER:
                for ts in TRADE_ORDER:
                    for device in devices:
                        X = flow["bilat_x"][(scn, rs, ts, device)]
                        Q = flow["bilat_q"][(scn, rs, ts, device)]
                        pairs = []
                        if ts == "不开放":
                            for k, pos in enumerate(flow["proc_pos"]):
                                pairs.append((pos, k))
                        else:
                            # 只保留矩阵中真实非零的路线（15条本地 + 255条跨区）
                            for pos in range(len(flow["m_regions"])):
                                for k in range(len(flow["proc_canon"])):
                                    if flow["T"][pos, k] > 0:
                                        pairs.append((pos, k))
                        for pos, k in pairs:
                            r_orig = flow["orig_of_canon"][flow["m_regions"][pos]]
                            j_orig = flow["orig_of_canon"][flow["proc_canon"][k]]
                            xa = X[pos, k]
                            qa = Q[pos, k]
                            bilat_count[ts] += 1
                            yield [scn, rs, ts, r_orig, j_orig, device, "入厂含铜量", "t Cu"], xa
                            yield [scn, rs, ts, r_orig, j_orig, device, "次生铜产量", "t Cu"], qa
    ws7, n7 = write_wide(wb, "双边废铜_面板",
                         ["Scenario", "回收情景", "贸易情景", "来源区", "处理区", "设备类别",
                          "流量类型", "单位"],
                         years, bilat_rows(), [11, 10, 10, 26, 26, 18, 14, 7], freeze_cols=8)

    # 8) 回收参数_面板
    def param_rows():
        for region in flow["panel_regions"]:
            for device in devices:
                for rs in REC_ORDER:
                    yield ([region, device, rs, "收集率", "%"], r3["C"][rs][(region, device)])
                    yield ([region, device, rs, "综合处理率", "%"], r3["P"][rs][(region, device)])
                    yield ([region, device, rs, "综合回收率", "%"], r3["combined"][rs][(region, device)])
    ws8, n8 = write_wide(wb, "回收参数_面板",
                         ["Region", "设备类别", "回收情景", "参数类型", "单位"],
                         years, param_rows(), [26, 18, 10, 14, 7], freeze_cols=5,
                         pct_fields=("收集率", "综合处理率", "综合回收率"))

    # 9) 贸易矩阵
    ws9 = wb.add_worksheet("贸易矩阵")
    hf = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                        "align": "center", "valign": "vcenter", "text_wrap": True})
    nf = wb.add_format({"num_format": "0.00%"})
    ws9.freeze_panes(1, 1)
    ws9.write_string(0, 0, "来源区\\处理区", hf)
    ws9.set_column(0, 0, 26)
    for j, name in enumerate(trade_data["regions"]):
        ws9.write_string(0, j + 1, name, hf)
        ws9.set_column(j + 1, j + 1, 13)
    ws9.write_string(0, len(trade_data["regions"]) + 1, "行和", hf)
    ws9.write_string(0, len(trade_data["regions"]) + 2, "处理能力", hf)
    ws9.set_column(len(trade_data["regions"]) + 1, len(trade_data["regions"]) + 2, 14)
    M = trade_data["matrix"]
    for i, name in enumerate(trade_data["regions"]):
        ws9.write_string(i + 1, 0, name)
        for j in range(M.shape[1]):
            if M[i, j]:
                ws9.write_number(i + 1, j + 1, float(M[i, j]), nf)
            else:
                ws9.write_number(i + 1, j + 1, 0.0, nf)
        ws9.write_number(i + 1, M.shape[1] + 1, float(M[i].sum()))
        ws9.write_string(i + 1, M.shape[1] + 2,
                         "有处理能力" if trade_data["regions"][i] in cfg["trade"]["processing_regions"]
                         else "无本地处理能力")
    ws9.autofilter(0, 0, M.shape[0], M.shape[1] + 2)

    # 10) 区域映射
    ws10 = wb.add_worksheet("区域映射")
    ws10.freeze_panes(1, 1)
    fields10 = ["原名", "规范名称", "来源文件", "映射依据"]
    for c, f in enumerate(fields10):
        ws10.write_string(0, c, f, hf)
        ws10.set_column(c, c, [26, 28, 60, 60][c])
    for i, region in enumerate(flow["panel_regions"], start=1):
        canon = flow["canon_of"][region]
        alias = cfg["region_aliases"].get(region)
        ws10.write_string(i, 0, region)
        ws10.write_string(i, 1, canon)
        ws10.write_string(i, 2, "；".join([Path(cfg["paths"]["recycling_xlsx"]).name,
                                          Path(cfg["paths"]["trade_xlsx"]).name]))
        ws10.write_string(i, 3, alias["basis"] if alias else "名称一致，无需别名")
    ws10.autofilter(0, 0, len(flow["panel_regions"]), 3)

    # 11) 核验结果
    ws11 = wb.add_worksheet("核验结果")
    ws11.freeze_panes(1, 2)
    fields11 = ["检查项目", "检查数量", "最大误差", "是否通过", "说明"]
    for c, f in enumerate(fields11):
        ws11.write_string(0, c, f, hf)
    for c, w in enumerate([46, 12, 18, 10, 96]):
        ws11.set_column(c, c, w)
    for i, chk in enumerate(checks, start=1):
        ws11.write_string(i, 0, str(chk["检查项目"]))
        if chk["检查数量"] is not None:
            ws11.write_number(i, 1, float(chk["检查数量"]))
        if chk["最大误差"] is not None:
            ws11.write_number(i, 2, float(chk["最大误差"]))
        ws11.write_string(i, 3, str(chk["是否通过"]))
        ws11.write_string(i, 4, str(chk["说明"]))
    ws11.autofilter(0, 0, len(checks), 4)

    # 12) 方法说明
    ws12 = wb.add_worksheet("方法说明")
    ws12.freeze_panes(1, 1)
    ws12.write_string(0, 0, "项目", hf)
    ws12.write_string(0, 1, "内容", hf)
    ws12.set_column(0, 0, 26)
    ws12.set_column(1, 1, 150)
    for i, (k, v) in enumerate(build_method_rows(cfg, flow, r3, step1), start=1):
        ws12.write_string(i, 0, str(k))
        ws12.write_string(i, 1, str(v))
    ws12.autofilter(0, 0, len(build_method_rows(cfg, flow, r3, step1)), 1)

    wb.close()
    counts = {"车辆次生铜_处理地": n1, "绿电次生铜_处理地": n2, "车辆回收铜_来源地": n3,
              "绿电回收铜_来源地": n4, "全球六组合_面板": n5, "情景差值_面板": n6,
              "双边废铜_面板": n7, "回收参数_面板": n8}
    return tmp, counts, bilat_count


def build_method_rows(cfg, flow, r3, step1):
    a = cfg["assumptions"]
    years = flow["years"]
    rows = [
        ("脚本", "step2_secondary_copper.py（第二步；import 不自动计算）"),
        ("输入", f"第一步成果 {step1['path']}（sha256={step1['sha256']}）；"
                 f"回收率 {cfg['paths']['recycling_xlsx']}；贸易矩阵 {cfg['paths']['trade_xlsx']}"),
        ("第一步核验", "读取第一步「核验结果」工作表，全部通过后才进入第二步"),
        ("第一步口径", "设备量驱动：新增=Q_t−Q_(t−1)+R_t；退役=自然退役+经济性提前退役；"
                  "铜强度按投入年固定；第一步输出四类结果（新增/退役设备量、新增/退役铜量）"),
        ("不开放公式", "Q[s,r,k,t]=W[r,k,t]×C[s,r,k,t]×P[s,r,k,t]，仅15个有处理能力区域；"
                    "17区不生成结果行，不建暂存/未利用/跨年库存账户"),
        ("全球开放公式", "X[s,r,j,k,t]=W[r,k,t]×C[s,r,k,t]×T[r,j]；Q[s,r,j,k,t]=X[s,r,j,k,t]×P[s,j,k,t]"),
        ("收集率归属", "取废铜产生地（来源区r）的收集率"),
        ("处理率归属", "取出口目的地/实际处理地j的综合处理率"),
        ("政策推动×开放", "来源区提升后的收集率C(t) × 处理区基准处理率P0"),
        ("技术进步×开放", "来源区基准收集率C0 × 处理区提升后的处理率P(t)"),
        ("综合回收率", "本地处理时可直接用W×综合回收率(H/I/J)；本脚本按C×P计算并核对与H/I/J一致，不再重复乘C或P，"
                   "也不另加拆解、分选、冶炼效率"),
        ("处理能力", a["no_capacity_limit"]),
        ("贸易矩阵", a["trade_matrix_scope"] + cfg["trade"]["static_matrix"]),
        ("区域别名", "China→Chinese Mainland；Taiwan→Chinese Taiwan；"
                 "European Free Trade Association→European_Free_Trade；"
                 "Central America and Caribbean→C. America and Caribbean；其余同名（见「区域映射」）"),
        ("设备回收类别", "BEV/PHEV车辆→新能源汽车/ELV；PV→光伏/WEEE；wind→陆上风电/IEW；"
                    "wind_offshore→海上风电/IEW；hydro→水电/IEW；nuclear→核电/IEW；六种CCS→CCS相关设备/IEW"),
        ("水电口径", "第一步内部按CON规模结构（小型/中型/大型）分别计算，对外只输出一条hydro退役铜；"
                 "第二步直接读取合并后的hydro并按“水电→IEW”统一套用回收与贸易参数，规模明细不进入次生铜面板，"
                 "不存在规模行与合计行重复计量"),
        ("水电CON份额", "1990—2021 小20%/中16%/大64%；2022—2050 中16%、小按76.1148/93/100 GW对380.574 GW"
                   "分段线性、大取余；中国规模结构外推到全球32区（模型假设）"),
        ("SSP5-8.5 说明", "该路径源表不含六种CCS技术，CCS相关设备退役与次生铜为结构性0，非缺失值填充"),
        ("不开放覆盖范围", "汇总仅覆盖15个有处理能力区域；17区退役量不进入不开放情景合计"),
        ("开放增产机制", "开放后17区已收集铜可出口处理，因此全球有效处理产出可高于不开放；"
                   "但不能据此要求两者相等"),
        ("不计算项", "运输成本、铜价值分摊、再生铜返还来源地、库存、原生铜替代、水压力"),
        ("报告期", a["recycling_period"]),
        ("单位", "全部铜量统一 t Cu；参数为0—1比例（显示0.00%）"),
        ("双边表结构", "「双边废铜_面板」在需求列示的 Scenario/回收情景/来源区/处理区/设备类别/流量类型/单位 之外"
                  "增加「贸易情景」列：对角线本地路线在不开放与开放两种情景下都有取值，若无该列会出现重复键。"),
        ("输出", cfg["paths"]["step2_out"]),
    ]
    return rows


# --------------------------------------------------------------------------- #
# 写盘核对
# --------------------------------------------------------------------------- #
def verify_workbook(path, flow, cfg):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    expected = ["车辆次生铜_处理地", "绿电次生铜_处理地", "车辆回收铜_来源地", "绿电回收铜_来源地",
                "全球六组合_面板", "情景差值_面板", "双边废铜_面板", "回收参数_面板",
                "贸易矩阵", "区域映射", "核验结果", "方法说明"]
    out = []
    out.append({"检查项目": "工作表数量与顺序", "检查数量": len(expected),
                "最大误差": None, "是否通过": "通过" if wb.sheetnames == expected else "未通过",
                "说明": str(wb.sheetnames)})

    def check_sheet(sheet, nkeys, n_years, expected_rows, sample_keys, label):
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        header = next(it)
        hdr_ok = len([h for h in header if isinstance(h, int)]) == n_years
        n = 0
        bad = 0
        maxdiff = 0.0
        seen = set()
        samples = {}
        for row in it:
            if row is None or row[0] is None:
                continue
            n += 1
            key = tuple(row[:nkeys])
            if key in seen:
                bad += 1
            seen.add(key)
            if key in sample_keys:
                samples[key] = np.array([row[nkeys + 1 + i] for i in range(n_years)], dtype=float)
        for key, want in sample_keys.items():
            got = samples.get(key)
            if got is None:
                bad += 1
                continue
            maxdiff = max(maxdiff, float(np.max(np.abs(got - want))))
        out.append({"检查项目": f"写盘核对-{label}", "检查数量": n, "最大误差": maxdiff,
                    "是否通过": "通过" if (hdr_ok and n == expected_rows and bad == 0
                                         and maxdiff <= 1e-6) else "未通过",
                    "说明": f"行数={n}（期望{expected_rows}）；表头年份列={hdr_ok}；"
                            f"重复/缺样={bad}；抽查最大绝对差={maxdiff:.3e}"})

    scn = cfg["scenarios"]["canonical"][0]
    proc_orig = [flow["orig_of_canon"][c] for c in flow["proc_canon"]]
    veh_samples = {}
    for j in (proc_orig[0], proc_orig[-1]):
        jj = proc_orig.index(j)
        d = flow["details"]["车辆"][0]
        veh_samples[(scn, "基准", "不开放", j, d.split("|")[0], d.split("|")[1])] = \
            flow["proc_panels"][(scn, "基准", "不开放", "车辆", d)][jj]
    check_sheet("车辆次生铜_处理地", 6, flow["n_t"], 3 * 3 * 2 * 15 * 14, veh_samples, "车辆次生铜_处理地")

    pow_samples = {}
    for j in (proc_orig[0], proc_orig[-1]):
        jj = proc_orig.index(j)
        d = flow["details"]["绿电"][0]
        pow_samples[(scn, "基准", "全球开放", j, cfg["devices"]["subtech_to_tech"][d], d)] = \
            flow["proc_panels"][(scn, "基准", "全球开放", "绿电", d)][jj]
    check_sheet("绿电次生铜_处理地", 6, flow["n_t"], 3 * 3 * 2 * 15 * 11, pow_samples, "绿电次生铜_处理地")

    src_samples = {}
    region = flow["panel_regions"][0]
    pos = flow["m_regions"].index(flow["canon_of"][region])
    d = flow["details"]["车辆"][0]
    src_samples[(scn, "基准", "全球开放", region, d.split("|")[0], d.split("|")[1])] = \
        flow["src_panels"][(scn, "基准", "全球开放", "车辆", d)][pos]
    check_sheet("车辆回收铜_来源地", 6, flow["n_t"], 3 * 3 * (15 + 32) * 14, src_samples, "车辆回收铜_来源地")

    srcp_samples = {}
    d = flow["details"]["绿电"][0]
    # 不开放情景的来源地只有15个处理区有行，因此用处理区做抽查样本
    proc_region = flow["orig_of_canon"][flow["proc_canon"][0]]
    pos_proc = flow["m_regions"].index(flow["canon_of"][proc_region])
    srcp_samples[(scn, "基准", "不开放", proc_region,
                  cfg["devices"]["subtech_to_tech"][d], d)] = \
        flow["src_panels"][(scn, "基准", "不开放", "绿电", d)][pos_proc]
    check_sheet("绿电回收铜_来源地", 6, flow["n_t"], 3 * 3 * (15 + 32) * 11, srcp_samples, "绿电回收铜_来源地")

    gl_samples = {}
    dev0 = flow["devices"][0]
    gl_samples[(scn, "基准", "不开放", dev0)] = flow["totals"][(scn, "基准", "不开放", dev0)]
    check_sheet("全球六组合_面板", 4, flow["n_t"], 3 * 3 * 2 * 8, gl_samples, "全球六组合_面板")

    # 双边废铜：检查行数与键唯一性（含贸易情景列，避免对角线路线在两个贸易情景下重复）
    ws = wb["双边废铜_面板"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    n = 0
    seen = set()
    dup = 0
    for row in it:
        if row[0] is None:
            continue
        n += 1
        key = tuple(row[:7])
        if key in seen:
            dup += 1
        seen.add(key)
    out.append({"检查项目": "写盘核对-双边废铜_面板", "检查数量": n, "最大误差": None,
                "是否通过": "通过" if (n == 35910 and dup == 0) else "未通过",
                "说明": f"行数={n}（期望35910）；重复键={dup}；表头={header[:8]}"})
    wb.close()
    return out


# --------------------------------------------------------------------------- #
# Markdown 方法说明
# --------------------------------------------------------------------------- #
def fmt(x, nd=0):
    if x is None:
        return "—"
    return f"{x:,.{nd}f}"


def build_markdown(cfg, rec, trade_data, flow, r3, step1, checks, counts, bilat_count,
                   run_info, baseline_hydro=None):
    devices = flow["devices"]
    years = flow["years"]
    scn_list = cfg["scenarios"]["canonical"]
    L = []
    A = L.append

    A("# 能源转型设备退役与次生铜核算：计算流程、方法与文献说明")
    A("")
    A("> 本文档由 `step2_secondary_copper.py` 在全部核验通过后自动生成，内容取自实际运行的代码、配置、"
      "输入文件与结果面板。文中数字均为本次运行的真实输出，不含占位符。")
    A("")
    A(f"- 生成时间：{run_info['time']}")
    A(f"- 代码版本（HEAD）：`{run_info['git_head']}`")
    A(f"- 配置：`{cfg['_config_path']}`（sha256 `{cfg['_config_sha256']}`）")
    A(f"- 运行环境：Python {run_info['python']}；numpy {run_info['numpy']}；openpyxl {run_info['openpyxl']}")
    A("")

    # ---------------- A ----------------
    A("## A. 研究范围与情景结构")
    A("")
    A("### A.1 设备范围")
    A("")
    A("只计算铜存量工作簿中实际出现的设备与组合，不扩展到 FCEV、传统燃油车或其他常规发电设备。")
    A("")
    A("| 部门 | 设备 / 子技术 | 原技术（技术明细） | 数量 |")
    A("|---|---|---|---|")
    A("| 车辆 | " + "、".join(cfg["devices"]["vehicle_subsectors"]) +
      " × BEV/PHEV | Subsector + Technology | 最多 7×2=14 |")
    A("| 发电 | PV、陆上风电、海上风电、水电、核电、六种CCS | Subtech | 11 |")
    A("")
    A("### A.2 情景结构")
    A("")
    A("| 维度 | 取值 | 说明 |")
    A("|---|---|---|")
    A("| 能源路径 | " + "、".join(scn_list) + " | SSP1-2.6 / SSP2-4.5 / SSP5-8.5；各文件写法不同，已显式映射 |")
    A("| 回收情景 | " + "、".join(REC_ORDER) + " | 收集率与综合处理率的三种组合 |")
    A("| 贸易情景 | " + "、".join(TRADE_ORDER) + " | 是否允许GCAM区域间废铜流动 |")
    A("")
    A("每个能源路径 3×2=6 个回收×贸易组合，三条路径共 18 个组合，全部生成。")
    A("")
    A("### A.3 时间范围")
    A("")
    A("| 项目 | 年份 | 说明 |")
    A("|---|---|---|")
    A(f"| 车辆历史追溯 | {cfg['years']['vehicle_start']}—{cfg['years']['end']} | 首个正值节点作为初始批次，"
      "其后按年插值；输出年份沿用原模板 |")
    A(f"| 发电历史追溯 | {cfg['years']['power_start']}—{cfg['years']['end']} | 首个正值节点（多数为1990）作为初始批次，其后按年插值 |")
    A(f"| 正式回收比较期 | {years[0]}—{years[-1]} | 退役铜与回收率数据的共同覆盖期 |")
    A(f"| 回收率校准年 / 目标年 | {cfg['years']['recycling_calibration_year']} / {cfg['years']['recycling_target_year']} | "
      "目标年仍为2060，不改为2050提前达标 |")
    A("")
    A("### A.4 处理边界（15区 / 17区）")
    A("")
    A("| 类别 | 区域数 | 处理能力设定 | 矩阵特征 |")
    A("|---|---|---|---|")
    A("| 本地处理区域 | 15 | 有本地处理能力 | 对角线=1，100%本地处理，并可接收进口 |")
    A("| 全出口区域 | 17 | 无本地处理能力 | 本地份额=0，100%出口至15个接收地 |")
    A("")
    A("这15/17的划分来自 `Nature_区域运输流向与铜价值分摊_全球贸易情景（运输成本）.xlsx` 的"
      "「贸易原始矩阵_32x32」对角线结构，并作为**研究设定**执行：它表示本研究假设的处理能力边界，"
      "不是对现实世界处理能力的实测判断，也不能仅凭矩阵证明某些地区现实中完全没有处理能力。")
    A("")
    A("### A.5 证据性质分层")
    A("")
    A("| 层次 | 内容 |")
    A("|---|---|")
    A("| 数据事实 | 保有量/装机面板（辆、MW）、CIPG、寿命参数、回收率逐年表、贸易矩阵（均来自本地输入文件） |")
    A("| 文献参数 | life 表中的 Weibull/Normal/Fixed 参数及其文献出处；水电规模寿命与规模铜强度（[4] 表1、表5） |")
    A("| 项目数据 | 中国水电总装机 380.574 GW（`保有量.xlsx` / 装机容量_九类_面板，三条能源路径一致） |")
    A("| 代理映射 | 设备类别→废弃物类别（ELV/WEEE/IEW）、Mini Car/Large Car 沿用 Passenger Car、"
      "水电子技术映射、中国CON规模结构外推到全球32区 |")
    A("| 研究假设 | 首个正值节点初始化、投入年铜强度固定、CIPG≤2021用2021值并在节点间插值、"
      "经济性提前退役、水电CON规模结构（1990—2021固定份额、大型内部49:15）、"
      "15/17处理能力边界、全量处理无产能上限、静态贸易矩阵、跨设备沿用同一矩阵 |")
    A("")

    # ---------------- B ----------------
    A("## B. 全流程与两个脚本的职责")
    A("")
    A("```mermaid")
    A("flowchart TD")
    A("  A1[保有量.xlsx 车辆保有量_面板 辆 / 装机容量_九类_面板 MW] --> B1[按序列取首个正值节点 t0]")
    A("  B1 --> C1[有效节点线性插值到年 Q(t)]")
    A("  A2[life 基准寿命分布] --> D1[S(a), p(a)=S(a-1)-S(a)]")
    A("  C1 --> E1[批次引擎：R_t=ΣI_τ·p(t-τ)]")
    A("  D1 --> E1")
    A("  E1 --> F1[I_t=Q_t-Q_t-1+R_t]")
    A("  F1 -->|I<0| G1[经济性提前退役 E_t=-I_t（最老批次）]")
    A("  F1 -->|I≥0| H1[新增投入批次]")
    A("  G1 --> I1[退役设备量 = 自然 + 提前]")
    A("  H1 --> J1[新增设备量]")
    A("  A3[数据.xlsx CIPG / Automotive] --> K1[投入年铜强度 k(τ)]")
    A("  I1 --> L1[退役铜量 = Σ退出×k(τ)]")
    A("  J1 --> M1[新增铜需求量 = I_t×k(t)]")
    A("  K1 --> L1")
    A("  K1 --> M1")
    A("  L1 --> N1[来源区收集 W×C]")
    A("  N1 --> O1{贸易情景}")
    A("  O1 -->|不开放| P1[15区本地处理 ×P]")
    A("  O1 -->|全球开放| P2[32区按矩阵 T 分配至15处理区 ×P]")
    A("  P1 --> Q1[次生铜产量]")
    A("  P2 --> Q1")
    A("  Q1 --> R1[处理地生产 / 来源地归属 / 双边流 / 情景差值]")
    A("```")
    A("")
    A("| 脚本 | 职责 | 主要输入（工作表） | 主要输出 |")
    A("|---|---|---|---|")
    A("| `step1_retired_copper.py` | 设备量驱动：t0 初始化、年度化、寿命离散、批次递推、经济性提前退役、投入年铜强度 | "
      "保有量.xlsx（车辆保有量_面板、装机容量_九类_面板）；数据.xlsx（CIPG、life、Automotive copper demand）；"
      "铜存量.xlsx（仅核对） | 能源转型设备退役铜量_第一步.xlsx（18表：四类结果+审计） |")
    A("| `step2_secondary_copper.py` | 收集、贸易分配、处理、次生铜产量与情景比较 | 第一步退役铜面板；"
      "全球退役设备铜回收率_四情景逐年结果.xlsx；Nature贸易矩阵 | 能源转型设备次生铜产量_第二步.xlsx（12表）；本MD |")
    A("")
    A("两个脚本均以 `main()` 为入口，`import` 时不执行计算；第二步只读取第一步的退役铜数值，不重算退役模型。"
      "第一步先写临时成果，第二步用 `--step1` 读取该临时成果；两步核验通过后由 `--promote` 备份并替换正式文件。")
    A("")
    A("**表结构说明**：「双边废铜_面板」在 Scenario/回收情景/来源区/处理区/设备类别/流量类型/单位 之外"
      "增加「贸易情景」列。原因是15条对角线本地路线在不开放与全球开放下都有取值，若无该列会出现重复键。")
    A("")

    # ---------------- C ----------------
    A("## C. 第一步：设备量驱动的公式与假设")
    A("")
    A("### C.1 唯一设备路径与四类结果")
    A("")
    A("本次不再以“铜存量面板”驱动退役，而是以**设备量本身**驱动：")
    A("")
    A("| 部门 | 驱动量 $Q_t$ | 单位 | 来源 |")
    A("|---|---|---|---|")
    A("| 车辆 | 保有量 | 辆 | `保有量.xlsx` / 车辆保有量_面板 |")
    A("| 发电 | 装机容量 | MW | `保有量.xlsx` / 装机容量_九类_面板（水电按 CON 份额拆规模） |")
    A("")
    A("四类结果分别输出：新增设备/装机量、退役设备/装机量、新增铜需求量、退役铜量；"
      "`铜存量.xlsx` 仅用于核对，不参与驱动。这样“新增铜需求”由设备新增决定，"
      "不会因为铜强度下降而被压低，也不会出现“铜凭空缩水导致的退役”。")
    A("")
    A("$$R_t=\\sum_{\\tau<t} I_\\tau\\,p(t-\\tau)$$")
    A("")
    A("$$I_t=Q_t-Q_{t-1}+R_t$$")
    A("")
    A("$$C^{in}_t=I_t\\cdot k(t),\\qquad C^{out}_t=\\sum_{\\tau<t}\\big(\\text{退出量}_\\tau\\big)\\cdot k(\\tau)$$")
    A("")
    A("其中 $k(\\tau)$ 是批次投入年的铜强度（车辆 kg/辆，发电 t/GW），投入后固定。")
    A("")
    A("### C.2 初始化与年度化")
    A("")
    A("- 取序列**首个正值节点**为投入起点 $t_0$；$t<t_0$ 时 $Q=0$、无投入无退役；"
      "$Q(t_0)$ 作为初始批次（年龄 0，**当年不退役**）。这是研究初始化假设，不代表真实投运年份。")
    A("- $t>t_0$ 在有效节点之间线性插值到年，保留全部节点；空白不当作 0，按相邻有效节点插值并登记；"
      "不在年份范围外外推；全零序列保持 0（结构性零）。")
    A("- 车辆输出年份沿用原模板 2010—2050，发电 1990—2050；内部历史可从 1990 年开始。")
    A("")
    A("### C.3 寿命分布与年度退役概率")
    A("")
    A("| 分布 | 生存函数 | 说明 |")
    A("|---|---|---|")
    A("| Weibull | $S(a)=\\exp[-(a/\\lambda)^k]$ | λ为尺度，不是均值；不重复做Gamma换算 |")
    A("| Normal | $S(a)=P(X>a\\mid X>0)$ | 截断在寿命>0并归一化 |")
    A("| Fixed | $S(a)=1\\ (a<L)$，$0\\ (a\\ge L)$ | 按确定年龄退役 |")
    A("")
    A("年度退役概率 $p(a)=S(a-1)-S(a)$，$p(0)=0$。对批次初始量使用 $p(a)$，"
      "与“先扣累计退役、再乘条件退役率 $1-S(a)/S(a-1)$”在数学上等价；二者不可混用。"
      "寿命映射见第一步「寿命参数」，文献来源 [4][5]。水电按 CON 三规模分别取 Normal(50,15)/(70,15)/(80,15)。")
    A("")
    A("### C.4 批次递推与经济性提前退役")
    A("")
    A("每年按顺序：①对全部既有批次计算自然退役 $R_t$；②由 $I_t=Q_t-Q_{t-1}+R_t$ 反推新增；"
      "③若 $I_t<0$，把缺口 $E_t=-I_t$ 记为**经济性提前退役**（按最老批次优先扣减），令 $I_t=0$；"
      "④更新批次与期末存量。")
    A("")
    A("$$\\text{期末存量}=\\text{期初存量}+\\text{新增}-\\text{自然退役}-\\text{提前退役}$$")
    A("")
    A("提前退役只由**设备量本身下降**触发（如 PHEV 被 BEV 替代、煤电/CCS 退出），"
      "不再用于吸收铜强度下降；它与自然退役分列，退役主面板填写两者之和。"
      "浮点级缺口（$|I_t|\\le\\max(10^{-6},10^{-9}Q_t)$）按 0 处理并登记；实质负需求全部转为提前退役。")
    A("")
    A("### C.5 铜强度：投入年固定（A 方案）")
    A("")
    A("- 车辆：`Automotive copper demand` 的 kg/vehicle 为常数，投入年即用此值。")
    A("- 发电：CIPG（t/GW）。投入年 $\\le 2021$ 用 2021 值（最早有效年限）；2021 之后在 "
      "2021/2030/2035/…/2060 节点之间**线性插值**，禁止把 2025 年直接跳配 2030 年值；不外推。")
    A("- 退役批次使用其**投入年**强度，不使用退役年强度——否则会把老设备按新强度重估，"
      "重新引入“强度下降导致铜凭空减少”的问题。")
    A("- 与 `铜存量.xlsx` 的差异属预期：旧表在 2025 年使用 2030 年 CIPG（如 PV 隐含 4 500.713），"
      "新规则插值后 2025 年 PV 约高 1.1%；该差异列入核对说明，不视为错误。")
    A("")
    A("### C.6 水电：CON 规模结构、容量拆分与汇总")
    A("")
    hc = cfg["life"]["hydro_con"]
    hw = step1.get("hydro_weights", {})
    node_years = [2021, 2025, 2030, 2035, 2040, 2045, 2050]

    def wval(kind, scale, year):
        return hw.get((kind, scale), {}).get(year)

    A("**为什么用 CON。** 补充方法9中，MOD/PRO 情景暂停新建小水电，会系统性低估小水电较短寿命对应的退役铜；"
      "CON 情景采用 UNIDO 预测（小水电 2030 年 93 GW、2050 年 100 GW、2060 年 102 GW）[4]，"
      "保留小水电发展。该目标在 [4] 中转引自 UNIDO[11]，本研究未直接核对 UNIDO 原始报告。")
    A("")
    A("**规模份额。** 1990—2021 固定 小型 20%、中型 16%、大型 64%（大型含抽蓄；"
      "由 2021 参考构成 RoR 20 / RES 65 / PSH 15 映射：中型 16 为 RES 的一部分，故大型水库=65−16=49）。"
      "2022—2050 中型固定 16%，小型按中国 CON 小水电装机路径换算，大型取余：")
    A("")
    A("$$K_{small}(t)=76.1148+(93-76.1148)\\frac{t-2021}{9},\\quad 2021\\le t\\le 2030$$")
    A("")
    A("$$K_{small}(t)=93+(100-93)\\frac{t-2030}{20},\\quad 2030<t\\le 2050$$")
    A("")
    A("$$w_{small}(t)=\\frac{K_{small}(t)}{380.574},\\quad w_{medium}=0.16,\\quad w_{large}=1-w_{small}-w_{medium}$$")
    A("")
    A("中国水电总装机 380.574 GW 取自 `保有量.xlsx`（China-hydro 2021—2050 均为 380 574 MW，三条能源路径一致，已复核）。"
      "该容量份额统一用于全球 32 区与三条能源路径，是**中国 CON 规模结构的全球代理外推**，不是各区域实测。")
    A("")
    A("容量份额关键节点（第一步「水电规模权重_面板」）：")
    A("")
    A("| 年份 | 小型 | 中型 | 大型（含抽蓄） |")
    A("|---|---:|---:|---:|")
    for y in node_years:
        ws_, wm_, wl_ = (wval("容量份额", s, y) for s in ("small", "medium", "large"))
        A(f"| {y} | {ws_:.3%} | {wm_:.3%} | {wl_:.3%} |")
    A("")
    A("**在装机层面拆分。** $Q_i(t)=Q_{hydro}(t)\\cdot w_i(t)$，三类分别建立批次模型，"
      "再用各自的规模铜强度把装机换算成铜；不再做“容量份额→铜量份额”的存量再分配。"
      "规模铜强度取补充表5：小型径流 4 209、中型水库 2 931、大型水库 1 713、抽蓄 3 166 t/GW；"
      "大型按 2021 参考构成 49:15 加权：")
    A("")
    A("$$I_{large}=\\frac{49\\times1713+15\\times3166}{64}=2053.547\\ \\text{t/GW}$$")
    A("")
    A("补充表5注明各强度按 3‰/年递减；为使规模强度与项目 hydro CIPG 路径一致，本实现按项目路径整体缩放，"
      "并用 2021 年 20/16/64 加权值归一（表5加权 2 625.0 vs 项目 2 617.155，归一后一致）：")
    A("")
    A("$$I_i(t)=I_i^{T5}\\cdot\\frac{CIPG_{hydro}(t)}{CIPG_{hydro}(2021)}\\cdot\\frac{CIPG_{hydro}(2021)}{\\sum_j w_j(2021)I_j^{T5}}$$")
    A("")
    A("大型内部固定 49:15 为本实现假设（论文未给未来固定内部比例）。规模寿命 Normal(50,15)/(70,15)/(80,15)，"
      "分别批次计算后汇总：新增、退役、新增铜、退役铜、铜存量均按规模求和；正式面板只输出 tech=hydro、Subtech=hydro 一行，"
      "规模明细只进「水电规模权重_面板」与核验表，避免与合计重复求和。")
    A("")
    A("**提前退役与水电。** CON 份额使大型份额下降、小型上升：小型需要新建，大型则可能出现"
      "“装机目标低于存活量”的年份，此时按最老批次记为经济性提前退役（单列）。"
      "这与自然寿命退役分开记录，不再用铜强度缩水解释。本次规模级结果见第一步核验表。")
    A("")

    # ---------------- D ----------------
    A("## D. 第二步：公式与假设")
    A("")
    A("记 $W[r,k,t]$ 为来源区 $r$、设备类别 $k$、年份 $t$ 的退役含铜量（来自第一步）。")
    A("")
    A("| 符号 | 含义 | 取值来源 |")
    A("|---|---|---|")
    A("| $C[s,r,k,t]$ | 收集率（废铜产生地） | 「逐年回收率」F列；基准情景冻结为2021年C0 |")
    A("| $P[s,j,k,t]$ | 综合处理率（出口目的地/实际处理地） | 「逐年回收率」G列；基准情景冻结为2021年P0 |")
    A("| $T[r,j]$ | 贸易分配比例 | 「贸易原始矩阵_32x32」，行=来源区，列=处理区，对角线=本地处理 |")
    A("| $X$ | 入厂含铜量 | 由 $W$、$C$、$T$ 计算 |")
    A("| $Q$ | 次生铜产量 | 由 $X$、$P$ 计算 |")
    A("")
    A("### D.1 不开放")
    A("")
    A("$$Q[s,r,k,t]=W[r,k,t]\\times C[s,r,k,t]\\times P[s,r,k,t],\\quad r\\in\\{15\\text{个处理区}\\}$$")
    A("")
    A("17区不生成结果行，不建立暂存量、未利用量或跨年库存账户；其退役量不被填为0，而是完全不进入该情景的合计。")
    A("")
    A("### D.2 全球开放")
    A("")
    A("$$X[s,r,j,k,t]=W[r,k,t]\\times C[s,r,k,t]\\times T[r,j]$$")
    A("")
    A("$$Q[s,r,j,k,t]=X[s,r,j,k,t]\\times P[s,j,k,t]$$")
    A("")
    A("收集率取废铜产生地，综合处理率取处理地。六种组合的规则：")
    A("")
    A("| 组合 | 收集率 | 综合处理率 |")
    A("|---|---|---|")
    A("| 基准 × 不开放 | $C_0$ | $P_0$ |")
    A("| 政策推动 × 不开放 | $C(t)$ | $P_0$ |")
    A("| 技术进步 × 不开放 | $C_0$ | $P(t)$ |")
    A("| 基准 × 全球开放 | $C_0[r]$ | $P_0[j]$ |")
    A("| 政策推动 × 全球开放 | $C(t)[r]$ | $P_0[j]$ |")
    A("| 技术进步 × 全球开放 | $C_0[r]$ | $P(t)[j]$ |")
    A("")
    A("综合回收率 $=C\\times P$ 已包含其定义覆盖的收集后损失，不再另加拆解、分选或冶炼效率重复扣损。"
      "本地处理可直接用 $W\\times$ 综合回收率；本脚本按 $C\\times P$ 计算并与原表H/I/J列逐值核对一致。")
    A("")
    A("### D.3 关键假设与边界")
    A("")
    A(f"- {cfg['assumptions']['no_capacity_limit']}")
    A(f"- {cfg['assumptions']['trade_matrix_scope']}{cfg['trade']['static_matrix']}")
    A("- 运输成本与铜价值分摊参数不参与物质量公式；不计算再生铜返还来源地、库存、原生铜替代与水压力。")
    A("- 来源归属回收铜量是“按来源区归属的回收铜”，不等于来源区本地产量，也不假定再生铜自动返还来源区。")
    A("- 开放可增加全球产量的原因：不开放时17区退役铜不处理；开放后其已收集铜可出口到15区处理。"
      "国际贸易本身不创造额外退役含铜量，因此不要求开放与不开放的全球产量相等，也不预设差值为正。")
    A("- **水电口径**：第一步内部按小型/中型/大型分别计算，但对外只输出一条 hydro 退役铜；"
      "第二步读到的是已合并的 hydro，再按“水电→IEW”统一套用回收与贸易参数。"
      "三种规模共用同一回收类别与贸易规则，因此先合并退役铜再做回收计算不改变结果，"
      "也避免规模明细与合计重复计量。")
    A("")

    # ---------------- E ----------------
    A("## E. 参数—文献对应表")
    A("")
    A("| 参数/规则 | 适用设备或区域 | 实际值/公式 | 原始文献 | 位置 | 本地文件及工作表 | 证据性质 | 本研究调整 |")
    A("|---|---|---|---|---|---|---|---|")
    A("| 乘用车寿命 | Car、Mini Car、Large Car and Truck | Weibull k=2.07，λ=14.46 | Zhang et al. (2025) [5]，DOI 10.1038/s41467-025-61481-y（上游 Hao et al. 2011） | Supplementary Table 2；Supplementary Eq.(6) | 数据.xlsx / life；寿命分布文献核验与映射说明.md | 文献直接值（经本地核验） | Mini Car、Large Car 为共享/代理映射，置信度中/低 |")
    A("| 轻/中/重卡寿命 | Light/Medium/Heavy truck | Weibull k=5.58，λ=8.02/10.09/12.8 | 同上 [5]（上游 Hao et al. 2011） | Supplementary Table 2 | 同上 | 文献直接值 | 无 |")
    A("| 公交车寿命 | Bus | Fixed 9年 | Hao et al. (2011) [5]，DOI 10.1007/s11431-010-4256-1 | Table 1 / Section 3.3 | 同上 | 文献直接值（代理） | 源表Bus未注明营运属性，按Transit bus代理 |")
    A("| PV寿命 | PV | Weibull k=5.3759，λ=30 | Dong, H. et al. (2025) [4]，DOI 10.1038/s41467-025-56592-5（上游 Kim & Park 2018） | Supplementary Note 2 / Table 1 | 同上 | 文献直接值 | 无 |")
    A("| 风电寿命 | 陆上/海上风电 | Weibull k=4.25，λ=19.48 | 同上 [4]（上游 Cao et al. 2019） | Supplementary Table 1 | 同上 | 文献直接值（共享） | 原文未分海陆，海上沿用同一分布 |")
    A("| 水电小型寿命 | hydro 内部 small（径流式） | Normal(μ=50, σ=15) | 同上 [4]（上游 Li et al. 2020，DOI 10.1016/j.resconrec.2020.104906） | 补充表1（本地中文版 PDF 第4页） | 数据.xlsx / life；第一步「寿命参数」 | 文献直接值 | 截断>0并归一化；内部规模，正式结果合并为hydro |")
    A("| 水电中型寿命 | hydro 内部 medium（水库式） | Normal(μ=70, σ=15) | 同上 [4] | 补充表1（PDF 第4页） | 同上 | 文献直接值 | 同上 |")
    A("| 水电大型寿命 | hydro 内部 large（含水抽蓄） | Normal(μ=80, σ=15) | 同上 [4] | 补充表1（PDF 第4页） | 同上 | 文献直接值 | 同上 |")
    A("| 水电规模铜强度 | 小型径流/中型水库/大型水库/抽蓄 | 4209 / 2931 / 1713 / 3166 t/GW | 同上 [4] | 补充表5（PDF 第11页） | 第一步「寿命参数」；config.life.hydro_con | 文献直接值 | 均按3‰/年递减；三规模共用同一系数，归一化中抵消 |")
    A("| 水电CON规模份额 | 全球32区（中国代理外推） | 1990—2021 小20%/中16%/大64%；2022—2050 中16%、小按76.1148/93/100 GW对380.574 GW线性、大取余 | 同上 [4] | 补充方法9（PDF 第15—16页） | config.life.hydro_con；第一步「水电规模权重_面板」 | 文献直接值＋项目数据（380.574 GW）＋插值 | 中国规模结构外推到全球32区；93/100 GW不按区域分配 |")
    A("| 水电大型内部构成 | 大型 = 大型水库49 + 抽蓄15 | 49:15 固定；$I_{large}=(49\\times1713+15\\times3166)/64$ | 同上 [4] | 补充方法9（2021参考构成） | config.life.hydro_con | 二次计算＋实现假设 | 论文未给未来固定49:15，固定该比例为本实现假设 |")
    A("| 水电规模历史权重 | 1990—2021 | 20%/16%/64% 固定 | 同上 [4] | 补充方法9（2021参考构成） | 同上 | 研究假设（历史回填） | 非1990—2020逐年观测 |")
    A("| 核电寿命 | 核电 | Normal(40,8) | 同上 [4]（上游 Li et al. 2020） | Supplementary Table 1 | 同上 | 文献直接值（共享） | Gen_II_LWR/Gen_III 映射 PWR 统一分布 |")
    A("| CCS寿命 | 六种CCS | Normal(35,10) | 同上 [4]（上游 Li et al. 2020） | Supplementary Table 1 | 同上 | 文献直接值（共享） | 未按燃料/捕集工艺区分 |")
    A("| 铜强度 CIPG | PV/风/水/核/CCS | 2021—2060 逐年（t/GW） | 数据.xlsx / CIPG；来源见 convert_power.py 与铜需求方法说明 | 全表 | 数据.xlsx / CIPG | 数据事实 | 2010/2015用2021值、2025用2030值（沿用既有代码口径）；1990—2009用2021值回填 |")
    A("| 容量因子 CF | 非CCS发电 | 分区域分技术（0—1） | 数据.xlsx / CF | 2010—2050 | 数据.xlsx / CF | 数据事实 | 1990—2009使用2010年值向前回填 |")
    A("| CCS容量因子 | 六种CCS | 0.80/0.85/0.80/0.80/0.80/0.80 | GCAM通用参数（电力部门_容量因子.xlsx） | electricity_water，1990年起 | 电力部门_容量因子.xlsx / electricity_water | 数据事实（GCAM通用参数） | 按“基础技术名+冷却方式后缀”显式匹配 |")
    A("| 基准收集率 C0 | 32区×7设备 | Global 0.58、China 0.56、EU-15 0.60、Japan 0.65、USA/Canada 0.61、India 0.58、Korea 0.61、Australia_NZ 0.58、Brazil 0.25、Africa 0.34，其余Global代理 | 全球退役设备铜回收率_四情景逐年结果.xlsx（基准与区域映射）；上游 ICSG/ICA [8] 与 Dong et al. (2020) [1] | 基准与区域映射 | 回收率工作簿 | 数据事实（含代理映射） | 区域代理关系沿用原表，不另设 |")
    A("| 基准处理率 P0 | ELV/WEEE/IEW | 0.54/0.54/0.69 | Glöser et al. (2013) [2]，DOI 10.1021/es400069b，Table 1 废弃物类别值 | Table 1 | 回收率工作簿（基准与区域映射）；上游说明见原表批注 | 文献直接值（类别层面） | 同类设备共用类别处理率 |")
    A("| 90%/95%目标与线性路径 | 32区×7设备 | 校准年2021、目标年2060，线性 | Dong et al. (2020) [1]，DOI 10.1016/j.resconrec.2020.104943 | Section 2.4 | 逐年回收率表（F/G列） | 二次计算（本地上游已生成逐年值） | 原文目标年为2100，本研究为2060，属更积极情景 |")
    A("| 贸易矩阵 T | 32来源区×15处理区 | 15条本地路线+255条跨区路线，行和=1 | Wang et al. (2026) [3]，DOI 10.1038/s41586-026-10905-w | Ext. Fig. 8b（源数据 MOESM15） | Nature_...贸易情景（运输成本）.xlsx / 贸易原始矩阵_32x32 | 数据事实（论文源数据） | 七类设备统一沿用光伏贸易矩阵，属**本研究外推** |")
    A("| 15/17处理能力 | 区域 | 15区有、17区无 | 同上 [3]（矩阵对角线结构） | 贸易矩阵 | 同上 | 研究设定 | 不代表现实产能实测 |")
    A("| 设备→废弃物类别 | 7类设备 | ELV/WEEE/IEW | Glöser et al. (2013) [2] | Table 1 | 回收率工作簿 | 代理映射 | 六种CCS与核电/水电/风电同归IEW |")
    A("| 初始化 | 车辆/发电 | 首个正值节点作为首批次；车辆2010起、发电1990起可见 | — | — | — | 研究假设 | 不表示真实投运年份 |")
    A("| 历史回填 | 投入年早于系数最早年 | CIPG 用 2021 值 | — | — | — | 研究假设 | 已在方法说明标注 |")
    A("| 经济性提前退役 | 全部序列 | 设备量下降且缺口超过自然退役时，按最老批次退出（缺口=提前退役量） | — | — | — | 模型假设 | 与自然退役分列；仅由设备量下降触发 |")
    A("")

    # ---------------- F ----------------
    A("## F. 引用与参考文献")
    A("")
    A("正文按编号引用；每条引用紧随其支持的方法或参数。")
    A("")
    A("[1] Dong, D., Tercero Espinoza, L. A., Loibl, A., Pfaff, M., Tukker, A., & van der Voet, E. (2020). "
      "Scenarios for anthropogenic copper demand and supply in China: Implications of a scrap import ban and a "
      "circular economy transition. *Resources, Conservation and Recycling*, 161, 104943. "
      "https://doi.org/10.1016/j.resconrec.2020.104943 —— 用于90%收集率/95%处理率目标与线性提升方法；"
      "原文中国在2100年达标，本研究把目标年设为2060年，属更积极的情景假设，不是原文对2060年的预测。")
    A("")
    A("[2] Glöser, S., Soulier, M., & Tercero Espinoza, L. A. (2013). Dynamic analysis of global copper flows: "
      "Global stocks, postconsumer material flows, recycling indicators, and uncertainty evaluation. "
      "*Environmental Science & Technology*, 47(12), 6564–6572. https://doi.org/10.1021/es400069b —— "
      "废弃物类别（ELV/WEEE/IEW）综合处理率 0.54/0.54/0.69 的类别层面来源。")
    A("")
    A("[3] Wang, C., Zuo, J., Chen, X., et al. (2026). Towards an equitable future of global photovoltaic waste "
      "recycling. *Nature*, 656, 888–896. https://doi.org/10.1038/s41586-026-10905-w —— "
      "开放贸易情景框架、32×32 区域流向矩阵与源数据（Ext. Fig. 8b，源数据文件 MOESM15）。"
      "本研究把该光伏贸易矩阵外推到全部七类设备，属本研究设定，不是原文结论。"
      "卷期页码沿用本地资料记录，本次未在出版商页面复核，标为待核实；DOI 与源数据文件对应关系已在本地核对。")
    A("")
    A("[4] Dong, H., Zhang, T., Geng, Y., Wang, P., Zhang, S., & Sarkis, J. (2025). Sub-technology market "
      "share strongly affects critical material constraints in power system transitions. "
      "*Nature Communications*, 16, 1285. https://doi.org/10.1038/s41467-025-56592-5 —— "
      "本研究水电规模结构的核心依据（本地中文本补充材料）："
      "补充表1（本地 PDF 第4页）给出小/中/大型水电寿命 Normal(50,15)/(70,15)/(80,15)；"
      "补充表5（PDF 第11页）给出规模铜强度 4209/2931/1713/3166 t/GW 及“每年3‰递减”；"
      "补充方法9（PDF 第15—16页）给出中型16%、CON情景小水电2030年93 GW、2050年100 GW、2060年102 GW，"
      "以及2021年 RoR 20%/RES 65%/PSH 15% 的参考构成。"
      "上游为 Li et al. (2020), *Resources, Conservation and Recycling*, 160, 104906, "
      "https://doi.org/10.1016/j.resconrec.2020.104906（寿命参数）。")
    A("")
    A("[5] Zhang, B., Xin, Q., Chen, S., Wang, B., Li, H., Wang, Z., & Bansal, P. (2025). Lithium-ion battery "
      "recycling relieves the threat to material scarcity amid China's electric vehicle ambitions. "
      "*Nature Communications*, 16, 6661. https://doi.org/10.1038/s41467-025-61481-y —— "
      "乘用车与轻/中/重卡 Weibull 参数的直接取值来源（Supplementary Table 2、Supplementary Eq.(6)）；"
      "上游为 Hao et al. (2011), *Science China Technological Sciences*, 54, 625–629, "
      "https://doi.org/10.1007/s11431-010-4256-1（Table 1 / Section 3.3，含 Transit bus 9年寿命）。"
      "作者与卷期页码已按 Crossref 元数据核对。")
    A("")
    A("[6] Ryter, J., Fu, X., Bhuwalka, K., et al. (2021). Emission impacts of China's solid waste import ban "
      "and COVID-19 in the copper supply chain. *Nature Communications*, 12, 3753. "
      "https://doi.org/10.1038/s41467-021-23874-7 —— 仅作为废铜贸易政策研究动机的背景依据，"
      "不提供本研究的贸易矩阵或回收参数。")
    A("")
    A("[7] Klose, S., & Pauliuk, S. (2021). Quantifying longevity and circularity of copper for different "
      "resource efficiency policies at the material and product levels. *Journal of Industrial Ecology*, "
      "25, 979–993. https://doi.org/10.1111/jiec.13092 —— 仅用于说明收集、处理等政策干预维度的背景，"
      "不是本研究全部数值的直接来源。")
    A("")
    A("[8] International Copper Study Group (2020). *The World Copper Factbook 2020*. "
      "https://copperalliance.org/wp-content/uploads/2021/01/2020_10_13_ICSG_Factbook_2020.pdf —— "
      "全球EoL收集率56%、处理率71%、综合回收率约40%的历史参考值（经本地回收率表间接使用）。")
    A("")
    A("[9] International Energy Agency (2021). *Hydropower Special Market Report – Analysis and forecast to "
      "2030*. https://iea.blob.core.windows.net/assets/4d2d4365-08c6-4171-9ea2-8549fabd1c8d/"
      "HydropowerSpecialMarketReport_corr.pdf —— 据 [4] 补充方法9转引（编号71），"
      "用于小水电与抽蓄的规模归属判断；本次**未直接核对 IEA 报告原文**。")
    A("")
    A("[10] 中华人民共和国水利部、中国国家统计局（2013）.《第一次全国水资源普查公报》，"
      "《中华人民共和国水利部公报》，53–57。—— 据 [4] 补充方法9转引（编号72），"
      "为中型水电（50—300 MW）占水电总装机16%的上游来源；本次**未直接核对公报原文**。")
    A("")
    A("[11] United Nations Industrial Development Organization (2020). *World Small Hydropower Development "
      "Report 2019 – Asia and Oceania Region*. "
      "https://www.unido.org/sites/default/files/files/2021-11/IDR%202022%20-%20EBOOK.pdf —— "
      "据 [4] 补充方法9转引（编号76），为 CON 情景小水电 2030 年 93 GW、2050 年 100 GW、"
      "2060 年 102 GW 的上游来源；本次**未直接核对 UNIDO 原始报告**。")
    A("")
    A("未用于本次计算的成本、价值分摊与延寿文献（如光伏运输成本分摊、Low/Medium/High延寿情景）"
      "仅作背景，不作为本次模型输入依据。")
    A("")
    A("**引用核对说明**：本次通过 Crossref 元数据核对了 [1][2][4][5] 的作者、题名、卷期与年份；"
      "[3] 的卷期页码沿用本地记录、未在出版商页面复核（已标待核实），DOI 与源数据文件对应关系在本地核对。"
      "本地《寿命分布文献核验与映射说明》把 DOI 10.1038/s41467-025-61481-y 的作者写作 “Dunn et al. (2021)”"
      "（实际是另一篇 ES&amp;T 2021 论文），本次按 Crossref 更正为 Zhang et al. (2025)；"
      "该更正只影响引用写法，不影响寿命参数取值（参数取自 数据.xlsx / life 并对应到该文补充材料表号）。"
      "[9][10][11] 为 [4] 补充方法9的上游引用，本次仅按 [4] 的中文本补充材料转引，"
      "未直接核对原始报告/公报，正文已按“据该补充材料转引”表述。")
    A("")

    # ---------------- G ----------------
    A("## G. 数值核验与结果摘要")
    A("")
    A("### G.1 输入与规模")
    A("")
    A("| 项目 | 数值 |")
    A("|---|---|")
    A(f"| 第一步成果 | `{step1['path']}`（sha256 `{step1['sha256']}`，来源={step1['source']}） |")
    A(f"| 车辆退役序列 | {len(step1['veh'])} 条 × {len(step1['v_years'])} 年（{step1['v_years'][0]}—{step1['v_years'][-1]}） |")
    A(f"| 绿电退役序列 | {len(step1['pow'])} 条 × {len(step1['p_years'])} 年（{step1['p_years'][0]}—{step1['p_years'][-1]}） |")
    A(f"| 回收率记录（GCAM 32区） | {rec['n_rows']} 条（已排除“给定地区”行 {sum(rec['skipped'].values())} 条） |")
    A(f"| 正式比较期 | {years[0]}—{years[-1]}（{len(years)} 年） |")
    A("")
    A("### G.2 输入文件 SHA-256")
    A("")
    A("| 文件 | 读取方式 | SHA-256 |")
    A("|---|---|---|")
    for k, v in run_info["hashes"].items():
        A(f"| {v['path']} | {v['source']} | `{v['sha256']}` |")
    A("")
    A("说明：`数据.xlsx` 在运行期间被 Excel 独占打开（无法以只读方式直接读取），因此读取项目 `cache/` 目录下的只读副本。"
      "该副本由 Excel `SaveCopyAs` 从内存工作簿导出，导出时工作簿 `Saved=True`（内存状态与磁盘一致）；"
      "上表记录的是副本的 SHA-256，原始文件解冻后可用同一算法复核。其余输入均为原文件直读。")
    A("")
    A("### G.3 核验结果（本次运行）")
    A("")
    A("#### G.3.a 第一步核验（复读第一步「核验结果」工作表，全部通过才进入第二步）")
    A("")
    A("| 检查项目 | 检查数量 | 最大误差 | 是否通过 | 说明 |")
    A("|---|---|---|---|---|")
    for chk in step1.get("first_checks", []):
        err = "—" if chk["最大误差"] is None else f"{chk['最大误差']:.3e}"
        n = "—" if chk["检查数量"] is None else f"{chk['检查数量']}"
        note = str(chk["说明"]).replace("|", "/")
        A(f"| {chk['检查项目']} | {n} | {err} | {chk['是否通过']} | {note} |")
    A("")
    A("#### G.3.b 第二步核验（本次运行）")
    A("")
    A("| 检查项目 | 检查数量 | 最大误差 | 是否通过 | 说明 |")
    A("|---|---|---|---|---|")
    for chk in checks:
        err = "—" if chk["最大误差"] is None else f"{chk['最大误差']:.3e}"
        n = "—" if chk["检查数量"] is None else f"{chk['检查数量']}"
        note = str(chk["说明"]).replace("|", "/")
        A(f"| {chk['检查项目']} | {n} | {err} | {chk['是否通过']} | {note} |")
    A("")
    n_fail = sum(1 for c in checks if c["是否通过"] != "通过")
    A(f"未通过项：**{n_fail}**。")
    A("")

    A("### G.4 主要结果摘要（t Cu）")
    A("")
    A("| Scenario | 回收情景 | 贸易情景 | 2021 | 2030 | 2040 | 2050 | 2021—2050 累计 |")
    A("|---|---|---|---|---|---|---|---|")
    for scn in scn_list:
        for rs in REC_ORDER:
            for ts in TRADE_ORDER:
                arr = sum((flow["totals"][(scn, rs, ts, d)] for d in devices), np.zeros(flow["n_t"]))
                A(f"| {scn} | {rs} | {ts} | {fmt(arr[0])} | {fmt(arr[years.index(2030)])} | "
                  f"{fmt(arr[years.index(2040)])} | {fmt(arr[-1])} | {fmt(arr.sum())} |")
    A("")
    A("### G.5 分设备类别（2050年，基准×不开放 / 基准×全球开放）")
    A("")
    A("| 设备类别 | 基准×不开放 2050 | 基准×全球开放 2050 | 开放-不开放 |")
    A("|---|---|---|---|")
    for device in devices:
        a = float(flow["totals"][("SSP2-4.5", "基准", "不开放", device)][-1])
        b = float(flow["totals"][("SSP2-4.5", "基准", "全球开放", device)][-1])
        A(f"| {device} | {fmt(a)} | {fmt(b)} | {fmt(b - a)} |")
    A("")
    A("### G.5.1 水电（hydro）结果与口径变化")
    A("")
    A("本轮把水电从“铜存量驱动 + 容量→铜量份额再分配”改为“装机层面按 CON 份额拆分 + 投入年规模强度”："
      "内部按 small/medium/large 分别计算，对外只保留一条 hydro。下表给出本次 hydro 退役含铜量，"
      "并在可用时与旧口径基线对比。")
    A("")
    if baseline_hydro:
        A("| Scenario | 2021—2050 退役（本次） | 2021—2050（旧口径） | 变化 | 2050（本次） | 2050（旧口径） | 变化 |")
        A("|---|---:|---:|---:|---:|---:|---:|")
    else:
        A("| Scenario | 2021—2050 退役 | 2050 退役 |")
        A("|---|---:|---:|")
    hydro_years = [y for y in years]
    for scn in scn_list:
        cum = sum(step1["pow"][k2][y] or 0.0 for k2 in step1["pow"]
                  if k2[0] == scn and k2[3] == "hydro" for y in hydro_years if y >= 2021)
        v50 = sum(step1["pow"][k2][2050] or 0.0 for k2 in step1["pow"]
                  if k2[0] == scn and k2[3] == "hydro")
        if baseline_hydro:
            old = baseline_hydro["by_scenario"].get(scn, {"cum": 0.0, "v2050": 0.0})
            dcum = (cum - old["cum"]) / old["cum"] if old["cum"] else 0.0
            d50 = (v50 - old["v2050"]) / old["v2050"] if old["v2050"] else 0.0
            A(f"| {scn} | {fmt(cum)} | {fmt(old['cum'])} | {dcum:+.2%} | "
              f"{fmt(v50)} | {fmt(old['v2050'])} | {d50:+.2%} |")
        else:
            A(f"| {scn} | {fmt(cum)} | {fmt(v50)} |")
    A("")
    if baseline_hydro:
        A(f"基线文件：`{baseline_hydro['path']}`（sha256 `{baseline_hydro['sha256']}`，"
          "旧的铜存量驱动口径）。差异不要求为零：新口径不含“铜强度下降导致的存量缩水退役”，"
          "但新增/退役的时点与铜量都按投入年强度重算；因此水电退役量通常低于旧口径，"
          "差异同时反映时点与铜量口径的变化，不能解释为同一物理量的误差。")
        A("")
    A("若基线文件缺失，本节只报告本次结果；旧口径成果仍保留在 `cache/` 备份中。")
    A("")
    A("### G.6 面板规模")
    A("")
    A("| 工作表 | 行数 | 年份列 |")
    A("|---|---|---|")
    for name, n in counts.items():
        if name == "双边废铜_面板":
            A(f"| {name} | {n}（不开放 {bilat_count['不开放']} 条、开放 {bilat_count['全球开放']} 条"
              f"“路线×设备类别”，每条含两个流量类型） | {years[0]}—{years[-1]} |")
        else:
            A(f"| {name} | {n} | {years[0]}—{years[-1]} |")
    A("")
    A("### G.7 异常与影响说明")
    A("")
    A("- `保有量.xlsx` 车辆面板有 20 个空白节点（SSP2-4.5 与 SSP5-8.5 的 Central Asia 车辆 2020 年）。"
      "未填0；按相邻有效节点插值并登记，影响限于这 20 个节点。")
    A("- 源表设备组合本身不完整：车辆缺少 84 个（情景×地区×车型×动力）组合、绿电缺少 204 个（主要为 SSP5-8.5 无CCS）。"
      "这些组合不补造；SSP5-8.5 的CCS相关设备为结构性0，已在核验结果登记。")
    A(f"- 设备量下降触发的**经济性提前退役**：车辆 {fmt(run_info.get('early_vehicle', 0))} 辆、"
      f"绿电 {fmt(run_info.get('early_power', 0))} MW；已单列于第一步提前退役面板，"
      "不与自然退役混报，也不来自铜强度下降。")
    A("")
    A("### G.8 结论性说明")
    A("")
    A("核验通过只表示模型内部一致（质量平衡、贸易分配、参数范围、写盘一致），不表示科学结论已被外部验证。"
      "情景间的差值是给定参数与假设下的条件结果，不构成对现实政策的预测。")
    A("")

    # ---------------- H ----------------
    A("## H. 复现说明")
    A("")
    A("### H.1 运行命令")
    A("")
    A("```powershell")
    A("cd C:\\Users\\22965\\OneDrive\\桌面\\MFA\\回收情景模拟")
    A("# 1) 第一步写临时成果（不覆盖正式文件）")
    A("python step1_retired_copper.py --out ..\\Output\\_tmp_step1.xlsx")
    A("# 2) 第二步读取该临时第一步，写临时成果与临时 MD")
    A("python step2_secondary_copper.py --step1 ..\\Output\\_tmp_step1.xlsx "
      "--out ..\\Output\\_tmp_step2.xlsx --md ..\\Output\\_tmp_method.md")
    A("# 3) 两步核验通过后提升为正式成果（自动备份旧文件到 cache/backup_*）")
    A("python step2_secondary_copper.py --step1 ..\\Output\\_tmp_step1.xlsx "
      "--out ..\\Output\\_tmp_step2.xlsx --md ..\\Output\\_tmp_method.md --promote")
    A("```")
    A("")
    A("两个脚本均支持 `--config`；第一步 `--out` 指定临时路径，第二步 `--step1/--out/--md/--promote` "
      "分别指定输入、输出、MD 与是否提升。")
    A("")
    A("### H.2 环境")
    A("")
    A(f"- Python {run_info['python']}；numpy {run_info['numpy']}；openpyxl {run_info['openpyxl']}；XlsxWriter {run_info['xlsxwriter']}")
    A("- 未使用随机数，无随机种子；全部计算为确定性。")
    A("")
    A("### H.3 输入与输出")
    A("")
    A("| 类型 | 路径 |")
    A("|---|---|")
    for k, v in run_info["hashes"].items():
        A(f"| 输入 | {v['path']}（{v['source']}） |")
    A(f"| 输出-第一步 | {cfg['paths']['step1_out']} |")
    A(f"| 输出-第二步 | {cfg['paths']['step2_out']} |")
    A(f"| 输出-本MD | {cfg['paths']['method_md']} |")
    A("")
    A("面板格式：年份横向展开（整数列，升序），一行对应唯一索引组合；铜量单位 t Cu，参数为0—1比例。")
    A("")
    A("### H.4 运行时间")
    A("")
    A(f"- 第一步：{run_info['t1']:.1f} s（本次记录）")
    A(f"- 第二步：{run_info['t2']:.1f} s（本次记录）")
    A("")
    A("### H.5 适用范围与局限")
    A("")
    A("- 处理能力为二值设定（15区有/17区无），不设产能上限、库存、暂存与运输损失。")
    A("- 贸易矩阵为静态矩阵（2021—2050 同一年份结构），且由光伏研究外推到全部设备类别。")
    A("- 回收率参数以2021年校准、2060年为目标做线性提升，属情景设定而非预测。")
    A("- 水电采用 CON 规模结构并以中国规模结构外推到全球：小型份额按 76.1148/93/100 GW 对 380.574 GW 计算，"
      "大型内部固定 49:15；这是代理假设，不是各区域实际规模构成，也不是论文给出的未来固定内部比例。")
    A("- 设备量驱动：铜强度只作用于投入批次，退役不再因强度下降而增加；"
      "模型铜存量按投入年强度累加，与 `铜存量.xlsx` 的“当年强度重估”分叉属预期，已在核对中说明。")
    A("- 经济性提前退役只按缺口扣减（不是整批退役），其规模取决于保有量/装机路径的下降幅度，属情景输入的结果。")
    A("- 不开展运输成本、再生铜返还、原生铜替代、水压力与成本核算。")
    A("")
    A("---")
    A("")
    A("本文件仅保存在 `MFA/Output`，不在代码目录保留副本；代码目录 README.md 指向上表路径。")
    A("")
    return "\n".join(L), n_fail


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main(argv=None):
    parser = argparse.ArgumentParser(description="第二步：次生铜产量")
    parser.add_argument("--config", default=None)
    parser.add_argument("--step1", default=None, help="第一步成果路径（默认取配置；可为临时成果）")
    parser.add_argument("--out", default=None, help="第二步输出路径（默认取配置）")
    parser.add_argument("--md", default=None, help="方法MD输出路径（默认取配置）")
    parser.add_argument("--promote", action="store_true",
                        help="两步核验通过后，备份旧正式成果并提升为正式文件")
    args = parser.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    started = time.time()
    cfg = load_config(args.config)
    step1_path = Path(args.step1) if args.step1 else Path(cfg["paths"]["step1_out"])
    out_path = Path(args.out) if args.out else Path(cfg["paths"]["step2_out"])
    md_path = Path(args.md) if args.md else Path(cfg["paths"]["method_md"])

    log("=" * 100)
    log("第二步：三回收情景 × 两贸易情景的逐年次生铜产量")
    log(f"配置：{cfg['_config_path']}（sha256={cfg['_config_sha256']}）")
    log(f"Python：{sys.version.split()[0]}；numpy={np.__version__}；openpyxl={openpyxl.__version__}")

    step1 = load_step1(cfg, step1_path)
    log(f"[核对] 第一步核验结果全部通过；车辆 {len(step1['veh'])} 条、绿电 {len(step1['pow'])} 条序列")

    baseline_hydro = load_baseline_hydro(cfg)
    if baseline_hydro:
        log(f"[基线] 旧口径成果：{baseline_hydro['path']}（sha256={baseline_hydro['sha256']}）")
    else:
        log("[基线] 未找到旧口径基线文件，MD 只报告本次结果")

    rec = load_recycling(cfg)
    log(f"[输入] 回收率：{rec['path']}（{rec['n_rows']} 条 GCAM 32区记录；"
        f"排除行={sum(rec['skipped'].values())}）sha256={rec['sha256']}")
    trade_data = load_trade(cfg)
    log(f"[输入] 贸易矩阵：{trade_data['path']}（{trade_data['matrix'].shape[0]}×{trade_data['matrix'].shape[1]}）"
        f"sha256={trade_data['sha256']}")

    # 面板区域全集
    panel_regions = sorted({k[1] for k in step1["pow"]} | {k[1] for k in step1["veh"]})
    if len(panel_regions) != 32:
        raise ValueError(f"区域数不为32：{len(panel_regions)}")
    cfg["_panel_regions"] = panel_regions
    alias_regions = {r for r in panel_regions if r in cfg["region_aliases"]}
    log(f"[区域] 32区；显式别名 {len(alias_regions)} 个：{sorted(alias_regions)}")

    r3 = build_r3(cfg, rec["rates"], step1)
    if r3["missing"]:
        raise ValueError(f"回收率缺少基准参数：{r3['missing'][:10]}")
    log(f"[参数] 回收率/处理率序列 {len(r3['C']['基准'])} 组；年份 {r3['years'][0]}—{r3['years'][-1]}")

    flow = compute_flows(cfg, step1, r3, trade_data, panel_regions)
    log("[计算] 18个组合已生成；开始核验")

    checks = verify_all(cfg, rec, trade_data, flow, r3, step1)
    for c in checks:
        log(f"  [{c['是否通过']}] {c['检查项目']} | 数量={c['检查数量']} | 误差={c['最大误差']} | {c['说明']}")
    n_fail = sum(1 for c in checks if c["是否通过"] != "通过")
    if n_fail:
        log(f"核验未通过 {n_fail} 项，停止：不生成正式成果与MD")
        write_run_log(cfg, started)
        raise SystemExit(1)

    tmp, counts, bilat_count = build_workbook(cfg, rec, trade_data, flow, r3, step1, checks)
    wb_checks = verify_workbook(tmp, flow, cfg)
    for c in wb_checks:
        log(f"[写盘核对] {c['检查项目']}：{c['是否通过']} | {c['说明']}")
    checks.extend(wb_checks)
    n_fail2 = sum(1 for c in checks if c["是否通过"] != "通过")
    if n_fail2:
        log(f"写盘核对未通过 {n_fail2} 项，停止；临时文件：{tmp}")
        write_run_log(cfg, started)
        raise SystemExit(1)

    # 把写盘核对写入正式成果的「核验结果」表，再重建一次并复核（保证 Excel 与 MD 的核验条目一致）
    tmp2, counts2, bilat2 = build_workbook(cfg, rec, trade_data, flow, r3, step1, checks)
    wb_checks2 = verify_workbook(tmp2, flow, cfg)
    key = lambda items: [(i["检查项目"], i["是否通过"], i["检查数量"]) for i in items]
    if key(wb_checks2) != key(wb_checks):
        log("重建后的写盘核对结果与首次不一致，停止替换")
        write_run_log(cfg, started)
        raise SystemExit(1)
    for item in wb_checks2:
        log(f"[写盘复核] {item['检查项目']}：{item['是否通过']} | {item['说明']}")
    shutil.rmtree(Path(tmp).parent, ignore_errors=True)
    tmp, counts, bilat_count = tmp2, counts2, bilat2

    # 写入目标（临时或正式）
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.replace(tmp, out_path)
    except PermissionError:
        log(f"替换失败：{out_path} 可能正被 Excel 打开；临时文件保留：{tmp}")
        write_run_log(cfg, started)
        raise
    shutil.rmtree(Path(tmp).parent, ignore_errors=True)
    log(f"第二步成果已写入：{out_path}")
    for name, n in counts.items():
        log(f"  {name}: {n} 行")

    # 生成 MD（核验通过后）
    run_info = {
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "git_head": git_head(),
        "python": sys.version.split()[0],
        "numpy": np.__version__,
        "openpyxl": openpyxl.__version__,
        "xlsxwriter": xlsxwriter.__version__,
        "t1": 0.0,
        "t2": time.time() - started,
        "hashes": {},
        "cf_backfill": 0,
        "n_extra": 0,
        "cipg_diag": 0.0, "cipg_diag_pv": 0.0, "cipg_diag_hydro": 0.0,
    }
    # 补充第一步的运行信息
    for key, path in (("保有量(设备路径)", cfg["paths"]["stock_xlsx"]),
                      ("铜存量(仅核对)", cfg["paths"]["copper_stock_xlsx"]),
                      ("参数", cfg["paths"]["data_xlsx"]),
                      ("回收率", cfg["paths"]["recycling_xlsx"]),
                      ("贸易矩阵", cfg["paths"]["trade_xlsx"])):
        p = Path(path)
        try:
            digest = sha256_of(p)
            src = "original"
        except PermissionError:
            cached = Path(cfg["paths"]["fallback_dir"]) / p.name
            digest = sha256_of(cached)
            src = "cache_copy"
            p = cached
        run_info["hashes"][key] = {"path": str(path), "source": src, "sha256": digest}
    run_info["hashes"]["第一步成果"] = {
        "path": cfg["paths"]["step1_out"],
        "source": ("original" if step1_path.resolve() == Path(cfg["paths"]["step1_out"]).resolve()
                   else f"temp_file:{step1_path.name}"),
        "sha256": sha256_of(step1_path),
    }
    run_info.update(read_step1_runinfo(cfg))

    md, _ = build_markdown(cfg, rec, trade_data, flow, r3, step1, checks, counts, bilat_count,
                           run_info, baseline_hydro)
    md_path.parent.mkdir(parents=True, exist_ok=True)
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write(md)
    log(f"方法与文献说明：{md_path}（{len(md.splitlines())} 行）")

    # 提升为正式成果（仅两步核验全部通过后）
    if args.promote:
        official_step1 = Path(cfg["paths"]["step1_out"])
        official_step2 = Path(cfg["paths"]["step2_out"])
        official_md = Path(cfg["paths"]["method_md"])
        backup = Path(cfg["paths"]["fallback_dir"]) / f"backup_{time.strftime('%Y%m%d_%H%M%S')}"
        backup.mkdir(parents=True, exist_ok=True)
        for src in (official_step1, official_step2, official_md):
            if src.exists():
                shutil.copy2(src, backup / src.name)
        log(f"[提升] 旧正式成果已备份：{backup}")
        if step1_path.resolve() != official_step1.resolve():
            shutil.copy2(step1_path, official_step1)
            log(f"[提升] 第一步：{step1_path} → {official_step1}")
        if out_path.resolve() != official_step2.resolve():
            shutil.copy2(out_path, official_step2)
            log(f"[提升] 第二步：{out_path} → {official_step2}")
        if md_path.resolve() != official_md.resolve():
            shutil.copy2(md_path, official_md)
            log(f"[提升] 方法MD：{md_path} → {official_md}")
    log(f"耗时：{time.time() - started:.1f} s")
    write_run_log(cfg, started)
    return 0


def read_step1_runinfo(cfg):
    """从第一步运行记录中提取补充信息（缺失则返回空）。"""
    info = {}
    log_path = Path(cfg["paths"]["output_dir"]) / "运行记录_第一步.txt"
    if not log_path.exists():
        return info
    text = log_path.read_text(encoding="utf-8", errors="replace")
    for line in text.splitlines():
        if "提前退役：车辆=" in line:
            try:
                v = line.split("车辆=")[1].split(" 辆")[0].replace(",", "")
                p = line.split("绿电=")[1].split(" MW")[0].replace(",", "")
                info["early_vehicle"] = float(v)
                info["early_power"] = float(p)
            except Exception:
                pass
        if "耗时" in line and "s" in line:
            try:
                info["t1"] = float(line.split("耗时")[1].split("s")[0].replace("：", "").strip())
            except Exception:
                pass
    return info


def write_run_log(cfg, started):
    out_dir = Path(cfg["paths"]["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = out_dir / "运行记录_第二步.txt"
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(RUN_LOG))
        fh.write(f"\n运行结束：{time.strftime('%Y-%m-%d %H:%M:%S')}；耗时 {time.time() - started:.1f} s\n")
    print(f"运行记录：{log_path}")


if __name__ == "__main__":
    sys.exit(main())
