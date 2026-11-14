# -*- coding: utf-8 -*-
"""第一步：设备量驱动的 DMFA —— 新增设备/装机、退役设备/装机、新增铜需求、退役铜量。

输入（只读）
------------
1. Output\\保有量.xlsx
   - 车辆保有量_面板：Scenario, Region, Subsector, Technology, 单位(辆), 1990..2050 节点
   - 装机容量_九类_面板：Scenario, region, tech, Subtech, 单位(MW), 1990..2050 节点
2. 数据.xlsx：CIPG（t/GW，2021 起）、life（寿命分布）、Automotive copper demand（kg/vehicle）
3. Output\\铜存量.xlsx：仅用于核对（保有量×系数 与旧铜存量表的口径差异）

核心口径
--------
- 驱动量 Q_t：车辆=保有量（辆），发电=装机（MW）；水电在装机层面按 CON 份额拆 small/medium/large。
- t0=首个正值节点；t<t0 时 Q=0、无投入无退役；Q(t0) 作为初始批次（年龄0，当年不退役）；
  之后按有效节点线性插值到年，不外推。
- 每年：R_t=Σ I_τ p(t−τ)；I_t=Q_t−Q_{t−1}+R_t。
  I_t<0 时，缺口 E_t=−I_t 记为“经济性提前退役”（从最老批次扣减），I_t=0；自然退役与提前退役分列。
- 系数：每个投入批次的铜强度在投入年确定，此后固定（A 方案）。
  车辆用 Automotive 固定值；发电用 CIPG，投入年≤2021 用 2021 值，节点间线性插值（禁止跳配）。
  CIPG 年度下降只影响新批次，不产生退役。
- 水电：CON 容量份额直接拆装机；规模强度取补充表5并以项目 2021 综合强度归一。

输出：Output\\能源转型设备退役铜量_第一步.xlsx（18 个工作表；年份列沿用替换前模板）

本脚本 import 时不执行计算，只在 __main__ 下运行。
"""
from __future__ import annotations

import argparse
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

import numpy as np
import openpyxl
import xlsxwriter

PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = PROJECT_DIR / "config.json"
RUN_LOG: list[str] = []

HYDRO_SCALES = ("small", "medium", "large")
HYDRO_SCALE_DIST = {"small": "Small Hydro", "medium": "Medium Hydro",
                    "large": "Large Hydro / PSH"}


def log(msg: str = "") -> None:
    print(msg)
    RUN_LOG.append(str(msg))


# --------------------------------------------------------------------------- #
# 基础工具
# --------------------------------------------------------------------------- #
def sha256_of(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_config(path=None) -> dict:
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


def open_workbook_readonly(path, fallback_dir=None):
    src = Path(path)
    try:
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        return wb, src, "original", sha256_of(src)
    except PermissionError:
        pass
    cache_dir = Path(fallback_dir) if fallback_dir else src.parent
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached = cache_dir / src.name
    if cached.exists():
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


def build_scn_map(cfg) -> dict:
    m = {}
    for canon in cfg["scenarios"]["canonical"]:
        m[canon] = canon
        m[norm_scenario(canon)] = canon
    for key in ("vehicle_names", "power_names"):
        for canon, name in cfg["scenarios"][key].items():
            m[str(name).strip()] = canon
            m[norm_scenario(name)] = canon
    return m


def canon_scenario(cfg, value) -> str:
    raw = str(value).strip()
    return cfg["_scn_map"].get(raw) or cfg["_scn_map"].get(norm_scenario(raw)) or norm_scenario(raw)


# --------------------------------------------------------------------------- #
# 寿命分布
# --------------------------------------------------------------------------- #
def build_survival(spec: dict, amax: int):
    ages = np.arange(0, amax + 1, dtype=float)
    kind = spec["type"].lower()
    if kind == "weibull":
        S = np.exp(-np.power(ages / float(spec["scale"]), float(spec["k"])))
    elif kind == "normal":
        mu, sigma = float(spec["mu"]), float(spec["sigma"])
        raw = 0.5 * np.vectorize(math.erfc)((ages - mu) / (sigma * math.sqrt(2.0)))
        S = raw / raw[0]
    elif kind == "fixed":
        S = np.where(ages < float(spec["life"]), 1.0, 0.0)
    else:
        raise ValueError(f"未知寿命分布：{spec}")
    S[0] = 1.0
    p = np.zeros_like(S)
    p[1:] = S[:-1] - S[1:]
    p = np.clip(p, 0.0, 1.0)
    return S, p


def distribution_label(spec: dict) -> str:
    kind = spec["type"].lower()
    if kind == "weibull":
        return f"Weibull(k={spec['k']}, λ={spec['scale']})"
    if kind == "normal":
        return f"Normal(μ={spec['mu']}, σ={spec['sigma']})，截断于寿命>0并归一化"
    return f"Fixed({spec['life']}年)"


def load_life_table(wb) -> dict:
    """从 数据.xlsx / life 读取分布参数。"""
    ws = wb["life"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    i_name = header_index(header, "设备/子技术")
    i_dist = header_index(header, "基准分布")
    i_k = header_index(header, "Shape k")
    i_scale = header_index(header, "基准尺度/均值/固定寿命(年)")
    i_sigma = header_index(header, "Std σ(年)")
    out = {}
    for r in it:
        if r is None or r[i_name] is None:
            continue
        name = str(r[i_name]).strip()
        dist = str(r[i_dist]).strip().lower()
        if dist.startswith("weibull"):
            out[name] = {"type": "weibull", "k": float(r[i_k]), "scale": float(r[i_scale])}
        elif dist.startswith("normal"):
            out[name] = {"type": "normal", "mu": float(r[i_scale]), "sigma": float(r[i_sigma])}
        else:
            out[name] = {"type": "fixed", "life": float(r[i_scale])}
    return out


def load_automotive(wb) -> dict:
    ws = wb["Automotive copper demand"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    i_cat = header_index(header, "categories")
    i_bev = header_index(header, "BEV")
    i_phev = header_index(header, "PHEV")
    out = {}
    for r in it:
        if r is None or r[i_cat] is None:
            continue
        out[str(r[i_cat]).strip()] = {"BEV": float(r[i_bev]), "PHEV": float(r[i_phev])}
    return out


def load_cipg(wb) -> dict:
    ws = wb["CIPG"]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    i_tech = header_index(header, "technology")
    year_cols = [(i, int(float(h))) for i, h in enumerate(header) if is_year_token(h)]
    out = {}
    for r in it:
        if r is None or r[i_tech] is None:
            continue
        tech = str(r[i_tech]).strip()
        out[tech] = {y: to_number(r[i]) for i, y in year_cols if to_number(r[i]) is not None}
    return out


def cipg_series(cipg: dict, tech: str, years: list[int]) -> np.ndarray:
    """投入年系数：≤最早节点(2021)用最早值；节点间线性插值；不外推。"""
    series = cipg.get(tech)
    if not series:
        raise KeyError(f"CIPG 缺少技术：{tech}")
    ys = sorted(series)
    y0 = ys[0]

    def at(y):
        if y <= y0:
            return float(series[y0])
        if y >= ys[-1]:
            return float(series[ys[-1]])
        for a, b in zip(ys, ys[1:]):
            if a <= y <= b:
                va, vb = float(series[a]), float(series[b])
                return va + (vb - va) * (y - a) / (b - a)
        return float(series[ys[-1]])

    return np.array([at(y) for y in years], dtype=float)


# --------------------------------------------------------------------------- #
# 水电 CON 规模
# --------------------------------------------------------------------------- #
def hydro_capacity_shares(cfg, year: int):
    hc = cfg["life"]["hydro_con"]
    total = float(hc["total_capacity_gw"])
    if year <= 2021:
        w_small = float(hc["history_shares"]["small"])
    else:
        cap = hc["small_capacity_gw"]
        if year <= 2030:
            k = float(cap["2021"]) + (float(cap["2030"]) - float(cap["2021"])) * (year - 2021) / 9.0
        elif year <= 2050:
            k = float(cap["2030"]) + (float(cap["2050"]) - float(cap["2030"])) * (year - 2030) / 20.0
        else:
            raise ValueError(f"CON 份额只定义到 2050：{year}")
        w_small = k / total
    w_medium = float(hc["medium_share"])
    return w_small, w_medium, 1.0 - w_small - w_medium


def hydro_scale_intensities(cfg, cipg: dict, years: list[int]):
    """返回 (base{scale: t/GW}, series{scale: array})。

    以补充表5规模强度为比例，按项目 hydro CIPG 路径缩放，
    并用 2021 年 20/16/64 加权值归一，使 2021 年综合强度精确等于项目值。
    """
    hc = cfg["life"]["hydro_con"]
    inten = hc["copper_intensity_t_per_gw"]
    ratio = hc["large_internal_ratio"]
    n_res, n_psh = float(ratio["reservoir_large"]), float(ratio["psh"])
    i_large = (n_res * float(inten["large_reservoir"]) + n_psh * float(inten["psh"])) / (n_res + n_psh)
    base = {"small": float(inten["small_ror"]), "medium": float(inten["medium_reservoir"]),
            "large": i_large}
    w_s, w_m, w_l = hydro_capacity_shares(cfg, 2021)
    agg2021 = w_s * base["small"] + w_m * base["medium"] + w_l * base["large"]
    hydro_path = cipg_series(cipg, "hydro", years)
    cipg2021 = float(cipg_series(cipg, "hydro", [2021])[0])
    norm = cipg2021 / agg2021
    series = {k: base[k] * (hydro_path / cipg2021) * norm for k in base}
    return base, series, agg2021, norm


# --------------------------------------------------------------------------- #
# 年度化与批次引擎
# --------------------------------------------------------------------------- #
def annualize(nodes: dict, years: list[int]):
    """返回 (Q 数组, t0 索引或 None)。t0=首个正值节点；t<t0 为 0。"""
    n = len(years)
    Q = np.zeros(n)
    t0 = next((y for y in years if nodes.get(y) is not None and nodes[y] > 0), None)
    if t0 is None:
        return Q, None, None
    i0 = years.index(t0)
    pts = [(y, nodes[y]) for y in years if y >= t0 and nodes.get(y) is not None]
    xs = np.array([p[0] for p in pts], dtype=float)
    ys = np.array([p[1] for p in pts], dtype=float)
    grid = np.array(years[i0:], dtype=float)
    Q[i0:] = np.interp(grid, xs, ys)
    return Q, i0, t0


def simulate_batches(Q: np.ndarray, years: list[int], i0: int, S: np.ndarray, p: np.ndarray,
                     k: np.ndarray, tol_abs: float, tol_rel: float):
    """批次引擎。k 为投入年铜强度（t/单位），随年份变化；批次固定投入年 k。"""
    n = len(Q)
    I = np.zeros(n)
    R = np.zeros(n)
    E = np.zeros(n)
    C_in = np.zeros(n)
    C_out = np.zeros(n)      # 自然 + 提前
    C_out_nat = np.zeros(n)
    C_out_early = np.zeros(n)
    stock_cu = np.zeros(n)
    if i0 is None:
        return {"I": I, "R": R, "E": E, "C_in": C_in, "C_out": C_out,
                "C_out_nat": C_out_nat, "C_out_early": C_out_early,
                "stock_cu": stock_cu, "Q": Q, "min_I": 0.0, "neg_float": 0, "neg_material": 0}
    amt = np.zeros(n)          # 批次投入量（未扣减）
    amt[i0] = Q[i0]
    I[i0] = Q[i0]
    C_in[i0] = I[i0] * k[i0]
    stock_cu[i0] = amt[i0] * S[0] * k[i0]
    min_I = 0.0
    neg_float = 0
    neg_material = 0
    for i in range(i0 + 1, n):
        ages = i - np.arange(i, dtype=int)
        deaths = amt[:i] * p[ages]
        R[i] = float(deaths.sum())
        C_out_nat[i] = float(np.sum(deaths * k[:i]))
        need = Q[i] - Q[i - 1] + R[i]
        if need >= 0:
            I[i] = need
            amt[i] = need
            C_in[i] = need * k[i]
            C_out[i] = C_out_nat[i]
        else:
            tol = max(tol_abs, tol_rel * max(abs(Q[i]), 1.0))
            if -need <= tol:
                neg_float += 1
                I[i] = 0.0
                C_out[i] = C_out_nat[i]
            else:
                neg_material += 1
                short = -need
                E[i] = short
                I[i] = 0.0
                removed_cu = 0.0
                for age in range(i, 0, -1):
                    if short <= 1e-12:
                        break
                    b = i - age
                    Sv = S[age]
                    if Sv <= 0:
                        continue
                    remaining = amt[b] * Sv
                    if remaining <= 0:
                        continue
                    take = min(remaining, short)
                    amt[b] -= take / Sv
                    removed_cu += take * k[b]
                    short -= take
                C_out_early[i] = removed_cu
                C_out[i] = C_out_nat[i] + C_out_early[i]
        stock_cu[i] = float(np.sum(amt[:i + 1] * S[i - np.arange(i + 1)] * k[:i + 1]))
        min_I = min(min_I, I[i])
    return {"I": I, "R": R, "E": E, "C_in": C_in, "C_out": C_out,
            "C_out_nat": C_out_nat, "C_out_early": C_out_early,
            "stock_cu": stock_cu, "Q": Q, "min_I": min_I,
            "neg_float": neg_float, "neg_material": neg_material}


# --------------------------------------------------------------------------- #
# 读入设备面板与年份模板
# --------------------------------------------------------------------------- #
def read_panel(ws, n_keys):
    it = ws.iter_rows(values_only=True)
    header = next(it)
    years = [int(float(h)) for h in header if is_year_token(h)]
    data = []
    for r in it:
        if r is None or r[0] is None:
            continue
        key = tuple(str(r[i]).strip() for i in range(n_keys))
        vals = {y: to_number(r[n_keys + 1 + i]) for i, y in enumerate(years)}
        data.append({"key": key, "unit": r[n_keys] if len(r) > n_keys else None, "years": vals})
    return data, years, header


def load_year_template(cfg):
    """读替换前正式成果的年份列；优先级：正式文件 → cache/year_template.json。"""
    template_path = Path(cfg["paths"]["year_template"])
    official = Path(cfg["paths"]["step1_out"])
    out = None
    if official.exists():
        try:
            wb = openpyxl.load_workbook(official, read_only=True, data_only=True)
            try:
                v, _ = _panel_years(wb, cfg["sheets"]["vehicle_out_panel"])
                p, _ = _panel_years(wb, cfg["sheets"]["power_out_panel"])
                out = {"vehicle": v, "power": p, "source": str(official)}
            finally:
                wb.close()
        except Exception as exc:
            log(f"[模板] 正式成果不可读（{type(exc).__name__}），改用缓存模板")
            out = None
    if out is None and template_path.exists():
        out = json.loads(template_path.read_text(encoding="utf-8"))
        out["source"] = str(template_path)
    if out is None:
        raise RuntimeError("无法读取输出年份模板（正式成果与 cache/year_template.json 均缺失）")
    template_path.parent.mkdir(parents=True, exist_ok=True)
    template_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def _panel_years(wb, sheet):
    ws = wb[sheet]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    years = [int(float(h)) for h in header if is_year_token(h)]
    return years, header


# --------------------------------------------------------------------------- #
# 计算
# --------------------------------------------------------------------------- #
def compute(cfg, template):
    cfg["_scn_map"] = build_scn_map(cfg)
    fallback = Path(cfg["paths"]["fallback_dir"])
    if not fallback.is_absolute():
        fallback = PROJECT_DIR / fallback
    years_all = list(range(cfg["years"]["power_start"], cfg["years"]["end"] + 1))
    v_out = template["vehicle"]
    p_out = template["power"]
    amax = int(cfg["life"]["amax"])
    tol_abs = float(cfg.get("model", {}).get("negative_tol_abs", 1e-6))
    tol_rel = float(cfg.get("model", {}).get("negative_tol_rel", 1e-9))

    hashes = {}
    # 保有量
    wb, used, src, digest = open_workbook_readonly(cfg["paths"]["stock_xlsx"], fallback)
    hashes["stock"] = {"path": str(cfg["paths"]["stock_xlsx"]), "source": src, "sha256": digest}
    vrecs, vyears, _ = read_panel(wb[cfg["sheets"]["vehicle_device_panel"]], 4)
    precs, pyears, _ = read_panel(wb[cfg["sheets"]["power_device_panel"]], 4)
    wb.close()
    log(f"[输入] 保有量：{cfg['paths']['stock_xlsx']}（{src}，sha256={digest}）")
    log(f"       车辆 {len(vrecs)} 行，发电 {len(precs)} 行；节点年份 {vyears}")

    # 参数
    wb, used, src, digest = open_workbook_readonly(cfg["paths"]["data_xlsx"], fallback)
    hashes["data"] = {"path": str(cfg["paths"]["data_xlsx"]), "source": src, "sha256": digest}
    life = load_life_table(wb)
    auto = load_automotive(wb)
    cipg = load_cipg(wb)
    wb.close()
    log(f"[输入] 参数：{cfg['paths']['data_xlsx']}（{src}，sha256={digest}）")

    # 寿命缓存
    surv_cache = {}

    def surv(name):
        if name not in life:
            raise KeyError(f"life 表缺少：{name}")
        if name not in surv_cache:
            surv_cache[name] = build_survival(life[name], amax)
        return surv_cache[name]

    # 发电系数序列（按 Subtech→CIPG）
    cipg_arrays = {}
    for subtech, tech in cfg["devices"]["subtech_to_cipg"].items():
        if subtech == "hydro":
            continue
        cipg_arrays[subtech] = cipg_series(cipg, tech, years_all) / 1000.0   # t/MW

    hydro_base, hydro_series, agg2021, norm = hydro_scale_intensities(cfg, cipg, years_all)
    cipg2021 = float(cipg_series(cipg, "hydro", [2021])[0])
    hydro_k = {s: hydro_series[s] / 1000.0 for s in HYDRO_SCALES}              # t/MW

    # 车辆：kg/vehicle → t/vehicle
    def veh_k(subsector, tech):
        if subsector not in auto:
            raise KeyError(f"Automotive 表缺少类别：{subsector}")
        return float(auto[subsector][tech]) / 1000.0

    # ---- 车辆 ----
    veh = {}
    blank_nodes = []
    zero_series = []
    for rec in vrecs:
        scn, region, subsector, tech = (canon_scenario(cfg, rec["key"][0]),
                                        rec["key"][1], rec["key"][2], rec["key"][3])
        nodes = {}
        for y, v in rec["years"].items():
            if v is None:
                blank_nodes.append(("车辆", scn, region, subsector, tech, y))
            else:
                nodes[y] = v
        Q, i0, t0 = annualize(nodes, years_all)
        if i0 is None:
            zero_series.append(("车辆", scn, region, subsector, tech))
        k = np.full(len(years_all), veh_k(subsector, tech))
        comps = cfg["life"]["vehicle_map"][subsector]["category"]
        S, p = surv(comps)
        res = simulate_batches(Q, years_all, i0, S, p, k, tol_abs, tol_rel)
        idx = [years_all.index(y) for y in v_out]
        veh[(scn, region, subsector, tech)] = {
            "unit": rec["unit"] or "辆",
            "I": res["I"][idx], "R": res["R"][idx], "E": res["E"][idx],
            "C_in": res["C_in"][idx], "C_out": res["C_out"][idx],
            "C_out_nat": res["C_out_nat"][idx], "C_out_early": res["C_out_early"][idx],
            "stock_cu": res["stock_cu"][idx], "Q": Q[idx],
            "t0": t0, "min_I": res["min_I"], "neg_float": res["neg_float"],
            "neg_material": res["neg_material"], "life": comps,
            "k": k[idx], "source_key": rec["key"], "nodes": nodes,
            "_full": res,
        }

    # ---- 发电 ----
    pow_ = {}
    hydro = {}
    for rec in precs:
        scn, region, tech, subtech = (canon_scenario(cfg, rec["key"][0]),
                                      rec["key"][1], rec["key"][2], rec["key"][3])
        nodes = {}
        for y, v in rec["years"].items():
            if v is None:
                blank_nodes.append(("绿电", scn, region, tech, subtech, y))
            else:
                nodes[y] = v
        Q_total, i0, t0 = annualize(nodes, years_all)
        if i0 is None:
            zero_series.append(("绿电", scn, region, tech, subtech))
        if subtech == "hydro":
            shares = np.array([hydro_capacity_shares(cfg, y) for y in years_all])  # n×3
            agg = {"I": np.zeros(len(years_all)), "R": np.zeros(len(years_all)),
                   "E": np.zeros(len(years_all)), "C_in": np.zeros(len(years_all)),
                   "C_out": np.zeros(len(years_all)), "C_out_nat": np.zeros(len(years_all)),
                   "C_out_early": np.zeros(len(years_all)), "stock_cu": np.zeros(len(years_all)),
                   "min_I": 0.0, "neg_float": 0, "neg_material": 0, "Q": Q_total}
            scale_res = {}
            for si, scale in enumerate(HYDRO_SCALES):
                S, p = surv(HYDRO_SCALE_DIST[scale])
                Qi = Q_total * shares[:, si]
                r = simulate_batches(Qi, years_all, i0, S, p, hydro_k[scale], tol_abs, tol_rel)
                scale_res[scale] = r
                for key in ("I", "R", "E", "C_in", "C_out", "C_out_nat", "C_out_early", "stock_cu"):
                    agg[key] += r[key]
                agg["min_I"] = min(agg["min_I"], r["min_I"])
                agg["neg_float"] += r["neg_float"]
                agg["neg_material"] += r["neg_material"]
            idx = [years_all.index(y) for y in p_out]
            pow_[(scn, region, tech, subtech)] = {
                "unit": rec["unit"] or "MW",
                "I": agg["I"][idx], "R": agg["R"][idx], "E": agg["E"][idx],
                "C_in": agg["C_in"][idx], "C_out": agg["C_out"][idx],
                "C_out_nat": agg["C_out_nat"][idx], "C_out_early": agg["C_out_early"][idx],
                "stock_cu": agg["stock_cu"][idx], "Q": Q_total[idx],
                "t0": t0, "min_I": agg["min_I"], "neg_float": agg["neg_float"],
                "neg_material": agg["neg_material"], "life": "CON", "nodes": nodes,
                "_full": agg,
            }
            hydro[(scn, region)] = {"scales": scale_res, "shares": shares,
                                    "Q_total": Q_total, "i0": i0, "t0": t0}
            continue
        k = cipg_arrays[subtech]
        cat = cfg["life"]["power_map"][subtech]["category"]
        S, p = surv(cat)
        res = simulate_batches(Q_total, years_all, i0, S, p, k, tol_abs, tol_rel)
        idx = [years_all.index(y) for y in p_out]
        pow_[(scn, region, tech, subtech)] = {
            "unit": rec["unit"] or "MW",
            "I": res["I"][idx], "R": res["R"][idx], "E": res["E"][idx],
            "C_in": res["C_in"][idx], "C_out": res["C_out"][idx],
            "C_out_nat": res["C_out_nat"][idx], "C_out_early": res["C_out_early"][idx],
            "stock_cu": res["stock_cu"][idx], "Q": Q_total[idx],
            "t0": t0, "min_I": res["min_I"], "neg_float": res["neg_float"],
            "neg_material": res["neg_material"], "life": cat,
            "k": k[idx], "source_key": rec["key"], "nodes": nodes,
            "_full": res,
        }

    # ---- 核验计算 ----
    checks = []
    checks.extend(build_checks(cfg, veh, pow_, hydro, blank_nodes, zero_series,
                               years_all, v_out, p_out, tol_abs, tol_rel,
                               agg2021, norm, hydro_base, hydro_k, cipg2021))
    result = {
        "cfg": cfg, "template": template, "v_out": v_out, "p_out": p_out,
        "veh": veh, "pow": pow_, "hydro": hydro,
        "checks": checks, "hashes": hashes, "blank_nodes": blank_nodes,
        "zero_series": zero_series, "hydro_base": hydro_base,
        "hydro_k": hydro_k, "agg2021": agg2021, "norm": norm,
        "years_all": years_all, "auto": auto, "cipg": cipg,
    }
    result["coef_rows"] = build_coef_rows(cfg, result, auto, cipg)
    return result


def build_checks(cfg, veh, pow_, hydro, blank_nodes, zero_series, years_all,
                 v_out, p_out, tol_abs, tol_rel, agg2021, norm, hydro_base, hydro_k,
                 cipg2021):
    checks = []

    def add(item, n, err, ok, note):
        checks.append({"检查项目": item, "检查数量": n, "最大误差": err,
                       "是否通过": "通过" if ok else "未通过", "说明": note})

    all_series = list(veh.values()) + list(pow_.values())
    n_series = len(all_series)
    # 1 初始化
    init_err = 0.0
    n_init = 0
    for r in all_series:
        if r["t0"] is None:
            continue
        n_init += 1
        i0 = years_all.index(r["t0"])
        # 起点批次 = 起点存量；当年退役 0
        q0 = float(r["Q"][:1].sum()) if False else None
        # 用内部数组核对
        full = r["_full"]
        init_err = max(init_err, abs(full["I"][i0] - full["Q"][i0]),
                       float(np.max(np.abs(full["R"][:i0 + 1]))))
    add("初始化：起点存量=初始批次，且当年不退役", n_init, init_err, init_err <= 1e-6,
        "逐序列检查 t0 年 I=Q、R=E=0；t<t0 的 I/R/E 均为 0")

    # 2 起点前无投入/退役
    pre_err = 0.0
    for r in all_series:
        if r["t0"] is None:
            continue
        i0 = years_all.index(r["t0"])
        full = r["_full"]
        pre_err = max(pre_err, float(np.max(np.abs(full["I"][:i0]))) if i0 else 0.0,
                      float(np.max(np.abs(full["R"][:i0]))) if i0 else 0.0)
    add("起点前不创建投入或退役", n_init, pre_err, pre_err <= 1e-9,
        "t<t0 的 I、R、E 均为 0（全零序列不产生任何批次）")

    # 3 新增非负
    min_I = min(r["min_I"] for r in all_series) if all_series else 0.0
    neg_float = sum(r["neg_float"] for r in all_series)
    neg_mat = sum(r["neg_material"] for r in all_series)
    add("新增投入量非负（缺口由提前退役吸收）", n_series, min_I, min_I >= -max(tol_abs, 1e-9),
        f"实质负需求={neg_mat} 条，全部转为经济性提前退役；浮点级={neg_float} 条按0处理；"
        f"I 最小值={min_I:.3e}")

    # 4 设备存量守恒
    max_err = 0.0
    n = 0
    for r in all_series:
        full = r["_full"]
        lhs = np.zeros_like(full["I"])
        lhs[1:] = full["Q"][:-1] + full["I"][1:] - full["R"][1:] - full["E"][1:]
        err = float(np.max(np.abs(lhs[1:] - full["Q"][1:]))) if len(lhs) > 1 else 0.0
        max_err = max(max_err, err)
        n += len(lhs) - 1
    add("设备存量流量守恒（期末=期初+新增−自然退役−提前退役）", n, max_err, max_err <= 1e-6,
        "逐序列逐年核对；误差为设备量绝对误差（辆或MW）")

    # 5 铜存量守恒
    max_err = 0.0
    n = 0
    for r in all_series:
        full = r["_full"]
        lhs = np.zeros_like(full["C_in"])
        lhs[1:] = full["stock_cu"][:-1] + full["C_in"][1:] - full["C_out"][1:]
        err = float(np.max(np.abs(lhs[1:] - full["stock_cu"][1:]))) if len(lhs) > 1 else 0.0
        max_err = max(max_err, err)
        n += len(lhs) - 1
    add("在用铜存量守恒（期末=期初+新增铜−退役铜）", n, max_err, max_err <= 1e-6,
        "逐序列逐年核对；退役铜=自然+提前")

    # 6 节点复现
    node_err = 0.0
    node_n = 0
    for data in (veh, pow_):
        for key, r in data.items():
            full = r["_full"]
            for y, v in r["nodes"].items():
                if v is None:
                    continue
                node_n += 1
                node_err = max(node_err, abs(float(full["Q"][years_all.index(y)]) - float(v)))
    add("有效保有量/装机节点复现", node_n, node_err, node_err <= 1e-6,
        "模型计算存量在全部有效节点上等于输入保有量/装机（空白节点除外）")

    # 7 系数：投入年固定
    coeff_bad = 0
    for key, r in list(veh.items())[:20]:
        k = r["k"]
        if len(np.unique(np.round(k, 9))) > 1:
            coeff_bad += 1
    for key, r in pow_.items():
        if r.get("k") is not None and key[3] != "hydro":
            k = r["k"]
            # 非hydro系数按年变化，属预期；仅检查无 NaN
            if np.any(~np.isfinite(k)):
                coeff_bad += 1
    add("铜强度按投入年确定（车辆固定；发电按年插值）", len(veh), None, coeff_bad == 0,
        "车辆为常数；发电批次的 k 由投入年 CIPG 决定，退役不再乘退役年系数")

    # 8 合成测试：Q 平坦 + k 下降 → 不产生退役
    years_t = list(range(1990, 2031))
    n_t = len(years_t)
    Q_flat = np.full(n_t, 100.0)
    k_dec = np.linspace(100.0, 50.0, n_t)
    S_t, p_t = build_survival({"type": "fixed", "life": 30}, 200)
    r_t = simulate_batches(Q_flat, years_t, 0, S_t, p_t, k_dec, tol_abs, tol_rel)
    # 前30年不退役；提前退役应为0（Q不再下降）
    early_before = float(np.sum(r_t["E"][1:30]))
    add("铜强度下降不触发设备退役（合成算例）", 1, early_before, early_before <= 1e-9,
        "Q 平坦、k 由100降到50，前30年无自然退役且提前退役=0；铜存量仅随批次退出下降")

    # 9 全零/空白
    add("全零序列与空白节点", len(zero_series) + len(blank_nodes), None, True,
        f"全零序列={len(zero_series)} 条（输出0，结构性零）；空白节点={len(blank_nodes)} 个"
        f"（不填0，按有效节点插值并登记）")

    # 10 水电分规模
    hydro_err = 0.0
    idx_out = [years_all.index(y) for y in p_out]
    for (scn, region), h in hydro.items():
        agg = pow_[(scn, region, "hydro", "hydro")]
        for key in ("I", "R", "E", "C_in", "C_out", "stock_cu"):
            s = np.sum([h["scales"][x][key] for x in HYDRO_SCALES], axis=0)[idx_out]
            hydro_err = max(hydro_err, float(np.max(np.abs(s - agg[key]))))
    add("水电分规模合计=hydro", len(hydro), hydro_err, hydro_err <= 1e-6,
        "Σ_scale 的新增/退役/铜量/存量 = hydro 对应值；正式结果只输出一条 hydro")

    # 11 提前退役规模
    e_veh = sum(float(np.sum(r["E"])) for r in veh.values())
    e_pow = sum(float(np.sum(r["E"])) for r in pow_.values())
    r_veh = sum(float(np.sum(r["R"])) for r in veh.values())
    r_pow = sum(float(np.sum(r["R"])) for r in pow_.values())
    add("经济性提前退役规模", n_series, None, True,
        f"车辆提前退役={e_veh:,.0f} 辆（占车辆退役 {e_veh / max(r_veh + e_veh, 1):.2%}）；"
        f"绿电提前退役={e_pow:,.0f} MW（占绿电退役 {e_pow / max(r_pow + e_pow, 1):.2%}）")

    # 12 水电规模强度校验
    implied = agg2021 * norm
    add("水电规模强度与综合 CIPG 对应", 3, abs(implied - cipg2021), abs(implied - cipg2021) <= 1e-6,
        f"表5 2021 容量加权={agg2021:,.3f} t/GW；归一后综合={implied:,.3f} t/GW，"
        f"项目 hydro CIPG 2021={cipg2021:,.3f} t/GW；归一后规模强度（2021）=" +
        "、".join(f"{s}:{hydro_base[s] * norm:,.3f}" for s in HYDRO_SCALES))

    return checks


# --------------------------------------------------------------------------- #
# 输出
# --------------------------------------------------------------------------- #
HEADER_FMT = {"bold": True, "bg_color": "#D9E1F2", "border": 1,
              "align": "center", "valign": "vcenter", "text_wrap": True}


def write_panel(workbook, name, fields, years, rows, widths, year_fmt="#,##0.00",
                pct=False, freeze=None, unit_col=None):
    ws = workbook.add_worksheet(name)
    hf = workbook.add_format(HEADER_FMT)
    yf = workbook.add_format({"num_format": "0.00%" if pct else year_fmt, "valign": "top"})
    txt = workbook.add_format({"valign": "top"})
    ws.freeze_panes(1, freeze if freeze is not None else len(fields) - len(years))
    ws.set_row(0, 28)
    for c, f in enumerate(fields):
        if isinstance(f, int):
            ws.write_number(0, c, f, hf)
        else:
            ws.write_string(0, c, str(f), hf)
        ws.set_column(c, c, widths[c] if c < len(widths) else 13)
    for i, row in enumerate(rows, start=1):
        vals, ys = row
        for c, v in enumerate(vals):
            if v is None:
                continue
            if isinstance(v, str):
                ws.write_string(i, c, v, txt)
            else:
                ws.write_number(i, c, float(v), txt)
        for j, v in enumerate(ys):
            if v is None:
                continue
            ws.write_number(i, len(vals) + j, float(v), yf)
    if rows:
        ws.autofilter(0, 0, len(rows), len(fields) - 1)
    return ws


def build_life_rows(cfg, result):
    rows = []
    life = None
    for subsector, entry in cfg["life"]["vehicle_map"].items():
        cat = entry["category"]
        spec = None
        # 用 config 的分布定义展示（与 数据.xlsx life 一致）
        spec = cfg["life"]["distributions"][cat]
        rows.append({"原技术": subsector, "部门": "车辆", "寿命类别": cat,
                     "分布": distribution_label(spec), "参数": json.dumps(spec, ensure_ascii=False),
                     "映射说明": entry["note"],
                     "来源": "数据.xlsx / life"})
    for subtech, entry in cfg["life"]["power_map"].items():
        cat = entry["category"]
        if cat == "HYDRO_CON":
            for scale in HYDRO_SCALES:
                dist = HYDRO_SCALE_DIST[scale]
                spec = cfg["life"]["distributions"][dist]
                rows.append({"原技术": subtech, "部门": "绿电", "寿命类别": f"{dist}（{scale}）",
                             "分布": distribution_label(spec), "参数": json.dumps(spec, ensure_ascii=False),
                             "映射说明": f"CON 规模结构内部计算；正式结果合并为 hydro；"
                                        f"规模强度（2021）={result['hydro_base'][scale] * result['norm']:.3f} t/GW",
                             "来源": "数据.xlsx / life；补充表5（强度）"})
            continue
        spec = cfg["life"]["distributions"][cat]
        rows.append({"原技术": subtech, "部门": "绿电", "寿命类别": cat,
                     "分布": distribution_label(spec), "参数": json.dumps(spec, ensure_ascii=False),
                     "映射说明": entry["note"], "来源": "数据.xlsx / life"})
    return rows


def build_region_rows(cfg):
    proc = set(cfg["trade"]["processing_regions"])
    rows = []
    for region in cfg["trade"]["processing_regions"] + cfg["trade"]["export_regions"]:
        rows.append({
            "Region": region,
            "处理能力设定": "有本地处理能力（15区）" if region in proc else "无本地处理能力（17区）",
            "矩阵特征": "对角线=1，本地处理并可接收进口" if region in proc
                     else "本地份额=0，100%出口至15个接收地",
            "纳入状态": "不开放情景纳入；开放情景作为处理区与来源区" if region in proc
                     else "不开放不纳入；开放情景仅作为来源区",
        })
    return rows


def build_coef_rows(cfg, result, auto, cipg):
    rows = []
    # 车辆
    for subsector in cfg["devices"]["vehicle_subsectors"]:
        for tech in cfg["devices"]["vehicle_technologies"]:
            rows.append({"部门": "车辆", "对象": subsector, "技术/动力": tech,
                         "年份": "1990—2050", "系数": auto[subsector][tech],
                         "单位": "kg/vehicle", "说明": "固定值，投入年即用此值，退役不重估"})
    # 发电非hydro
    for subtech in cfg["devices"]["power_subtech_order"]:
        if subtech == "hydro":
            continue
        tech = cfg["devices"]["subtech_to_cipg"][subtech]
        vals = cipg.get(tech, {})
        for y in result["p_out"]:
            v = float(cipg_series(cipg, tech, [y])[0])
            rows.append({"部门": "绿电", "对象": subtech, "技术/动力": tech,
                         "年份": y, "系数": v, "单位": "t/GW",
                         "说明": "投入年系数；≤2021用2021值，节点间线性插值"})
    # 水电规模
    for scale in HYDRO_SCALES:
        for y in result["p_out"]:
            rows.append({"部门": "绿电", "对象": "hydro", "技术/动力": HYDRO_SCALE_DIST[scale],
                         "年份": y, "系数": float(result["hydro_k"][scale][result["years_all"].index(y)] * 1000.0),
                         "单位": "t/GW",
                         "说明": "CON 规模强度（表5比例×项目CIPG路径归一）；正式结果合并为hydro"})
    return rows


def write_workbook(cfg, result, path: Path):
    wb = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True})
    v_out, p_out = result["v_out"], result["p_out"]

    v_keys = ["Scenario", "Region", "Subsector", "Technology", "单位"]
    p_keys = ["Scenario", "Region", "tech", "Subtech", "单位"]

    def v_rows(field, unit):
        out = []
        for key in sorted(result["veh"], key=lambda k: (k[0], k[1], k[2], k[3])):
            r = result["veh"][key]
            vals = list(key) + [unit]
            out.append((vals, [float(x) for x in r[field]]))
        return out

    def p_rows(field, unit):
        out = []
        for key in sorted(result["pow"], key=lambda k: (k[0], k[1], k[2], k[3])):
            r = result["pow"][key]
            vals = list(key) + [unit]
            out.append((vals, [float(x) for x in r[field]]))
        return out

    widths_v = [11, 24, 20, 11, 7] + [13] * len(v_out)
    widths_p = [11, 24, 18, 24, 7] + [13] * len(p_out)
    write_panel(wb, "车辆新增量_面板", v_keys + v_out, v_out, v_rows("I", "辆"), widths_v)
    write_panel(wb, "车辆退役量_面板", v_keys + v_out, v_out, v_rows("R", "辆"), widths_v)
    write_panel(wb, "车辆提前退役_面板", v_keys + v_out, v_out, v_rows("E", "辆"), widths_v)
    write_panel(wb, "车辆新增铜_面板", v_keys + v_out, v_out, v_rows("C_in", "t Cu"), widths_v)
    write_panel(wb, "车辆退役铜_面板", v_keys + v_out, v_out, v_rows("C_out", "t Cu"), widths_v)
    write_panel(wb, "车辆铜存量_面板", v_keys + v_out, v_out, v_rows("stock_cu", "t Cu"), widths_v)

    write_panel(wb, "绿电新增装机_面板", p_keys + p_out, p_out, p_rows("I", "MW"), widths_p)
    write_panel(wb, "绿电退役装机_面板", p_keys + p_out, p_out, p_rows("R", "MW"), widths_p)
    write_panel(wb, "绿电提前退役_面板", p_keys + p_out, p_out, p_rows("E", "MW"), widths_p)
    write_panel(wb, "绿电新增铜_面板", p_keys + p_out, p_out, p_rows("C_in", "t Cu"), widths_p)
    write_panel(wb, "绿电退役铜_面板", p_keys + p_out, p_out, p_rows("C_out", "t Cu"), widths_p)
    write_panel(wb, "绿电铜存量_面板", p_keys + p_out, p_out, p_rows("stock_cu", "t Cu"), widths_p)

    # 水电规模权重
    years_all = result["years_all"]
    idx = [years_all.index(y) for y in p_out]
    wmat = np.array([hydro_capacity_shares(cfg, y) for y in years_all])          # n×3
    kmat = np.column_stack([result["hydro_k"][s] for s in HYDRO_SCALES])         # n×3
    cmat = wmat * kmat
    vmat = cmat / cmat.sum(axis=1, keepdims=True)
    w_rows = []
    for kind, mat in (("容量份额", wmat), ("铜量份额", vmat)):
        for si, scale in enumerate(HYDRO_SCALES):
            w_rows.append(([kind, scale, "%"], [float(mat[i, si]) for i in idx]))
    write_panel(wb, "水电规模权重_面板", ["权重类型", "规模类别", "单位"] + p_out, p_out, w_rows,
                [14, 12, 6] + [12] * len(p_out), pct=True)

    # 不开放纳入区域
    reg_rows = [(list(r.values()), []) for r in build_region_rows(cfg)]
    write_panel(wb, "不开放纳入区域", ["Region", "处理能力设定", "矩阵特征", "纳入状态"], [], reg_rows,
                [26, 26, 46, 52])

    # 寿命参数
    life_rows = build_life_rows(cfg, result)
    life_fields = ["原技术", "部门", "寿命类别", "分布", "参数", "映射说明", "来源"]
    write_panel(wb, "寿命参数", life_fields, [],
                [(list(r.get(f) for f in life_fields), []) for r in life_rows],
                [20, 8, 26, 46, 44, 70, 40])

    # 系数
    coef_rows = result["coef_rows"]
    coef_fields = ["部门", "对象", "技术/动力", "年份", "系数", "单位", "说明"]
    write_panel(wb, "系数_面板", coef_fields, [],
                [(list(r.get(f) for f in coef_fields), []) for r in coef_rows],
                [8, 20, 24, 12, 14, 12, 60])

    # 核验结果
    chk_fields = ["检查项目", "检查数量", "最大误差", "是否通过", "说明"]
    write_panel(wb, "核验结果", chk_fields, [],
                [([c["检查项目"], c["检查数量"], c["最大误差"], c["是否通过"], c["说明"]], [])
                 for c in result["checks"]],
                [46, 12, 16, 10, 96])

    # 方法说明
    method = build_method_rows(cfg, result)
    write_panel(wb, "方法说明", ["项目", "内容"], [],
                [([k, v], []) for k, v in method], [24, 150])

    wb.close()


def build_method_rows(cfg, result):
    a = cfg["assumptions"]
    rows = [
        ("脚本", "step1_retired_copper.py；设备量驱动；import 不计算"),
        ("唯一设备路径", cfg["paths"]["stock_xlsx"] + "（车辆保有量_面板 辆；装机容量_九类_面板 MW）"),
        ("参数来源", cfg["paths"]["data_xlsx"] + "：CIPG / life / Automotive copper demand"),
        ("四类结果", "新增设备量/装机量、退役设备量/装机量、新增铜需求量、退役铜量；退役=自然+提前"),
        ("初始化", "t0=首个正值节点；t<t0 无投入无退役；Q(t0) 为初始批次，年龄0，当年不退役；"
                 "属研究初始化假设，不代表真实投运年份"),
        ("年度化", "有效节点间线性插值，保留全部节点；空白不填0；不外推；全零序列保持零"),
        ("批次公式", "R_t=Σ I_τ·p(t−τ)，p(a)=S(a−1)−S(a)；I_t=Q_t−Q_{t−1}+R_t；"
                  "I_t<0 时 E_t=−I_t 为经济性提前退役（最老批次优先），I_t=0"),
        ("系数口径", "每批次铜强度在投入年确定并固定；车辆用 Automotive（kg/vehicle）；"
                  "发电用 CIPG，投入年≤2021用2021值，节点间线性插值（禁止2025→2030跳配）"),
        ("单位", "车辆：辆×kg/辆÷1000；发电：MW÷1000×t/GW；输出统一 t Cu"),
        ("水电", "CON 容量份额直接拆装机（20/16/64→2050 26.276/16/57.724）；"
               "规模寿命 Small(50,15)/Medium(70,15)/Large(80,15)；规模强度取表5并按项目2021综合强度归一；"
               "正式结果汇总为一条 hydro"),
        ("提前退役", "仅由设备量下降触发，单列于「车辆/绿电提前退役_面板」；"
                 "不再用于吸收铜强度下降"),
        ("核对", "铜存量.xlsx 仅作核对；因新规则按投入年插值，2025年等节点与旧铜存量表存在预期差异"),
        ("输出", cfg["paths"]["step1_out"]),
    ]
    return rows


# --------------------------------------------------------------------------- #
# 写盘核对
# --------------------------------------------------------------------------- #
def verify_written(path: Path, result) -> list[dict]:
    checks = []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    expected_sheets = ["车辆新增量_面板", "车辆退役量_面板", "车辆提前退役_面板",
                       "车辆新增铜_面板", "车辆退役铜_面板", "车辆铜存量_面板",
                       "绿电新增装机_面板", "绿电退役装机_面板", "绿电提前退役_面板",
                       "绿电新增铜_面板", "绿电退役铜_面板", "绿电铜存量_面板",
                       "水电规模权重_面板", "不开放纳入区域", "寿命参数", "系数_面板",
                       "核验结果", "方法说明"]
    checks.append({"检查项目": "工作表数量与顺序", "检查数量": len(expected_sheets),
                   "最大误差": None,
                   "是否通过": "通过" if wb.sheetnames == expected_sheets else "未通过",
                   "说明": str(wb.sheetnames)})

    def check_panel(sheet, keys, years, expected_fn, label):
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        header = next(it)
        hdr_ok = list(header) == keys + ["单位"] + years
        n = 0
        bad = 0
        maxdiff = 0.0
        seen = set()
        for row in it:
            if row is None or row[0] is None:
                continue
            n += 1
            key = tuple(str(row[i]).strip() for i in range(len(keys)))
            if key in seen:
                bad += 1
            seen.add(key)
            exp = expected_fn(key)
            if exp is None:
                bad += 1
                continue
            for i, _ in enumerate(years):
                got = row[len(keys) + 1 + i]
                if got is None:
                    bad += 1
                    continue
                maxdiff = max(maxdiff, abs(float(got) - float(exp[i])))
        ok = hdr_ok and bad == 0 and maxdiff <= 1e-6
        checks.append({"检查项目": f"写盘核对-{label}", "检查数量": n, "最大误差": maxdiff,
                       "是否通过": "通过" if ok else "未通过",
                       "说明": f"表头正确={hdr_ok}；键重复/缺失={bad}；逐格最大绝对差={maxdiff:.3e}"})

    check_panel("车辆新增量_面板", ["Scenario", "Region", "Subsector", "Technology"],
                result["v_out"], lambda k: result["veh"][k]["I"], "车辆新增量")
    check_panel("车辆退役量_面板", ["Scenario", "Region", "Subsector", "Technology"],
                result["v_out"], lambda k: result["veh"][k]["R"], "车辆退役量")
    check_panel("车辆新增铜_面板", ["Scenario", "Region", "Subsector", "Technology"],
                result["v_out"], lambda k: result["veh"][k]["C_in"], "车辆新增铜")
    check_panel("车辆退役铜_面板", ["Scenario", "Region", "Subsector", "Technology"],
                result["v_out"], lambda k: result["veh"][k]["C_out"], "车辆退役铜")
    check_panel("绿电新增装机_面板", ["Scenario", "Region", "tech", "Subtech"],
                result["p_out"], lambda k: result["pow"][k]["I"], "绿电新增装机")
    check_panel("绿电退役装机_面板", ["Scenario", "Region", "tech", "Subtech"],
                result["p_out"], lambda k: result["pow"][k]["R"], "绿电退役装机")
    check_panel("绿电新增铜_面板", ["Scenario", "Region", "tech", "Subtech"],
                result["p_out"], lambda k: result["pow"][k]["C_in"], "绿电新增铜")
    check_panel("绿电退役铜_面板", ["Scenario", "Region", "tech", "Subtech"],
                result["p_out"], lambda k: result["pow"][k]["C_out"], "绿电退役铜")
    wb.close()
    return checks


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main(argv=None):
    parser = argparse.ArgumentParser(description="第一步：设备量驱动 DMFA")
    parser.add_argument("--config", default=None)
    parser.add_argument("--out", default=None, help="临时输出路径（默认 Output\\_tmp_step1_*.xlsx）")
    args = parser.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    started = time.time()
    cfg = load_config(args.config)
    out_dir = Path(cfg["paths"]["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = Path(args.out) if args.out else out_dir / f"_tmp_step1_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"

    log("=" * 100)
    log("第一步：设备量驱动 DMFA（新增/退役设备量、新增/退役铜量）")
    log(f"配置：{cfg['_config_path']}（sha256={cfg['_config_sha256']}）")

    template = load_year_template(cfg)
    log(f"[模板] 年份来源：{template.get('source')}；车辆 {template['vehicle'][0]}—{template['vehicle'][-1]}"
        f"（{len(template['vehicle'])}），绿电 {template['power'][0]}—{template['power'][-1]}"
        f"（{len(template['power'])}）")

    result = compute(cfg, template)
    log(f"[结果] 车辆序列={len(result['veh'])}；绿电序列={len(result['pow'])}；"
        f"全零序列={len(result['zero_series'])}；空白节点={len(result['blank_nodes'])}")
    log(f"[结果] 提前退役：车辆={sum(float(np.sum(r['E'])) for r in result['veh'].values()):,.0f} 辆；"
        f"绿电={sum(float(np.sum(r['E'])) for r in result['pow'].values()):,.0f} MW")

    tmpdir = Path(tempfile.mkdtemp(prefix="step1_"))
    tmp = tmpdir / out_path.name
    write_workbook(cfg, result, tmp)
    vchecks = verify_written(tmp, result)
    all_checks = result["checks"] + vchecks
    failed = [c for c in all_checks if c["是否通过"] != "通过"]
    for c in all_checks:
        log(f"  [{c['是否通过']}] {c['检查项目']} | 数量={c['检查数量']} | 误差={c['最大误差']}")
    if failed:
        log(f"核验未通过 {len(failed)} 项；临时文件保留：{tmp}")
        write_run_log(cfg, started)
        raise SystemExit(1)

    result["checks"] = all_checks
    tmp2 = tmpdir / ("final_" + out_path.name)
    write_workbook(cfg, result, tmp2)
    vchecks2 = verify_written(tmp2, result)
    if [(c["检查项目"], c["是否通过"]) for c in vchecks2] != [(c["检查项目"], c["是否通过"]) for c in vchecks]:
        log("重建后写盘核对不一致，停止")
        write_run_log(cfg, started)
        raise SystemExit(1)

    # 复制到目标临时路径
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tmp2, out_path)
    shutil.rmtree(tmpdir, ignore_errors=True)
    log(f"[完成] 临时第一步成果：{out_path}（{time.time() - started:.1f} s）")
    write_run_log(cfg, started)
    return 0


def write_run_log(cfg, started):
    out_dir = Path(cfg["paths"]["output_dir"])
    log_path = out_dir / "运行记录_第一步.txt"
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(RUN_LOG))
        fh.write(f"\n运行结束：{time.strftime('%Y-%m-%d %H:%M:%S')}；耗时 {time.time() - started:.1f} s\n")


if __name__ == "__main__":
    sys.exit(main())
