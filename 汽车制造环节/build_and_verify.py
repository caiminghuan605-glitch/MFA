# -*- coding: utf-8 -*-
"""把电力重分类结果导出到最终双面板工作簿，并做全量核验。

最终成果：C:\\Users\\22965\\OneDrive\\桌面\\MFA\\Output\\能源转型设备所有量.xlsx
  - 车辆保有量_面板：既有年份逐格保留，并新增 1990、2005 两列（源服务量为 0，
    BEV/PHEV 均记 0）
  - 装机容量_九类_面板：保留历史表名；实际列为 Scenario、region、tech、Subtech、
    单位、1990、2005、2010…2050，内容为八类 tech、十一类 Subtech（不混入 tech 合计行）

流程：
  1. convert_power.compute() 从原始技术重新计算未取整装机，再按
     Scenario+region+Year+tech+Subtech 汇总并四舍五入；
  2. 构造电力面板（剔除该 Scenario-region 下全部年份均无源记录的 Subtech 行，
     与旧最终面板“源表无记录则不出行”的口径一致）；
  3. 打开现有最终工作簿，扩展车辆面板年份，替换电力面板后原子写回；
  4. 重新打开文件，核对车辆既有年份逐格不变、1990/2005 为 0，检查面板主键
     唯一性与映射一致性，并对照更新前旧面板核对 2010–2050 既有值。

若 Output 下存在旧明细基线（_baseline_power_detail_pre_reclass.csv 等），
会额外做原始技术发电量/CF/未取整装机的逐键核对；不存在则跳过并在报告中注明。
"""
import csv
import math
import os
import sys
from collections import Counter, defaultdict

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g
import convert_power as cp

VEHICLE_SHEET = g.VEHICLE_PANEL_SHEET
POWER_SHEET = g.POWER_PANEL_SHEET
VEHICLE_KEY_FIELDS = ["Scenario", "Region", "Subsector", "Technology", "单位"]
VEHICLE_FIELDS = VEHICLE_KEY_FIELDS + list(g.VEHICLE_PANEL_YEARS)
POWER_FIELDS = ["Scenario", "region", "tech", "Subtech", "单位"] + list(g.POWER_PANEL_YEARS)
# 旧面板/旧明细已有年份（2010 起），用于核对本次扩展没有改变既有结果
BASELINE_POWER_YEARS = [y for y in g.POWER_PANEL_YEARS if y >= 2010]

BASELINE_DETAIL = os.path.join(g.OUTPUT_DIR, "_baseline_power_detail_pre_reclass.csv")
BASELINE_CATEGORIES = os.path.join(g.OUTPUT_DIR, "_baseline_power_categories_pre_reclass.csv")
BASELINE_FINAL = os.path.join(g.OUTPUT_DIR, "_baseline_能源转型设备所有量_pre_reclass.xlsx")
REPORT_TXT = os.path.join(g.OUTPUT_DIR, "_power_reclass_verification.txt")

# 旧归并类别 -> 新 Subtech 的 1:1 对照（用于核对未调整分类的结果不变）
ONE_TO_ONE = [
    ("gas_CCS", "gas_CC_CCS"),
    ("hydro", "hydro"),
    ("nuclear", "nuclear"),
    ("refined liquids_CCS", "refined liquids_CC_CCS"),
    ("PV", "PV"),
    ("wind", "wind"),
    ("wind_offshore", "wind_offshore"),
]
# 拆分分类：旧类别 = 多个新 Subtech 之和
SPLIT = [
    ("biomass_CCS", ("biomass_IGCC_CCS", "biomass_conv_CCS")),
    ("coal_CCS", ("coal_IGCC_CCS", "coal_conv_pul_CCS")),
]


# --------------------------------------------------------------------------- #
# 面板构造
# --------------------------------------------------------------------------- #
def build_power_panel(subtechs):
    """把长表 Subtech 结果透视为最终电力面板行。

    同一 Scenario+region 下，若某 Subtech 所有年份均为“源表无记录”（值全空），
    则该 Subtech 不出行；只要任一年有记录就保留整行，缺年保持空白。
    """
    by_key = defaultdict(dict)
    for r in subtechs:
        if r["tech"] is None or r["Subtech"] is None:
            continue
        key = (r["Scenario"], r["region"], r["tech"], r["Subtech"])
        by_key[key][r["Year"]] = r["InstalledCapacity_MW"]

    rows = []
    for key in sorted(by_key):
        vals = [by_key[key].get(y) for y in g.POWER_PANEL_YEARS]
        if all(v is None for v in vals):
            continue
        row = {
            "Scenario": key[0], "region": key[1],
            "tech": key[2], "Subtech": key[3], "单位": "MW",
        }
        for y in g.POWER_PANEL_YEARS:
            row[y] = by_key[key].get(y)
        rows.append(row)
    return rows


# --------------------------------------------------------------------------- #
# 工作簿读写
# --------------------------------------------------------------------------- #
def _sheet_values(ws):
    return [tuple(c.value for c in row) for row in ws.iter_rows()]


def _read_vehicle_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    if VEHICLE_SHEET not in wb.sheetnames:
        wb.close()
        raise SystemExit("最终工作簿缺少工作表：%s" % VEHICLE_SHEET)
    vals = _sheet_values(wb[VEHICLE_SHEET])
    wb.close()
    return vals


def _read_power_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if POWER_SHEET not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[POWER_SHEET]
    header = [c.value for c in ws[1]]
    rows = [tuple(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return header, rows


def _extend_vehicle_sheet(ws):
    """把车辆面板年份扩展到 g.VEHICLE_PANEL_YEARS。

    已有年份逐格保留；新增的 1990、2005 两年，交通源表服务量为 0，
    因此 BEV/PHEV 均记 0。返回是否发生了扩展。
    """
    header = [c.value for c in ws[1]]
    if header == VEHICLE_FIELDS:
        return False
    missing = [f for f in VEHICLE_KEY_FIELDS if f not in header]
    if missing:
        raise SystemExit("车辆面板缺少字段：%s" % missing)
    old_idx = {h: i for i, h in enumerate(header)}
    old_rows = [tuple(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    new_rows = []
    for r in old_rows:
        row = [r[old_idx[f]] for f in VEHICLE_KEY_FIELDS]
        for y in g.VEHICLE_PANEL_YEARS:
            row.append(r[old_idx[y]] if y in old_idx else 0)
        new_rows.append(row)
    ws.delete_rows(1, ws.max_row)
    ws.append(VEHICLE_FIELDS)
    for row in new_rows:
        ws.append(row)
    return True


def update_final_workbook(path, power_rows):
    """保留车辆面板并扩展年份，替换电力面板，原子写回同一文件。

    返回 (vehicle_before, vehicle_after, old_power_header, old_power_rows, power_after)。
    """
    if not os.path.exists(path):
        raise SystemExit(
            "找不到最终工作簿：%s\n"
            "该文件中的车辆保有量_面板是本次交付的既有内容，脚本只更新电力面板，"
            "不会重建车辆面板。" % path)

    wb = openpyxl.load_workbook(path, data_only=False)
    if VEHICLE_SHEET not in wb.sheetnames:
        wb.close()
        raise SystemExit("最终工作簿缺少工作表：%s" % VEHICLE_SHEET)

    vehicle_before = _sheet_values(wb[VEHICLE_SHEET])
    _extend_vehicle_sheet(wb[VEHICLE_SHEET])

    old_power_header = None
    old_power_rows = None
    if POWER_SHEET in wb.sheetnames:
        old_ws = wb[POWER_SHEET]
        old_power_header = [c.value for c in old_ws[1]]
        old_power_rows = [tuple(r) for r in old_ws.iter_rows(min_row=2, values_only=True)]
        del wb[POWER_SHEET]
    ws = wb.create_sheet(POWER_SHEET, 1)
    ws.append(POWER_FIELDS)
    for row in power_rows:
        ws.append([row.get(f) for f in POWER_FIELDS])

    tmp = path + ".tmp"
    wb.save(tmp)
    wb.close()
    os.replace(tmp, path)

    vehicle_after = _read_vehicle_sheet(path)
    power_after = _read_power_sheet(path)
    return (vehicle_before, vehicle_after, old_power_header, old_power_rows,
            power_after)


# --------------------------------------------------------------------------- #
# 核验
# --------------------------------------------------------------------------- #
def verify_vehicle_preserved(before, after):
    """核对车辆面板：既有年份逐单元格不变；新增 1990/2005 两列为 0。"""
    diffs = []
    bh = list(before[0])
    ah = list(after[0])
    if ah != VEHICLE_FIELDS:
        diffs.append(("表头", ah, VEHICLE_FIELDS))
        return diffs
    if len(before) - 1 != len(after) - 1:
        diffs.append(("行数", len(before) - 1, len(after) - 1))
        return diffs

    old_years = [y for y in bh if isinstance(y, int)]
    new_years = [y for y in ah if isinstance(y, int)]
    added = [y for y in new_years if y not in old_years]
    missing_new = [y for y in (1990, 2005) if y not in new_years]
    if missing_new:
        diffs.append(("缺少新增年份", missing_new, [1990, 2005]))
    if added and set(added) != {1990, 2005}:
        diffs.append(("新增年份", added, [1990, 2005]))

    bi = {h: i for i, h in enumerate(bh)}
    ai = {h: i for i, h in enumerate(ah)}
    for i, (rb, ra) in enumerate(zip(before[1:], after[1:])):
        for f in VEHICLE_KEY_FIELDS:
            if rb[bi[f]] != ra[ai[f]]:
                diffs.append(("行%d %s" % (i + 2, f), rb[bi[f]], ra[ai[f]]))
        for y in old_years:
            if rb[bi[y]] != ra[ai[y]]:
                diffs.append(("行%d 年%d" % (i + 2, y), rb[bi[y]], ra[ai[y]]))
        for y in (1990, 2005):
            if y in ai and ra[ai[y]] not in (0, None):
                diffs.append(("行%d 年%d 应为0" % (i + 2, y), ra[ai[y]], 0))
    return diffs


def verify_power_panel(header, rows):
    """校验电力面板表头、主键唯一性、tech/Subtech 映射与取值类型。"""
    problems = []
    if header != POWER_FIELDS:
        problems.append("表头不符：期望 %s，实际 %s" % (POWER_FIELDS, header))

    keys = [(r[0], r[1], r[2], r[3]) for r in rows]
    dup = [k for k, n in Counter(keys).items() if n > 1]
    if dup:
        problems.append("主键重复 %d 个，示例 %s" % (len(dup), dup[:3]))

    tech_set = {r[2] for r in rows}
    sub_set = {r[3] for r in rows}
    if tech_set != set(cp.TECH_ORDER):
        problems.append("tech 集合不符：%s" % sorted(tech_set))
    if sub_set != set(cp.SUBTECH_ORDER):
        problems.append("Subtech 集合不符：%s" % sorted(sub_set))
    for r in rows:
        if cp.SUBTECH_TO_TECH.get(r[3]) != r[2]:
            problems.append("Subtech→tech 映射不一致：%s -> %s" % (r[3], r[2]))
            break

    bad_type = 0
    for r in rows:
        for v in r[5:]:
            if v is not None and not isinstance(v, int):
                bad_type += 1
    if bad_type:
        problems.append("年份列存在非整数/非空值 %d 个" % bad_type)

    return problems


def _load_baseline_detail():
    if not os.path.exists(BASELINE_DETAIL):
        return None
    with open(BASELINE_DETAIL, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _detail_key(r, year_field="Year"):
    return (r["Scenario"], r["region"], r["subcetor"], r["technology"],
            str(r[year_field]))


def verify_detail_unchanged(detail, baseline):
    """原始技术的发电量、CF、未取整装机逐键与修改前基线比对。

    基线只有 2010–2050；本次新增的 1990/2005 单独计数，不参与旧基线比对。
    """
    if baseline is None:
        return {"skipped": True}
    base_years = {str(y) for y in BASELINE_POWER_YEARS}
    old = {_detail_key(r): r for r in baseline}
    new_all = {_detail_key(r): r for r in detail}
    new = {k: v for k, v in new_all.items() if k[4] in base_years}
    key_diff = set(old) ^ set(new)
    fields = ["Generation_EJ", "CapacityFactor", "InstalledCapacity_MW_unrounded"]
    mismatches = []
    for k in old.keys() & new.keys():
        for f in fields:
            a, b = old[k][f], new[k][f]
            if (a == "" or a is None) and (b is None or b == ""):
                continue
            if a == "" or b is None:
                mismatches.append((k, f, a, b))
                continue
            try:
                if not math.isclose(float(a), float(b), rel_tol=1e-12, abs_tol=1e-9):
                    mismatches.append((k, f, a, b))
            except (TypeError, ValueError):
                mismatches.append((k, f, a, b))
    return {
        "skipped": False,
        "old_rows": len(baseline),
        "new_rows": len(detail),
        "new_hist_rows": len(new_all) - len(new),
        "key_diff": len(key_diff),
        "mismatches": mismatches,
    }


def _sum_by(rows, cat_field, val_field):
    sums = defaultdict(float)
    seen = set()
    has_none = set()
    for r in rows:
        k = (r["Scenario"], r["region"], str(r["Year"]), r[cat_field])
        seen.add(k)
        v = r[val_field]
        if v is None or v == "":
            has_none.add(k)
        else:
            sums[k] += float(v)
    return sums, seen, has_none


def verify_unrounded_unchanged(detail, baseline):
    """未取整口径核对：1:1 类别不变；biomass/coal 拆分合计等于旧类别；风电 tech 合计对应旧两类之和。"""
    if baseline is None:
        return {"skipped": True}
    old_sums, old_seen, _ = _sum_by(baseline, "Category", "InstalledCapacity_MW_unrounded")
    new_sums, new_seen, _ = _sum_by(detail, "Subtech", "InstalledCapacity_MW_unrounded")

    one_to_one = []
    for old_cat, new_sub in ONE_TO_ONE:
        groups = [k for k in old_seen if k[3] == old_cat]
        mism = 0
        missing = 0
        for k in groups:
            nk = (k[0], k[1], k[2], new_sub)
            if nk not in new_seen:
                missing += 1
            elif not math.isclose(old_sums.get(k, 0.0), new_sums.get(nk, 0.0),
                                  rel_tol=1e-12, abs_tol=1e-9):
                mism += 1
        one_to_one.append((old_cat, new_sub, len(groups), mism, missing))

    split = []
    for old_cat, subs in SPLIT:
        groups = [k for k in old_seen if k[3] == old_cat]
        mism = 0
        missing = 0
        for k in groups:
            nv = 0.0
            found = False
            for s in subs:
                nk = (k[0], k[1], k[2], s)
                if nk in new_seen:
                    found = True
                    nv += new_sums.get(nk, 0.0)
            if not found:
                missing += 1
            elif not math.isclose(old_sums.get(k, 0.0), nv, rel_tol=1e-12, abs_tol=1e-9):
                mism += 1
        split.append((old_cat, "+".join(subs), len(groups), mism, missing))

    # 风电：旧 wind + wind_offshore 的未取整合计 == 新 tech=wind 两个 Subtech 的合计
    old_wind = defaultdict(float)
    for k, v in old_sums.items():
        if k[3] in ("wind", "wind_offshore"):
            old_wind[(k[0], k[1], k[2])] += v
    new_wind = defaultdict(float)
    for k, v in new_sums.items():
        if k[3] in ("wind", "wind_offshore"):
            new_wind[(k[0], k[1], k[2])] += v
    wind_mism = sum(1 for k, v in old_wind.items()
                    if not math.isclose(v, new_wind.get(k, 0.0), rel_tol=1e-12, abs_tol=1e-9))
    return {
        "skipped": False,
        "one_to_one": one_to_one,
        "split": split,
        "wind_groups": len(old_wind),
        "wind_mismatch": wind_mism,
    }


def _load_old_categories():
    if not os.path.exists(BASELINE_CATEGORIES):
        return None
    out = {}
    with open(BASELINE_CATEGORIES, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            out[(r["Scenario"], r["region"], str(r["Year"]), r["Category"])] = r["InstalledCapacity_MW"]
    return out


def _load_new_subtechs():
    path = cp.MID_SUBTECH_CSV
    if not os.path.exists(path):
        return None
    out = {}
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            out[(r["Scenario"], r["region"], str(r["Year"]), r["Subtech"])] = r["InstalledCapacity_MW"]
    return out


def _int_or_none(v):
    return None if v is None or v == "" else int(v)


def verify_rounded(old_cats, new_subs):
    """取整口径核对：1:1 类别整数完全相同；biomass/coal 分别取整允许 ±1。"""
    if old_cats is None or new_subs is None:
        return {"skipped": True}

    exact = []
    for old_cat, new_sub in ONE_TO_ONE:
        diff = 0
        examples = []
        for k, ov in old_cats.items():
            if k[3] != old_cat:
                continue
            nv = _int_or_none(new_subs.get((k[0], k[1], k[2], new_sub)))
            if nv is None:
                continue
            ov = _int_or_none(ov)
            if ov != nv:
                diff += 1
                if len(examples) < 3:
                    examples.append((k, ov, nv))
        exact.append((old_cat, new_sub, diff, examples))

    split = []
    for old_cat, subs in SPLIT:
        diff = 0
        max_abs = 0
        examples = []
        for k, ov in old_cats.items():
            if k[3] != old_cat:
                continue
            vals = [_int_or_none(new_subs.get((k[0], k[1], k[2], s))) for s in subs]
            if any(v is None for v in vals):
                continue
            nv = sum(vals)
            ov = _int_or_none(ov)
            if ov != nv:
                diff += 1
                max_abs = max(max_abs, abs(nv - ov))
                if len(examples) < 3:
                    examples.append((k, ov, nv))
        split.append((old_cat, "+".join(subs), diff, max_abs, examples))

    return {"skipped": False, "one_to_one": exact, "split": split}


def verify_panel_against_old(old_header, old_rows, power_rows):
    """最终面板对照更新前的同结构面板：2010–2050 既有年份逐格不变。"""
    if old_header is None or old_rows is None:
        return {"skipped": True}
    old_years = [y for y in old_header if isinstance(y, int)]
    old_vals = {}
    for r in old_rows:
        for j, y in enumerate(old_years):
            old_vals[(r[0], r[1], r[2], r[3], str(y))] = r[5 + j]

    new_vals = {}
    for r in power_rows:
        for y in g.POWER_PANEL_YEARS:
            new_vals[(r["Scenario"], r["region"], r["tech"],
                      r["Subtech"], str(y))] = r[y]

    diff = 0
    compared = 0
    examples = []
    for k, ov in old_vals.items():
        compared += 1
        nv = new_vals.get(k)
        if ov != nv:
            diff += 1
            if len(examples) < 5:
                examples.append((k, ov, nv))
    return {"skipped": False, "compared": compared,
            "diff": diff, "examples": examples}


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main():
    g.ensure_dir(g.OUTPUT_DIR)

    detail, subtechs, anomalies = cp.compute()
    power_rows = build_power_panel(subtechs)

    print("=== 电力重分类导出与核验 ===")
    print("换算明细行数:", len(detail))
    print("Subtech 聚合行数:", len(subtechs))
    print("最终电力面板行数:", len(power_rows))
    print("异常数:", len(anomalies))

    (vehicle_before, vehicle_after, old_power_header, old_power_rows,
     (header, panel_rows)) = update_final_workbook(g.FINAL_XLSX, power_rows)

    problems = []
    vdiff = verify_vehicle_preserved(vehicle_before, vehicle_after)
    if vdiff:
        problems.append("车辆面板既有年份或新增年份处理异常 %d 处" % len(vdiff))

    # 与修改前备份的车辆面板逐单元格比对（若备份仍存在）
    baseline_vehicle_diff = None
    if os.path.exists(BASELINE_FINAL):
        try:
            baseline_vehicle = _read_vehicle_sheet(BASELINE_FINAL)
            baseline_vehicle_diff = verify_vehicle_preserved(
                baseline_vehicle, vehicle_after)
            if baseline_vehicle_diff:
                problems.append("车辆面板与修改前备份不一致 %d 处"
                                % len(baseline_vehicle_diff))
        except Exception as e:  # noqa: BLE001
            print("基线车辆面板读取失败，跳过比对:", e)

    pp = verify_power_panel(header, panel_rows)
    problems.extend(pp)

    baseline_detail = _load_baseline_detail()
    dstat = verify_detail_unchanged(detail, baseline_detail)
    if not dstat["skipped"]:
        if dstat["key_diff"] or dstat["mismatches"]:
            problems.append("原始技术发电量/CF/未取整装机与基线不一致")

    ustat = verify_unrounded_unchanged(detail, baseline_detail)
    if not ustat["skipped"]:
        for old_cat, new_sub, groups, mism, missing in ustat["one_to_one"]:
            if mism or missing:
                problems.append("未取整 1:1 核对失败：%s -> %s" % (old_cat, new_sub))
        for old_cat, new_sub, groups, mism, missing in ustat["split"]:
            if mism or missing:
                problems.append("未取整拆分合计核对失败：%s -> %s" % (old_cat, new_sub))
        if ustat["wind_mismatch"]:
            problems.append("未取整风电 tech 合计核对失败")

    old_cats = _load_old_categories()
    new_subs = _load_new_subtechs()
    rstat = verify_rounded(old_cats, new_subs)

    pstat = verify_panel_against_old(old_power_header, old_power_rows, power_rows)
    if not pstat["skipped"] and pstat["diff"]:
        problems.append("2010–2050 电力面板既有值发生变化 %d 处" % pstat["diff"])

    # ---------------- 报告 ----------------
    hist_detail = [r for r in detail if r["Year"] < 2010]
    hist_ok = sum(1 for r in hist_detail if r["Status"] == "ok")
    hist_zero = sum(1 for r in hist_detail if r["Status"] == "zero_generation")
    hist_cf = sum(1 for r in hist_detail if r["CF_Source"] == cp.CF_HISTORY_SOURCE)

    lines = []
    add = lines.append
    add("电力重分类与年份扩展核验报告")
    add("=" * 60)
    add("最终工作簿: %s" % g.FINAL_XLSX)
    add("车辆面板: %d 行；电力面板: %d 行；年份范围 1990–2050"
        % (len(vehicle_after) - 1, len(panel_rows)))
    add("")

    add("[1] 车辆面板")
    add("  年份扩展: 新增 1990、2005 两列（交通源表服务量为 0，BEV/PHEV 均记 0）")
    add("  回写前后差异: %d（既有年份逐格 + 新增年份应为0）" % len(vdiff))
    if baseline_vehicle_diff is not None:
        add("  与修改前备份差异: %d" % len(baseline_vehicle_diff))
    add("")

    add("[2] 电力面板结构与映射")
    add("  表头: %s" % header)
    add("  主键 Scenario+region+tech+Subtech 重复数: %d"
        % len([k for k, n in Counter((r[0], r[1], r[2], r[3]) for r in panel_rows).items() if n > 1]))
    add("  tech 数: %d；Subtech 数: %d" % (len({r[2] for r in panel_rows}), len({r[3] for r in panel_rows})))
    add("  Subtech→tech 唯一映射: %s"
        % ("通过" if all(cp.SUBTECH_TO_TECH.get(r[3]) == r[2] for r in panel_rows) else "失败"))
    add("")

    add("[3] 原始技术计算口径（对照修改前明细）")
    add("  本次明细总行数: %d；其中 1990/2005 新增 %d 行"
        % (len(detail), len(hist_detail)))
    if dstat["skipped"]:
        add("  旧明细基线未找到，跳过逐键旧值核对；2010–2050 面板层见 [6]")
    else:
        add("  基线/本次明细行数: %d / %d" % (dstat["old_rows"], dstat["new_rows"]))
        add("  匹配键差异数: %d" % dstat["key_diff"])
        add("  发电量/CF/未取整装机逐键差异数: %d" % len(dstat["mismatches"]))
        for m in dstat["mismatches"][:3]:
            add("    示例: %s" % (m,))
    add("")

    add("[4] 未取整汇总核对")
    if ustat["skipped"]:
        add("  跳过：找不到基线明细")
    else:
        for old_cat, new_sub, groups, mism, missing in ustat["one_to_one"]:
            add("  %-22s -> %-22s 组数 %4d，差异 %d，缺记录 %d"
                % (old_cat, new_sub, groups, mism, missing))
        for old_cat, new_sub, groups, mism, missing in ustat["split"]:
            add("  %-22s -> %-22s 组数 %4d，差异 %d，缺记录 %d"
                % (old_cat, new_sub, groups, mism, missing))
        add("  风电旧两类未取整合计 vs 新 tech=wind: 组数 %d，差异 %d"
            % (ustat["wind_groups"], ustat["wind_mismatch"]))
    add("")

    add("[5] 取整结果核对（传统四舍五入，先汇总未取整再取整）")
    if rstat["skipped"]:
        add("  跳过：缺少基线分类或本次 Subtech CSV")
    else:
        for old_cat, new_sub, diff, examples in rstat["one_to_one"]:
            add("  %-22s -> %-22s 整数差异 %d%s"
                % (old_cat, new_sub, diff,
                   ("，示例 %s" % (examples[:2],)) if examples else ""))
        for old_cat, new_sub, diff, max_abs, examples in rstat["split"]:
            add("  %-22s -> %-22s 分别取整差异 %d（最大 |差| %d）%s"
                % (old_cat, new_sub, diff, max_abs,
                   ("，示例 %s" % (examples[:2],)) if examples else ""))
        add("  说明：拆分项由“旧类别合计后取整”改为“各 Subtech 分别取整”，")
        add("        差异仅可能为 ±1，属原取整规则的正常结果，未人为调整。")
    add("")

    add("[6] 最终面板对照更新前旧面板（2010–2050）")
    if pstat["skipped"]:
        add("  跳过：更新前工作簿没有旧电力面板")
    else:
        add("  逐格对比 %d 个既有年份值，差异 %d"
            % (pstat["compared"], pstat["diff"]))
        for ex in pstat["examples"]:
            add("    示例: %s" % (ex,))
    add("")

    add("[7] 1990/2005 新增年份")
    add("  电力明细: %d 行（ok %d，zero_generation %d，异常 %d）"
        % (len(hist_detail), hist_ok, hist_zero, len(anomalies)))
    add("  历史 CF 来源（CF表(历史长表)）: %d 条" % hist_cf)
    add("  车辆面板: 1990/2005 全部为 0（交通源表服务量为 0）")
    add("")

    add("[8] 结论")
    if problems:
        add("  未通过：")
        for p in problems:
            add("    - %s" % p)
    else:
        add("  通过：车辆既有年份逐格不变、1990/2005 为 0；电力既有年份逐格一致；")
        add("        年份扩展到 1990–2050；面板主键唯一；Subtech→tech 唯一映射。")

    text = "\n".join(lines)
    print(text)
    with open(REPORT_TXT, "w", encoding="utf-8") as f:
        f.write(text + "\n")

    if problems:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
