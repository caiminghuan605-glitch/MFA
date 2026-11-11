# -*- coding: utf-8 -*-
"""合成最终成果工作簿并做全量核验。

读取 convert_transport 与 convert_power 的计算结果，生成
Output/GCAM_保有量与九类装机容量_2010-2050.xlsx，包含四个工作表：
  - 车辆保有量（长表，活动量表全部原始字段 + 计算列）
  - 装机换算明细（原始技术、发电量、CF 及来源、未取整装机、CIPG、归并类别）
  - 装机容量_九类（情景、地区、年份、九类标签、整数 MW、计算状态）
  - 异常与核验（异常记录 + 核验统计 + 手算 + 一致性核对）

导出后重新打开文件，检查年份范围、因子匹配唯一性、原始发电记录覆盖、
九类映射与取整一致性，并计算文件 SHA-256。
"""
import csv
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g
import convert_transport as ct
import convert_power as cp

import openpyxl
from openpyxl.styles import Font

# 异常与核验工作表统一列
ANOMALY_COLS = ["Table", "Scenario", "Region", "Sector", "Technology",
                "Year", "Type", "Detail"]


def _unify_anomaly(a):
    if a["Table"] == "transport_service":
        return {
            "Table": a["Table"], "Scenario": a["Scenario"], "Region": a["Region"],
            "Sector": "%s / %s" % (a["Sector"], a["Subsector"]),
            "Technology": a["Technology"], "Year": a["Year"],
            "Type": a["Type"], "Detail": a["Detail"],
        }
    return {
        "Table": a["Table"], "Scenario": a["Scenario"], "Region": a["region"],
        "Sector": a["subcetor"], "Technology": a["technology"], "Year": a["Year"],
        "Type": a["Type"], "Detail": a["Detail"],
    }


def _check_unique(rows, key_fields):
    seen = set()
    dup = []
    for r in rows:
        k = tuple(r[f] for f in key_fields)
        if k in seen:
            dup.append(k)
        seen.add(k)
    return dup


def _hand_calc_pv(detail, categories):
    """选一组由不同容量因子构成的合并项（PV：rooftop_pv/PV/PV_storage）手算加总。"""
    pv = [r for r in detail if r["Category"] == "PV"]
    combo = defaultdict(set)
    for r in pv:
        combo[(r["Scenario"], r["region"], r["Year"])].add(r["technology"])
    key = next(k for k, v in combo.items() if len(v) >= 3)
    subs = [r for r in pv
            if (r["Scenario"], r["region"], r["Year"]) == key]
    lines = []
    total = 0.0
    for r in subs:
        mw = float(r["InstalledCapacity_MW_unrounded"]) if r["InstalledCapacity_MW_unrounded"] is not None else 0.0
        total += mw
        lines.append("%s(gen=%s,CF=%s)=%s" % (r["technology"], r["Generation_EJ"],
                                              r["CapacityFactor"], r["InstalledCapacity_MW_unrounded"]))
    cat_row = next(c for c in categories
                   if c["Scenario"] == key[0] and c["region"] == key[1]
                   and c["Year"] == key[2] and c["Category"] == "PV")
    expected = g.round_half_up(total)
    actual = cat_row["InstalledCapacity_MW"]
    return key, lines, total, expected, actual


def _build_verification(rt, at, pd, pc, ap):
    ty = sorted({r["Year"] for r in rt})
    py = sorted({r["Year"] for r in pd})
    t_sc = sorted({r["Scenario"] for r in rt})
    p_sc = sorted({r["Scenario"] for r in pd})

    dup_t = _check_unique(rt, ["Scenario", "Region", "Sector", "Subsector",
                               "Technology", "Year"])
    dup_detail = _check_unique(pd, ["Scenario", "region", "subcetor",
                                    "technology", "Year"])
    dup_cat = _check_unique(pc, ["Scenario", "region", "Year", "Category"])

    out_t = sum(1 for r in rt if not g.in_range(r["Year"]))
    out_p = sum(1 for r in pd if not g.in_range(r["Year"]))

    unmapped = [r["technology"] for r in pd if r["Category"] is None]
    n_supplement = sum(1 for r in pd if r["CF_Source"] and r["CF_Source"].startswith("匹配明细"))

    # 手算 PV 多技术合并项
    hc_key, hc_lines, hc_total, hc_expected, hc_actual = _hand_calc_pv(pd, pc)

    # 取整前一致性：原始技术合计 vs 九类合计
    cat_idx = {(c["Scenario"], c["region"], c["Year"], c["Category"]): c for c in pc}
    detail_sum = defaultdict(float)
    for r in pd:
        if r["InstalledCapacity_MW_unrounded"] is not None:
            detail_sum[(r["Scenario"], r["region"], r["Year"], r["Category"])] += float(r["InstalledCapacity_MW_unrounded"])
    n_consist = 0
    n_inconsist = 0
    for key, val in detail_sum.items():
        c = cat_idx.get(key)
        if c is None or c["InstalledCapacity_MW"] is None:
            continue
        if int(c["InstalledCapacity_MW"]) == g.round_half_up(val):
            n_consist += 1
        else:
            n_inconsist += 1

    half_cases = [(0.5, 1), (1.5, 2), (2.5, 3), (2.4, 2), (2.6, 3)]
    half_ok = all(g.round_half_up(x) == e for x, e in half_cases)

    v = []
    v.append(("车辆保有量-结果行数", len(rt)))
    v.append(("装机换算明细-行数", len(pd)))
    v.append(("装机容量_九类-行数", len(pc)))
    v.append(("车辆保有量-异常数", len(at)))
    v.append(("电力-异常数", len(ap)))
    v.append(("交通实际年份", str(ty)))
    v.append(("电力实际年份", str(py)))
    v.append(("交通含 2010/2020/2050", str(all(y in ty for y in [2010, 2020, 2050]))))
    v.append(("电力含 2010/2021/2050", str(all(y in py for y in [2010, 2021, 2050]))))
    v.append(("交通范围外记录数", out_t))
    v.append(("电力范围外记录数", out_p))
    v.append(("交通匹配唯一性", "通过" if not dup_t else "重复 %d" % len(dup_t)))
    v.append(("装机明细匹配唯一性", "通过" if not dup_detail else "重复 %d" % len(dup_detail)))
    v.append(("九类匹配唯一性", "通过" if not dup_cat else "重复 %d" % len(dup_cat)))
    v.append(("交通情景数", "%d %s" % (len(t_sc), t_sc)))
    v.append(("电力情景数", "%d %s" % (len(p_sc), p_sc)))
    v.append(("九类映射未覆盖技术数", len(set(unmapped))))
    v.append(("补充因子使用记录数", n_supplement))
    v.append((".5 取整自检", "通过" if half_ok else "失败"))
    v.append(("手算-九类合并项(PV)", "%s %s %s: %s = 未取整合计 %s → %d MW (表内 %s)"
              % (hc_key[0], hc_key[1], hc_key[2], " + ".join(hc_lines),
                 hc_total, hc_expected, hc_actual)))
    v.append(("取整前一致核对", "一致 %d / 不一致 %d" % (n_consist, n_inconsist)))
    v.append(("因子来源说明", "主表 CF（region+subcetor+technology+year）；"
                             "六种 CCS 从 匹配明细 按 Scenario+region+subcetor+technology+year 补齐"))
    return v


def _write_sheet(ws, fields, rows):
    ws.append(fields)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([r.get(f) for f in fields])


def _reopen_check(path, n_rt, n_pd, n_pc, n_anom):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws1 = wb["车辆保有量"]
    ws2 = wb["装机换算明细"]
    ws3 = wb["装机容量_九类"]
    ws4 = wb["异常与核验"]

    h1 = [c.value for c in ws1[1]]
    i = h1.index("VehicleStock_vehicles")
    bad1 = sum(1 for row in ws1.iter_rows(min_row=2, values_only=True)
               if row[i] is not None and not isinstance(row[i], int))

    h2 = [c.value for c in ws2[1]]
    i2 = h2.index("InstalledCapacity_MW_unrounded")
    bad2 = sum(1 for row in ws2.iter_rows(min_row=2, values_only=True)
               if row[i2] is not None and not isinstance(row[i2], (int, float)))

    h3 = [c.value for c in ws3[1]]
    i3 = h3.index("InstalledCapacity_MW")
    bad3 = sum(1 for row in ws3.iter_rows(min_row=2, values_only=True)
               if row[i3] is not None and not isinstance(row[i3], int))

    anom_rows = 0
    for row in ws4.iter_rows(min_row=2, values_only=True):
        if row[0] in ("transport_service", "power generation"):
            anom_rows += 1
        elif row[0] is None and all(v is None for v in row):
            break

    print("=== 重开工作簿检查 ===")
    print("车辆保有量: %d 行(应 %d), 整数列异常 %d" % (ws1.max_row - 1, n_rt, bad1))
    print("装机换算明细: %d 行(应 %d), 未取整列类型异常 %d" % (ws2.max_row - 1, n_pd, bad2))
    print("装机容量_九类: %d 行(应 %d), 整数列异常 %d" % (ws3.max_row - 1, n_pc, bad3))
    print("异常记录行数: %d (应 %d)" % (anom_rows, n_anom))
    return bad1 + bad2 + bad3


def main():
    g.ensure_dir(g.OUTPUT_DIR)

    rt, at = ct.compute()
    pd, pc, ap = cp.compute()
    verification = _build_verification(rt, at, pd, pc, ap)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "车辆保有量"
    ws2 = wb.create_sheet("装机换算明细")
    ws3 = wb.create_sheet("装机容量_九类")
    ws4 = wb.create_sheet("异常与核验")

    _write_sheet(ws1, ct.RESULT_FIELDS, rt)
    _write_sheet(ws2, cp.RESULT_FIELDS, pd)
    _write_sheet(ws3, cp.CATEGORY_FIELDS, pc)

    ws4.append(ANOMALY_COLS)
    for c in ws4[1]:
        c.font = Font(bold=True)
    for a in at + ap:
        u = _unify_anomaly(a)
        ws4.append([u.get(f) for f in ANOMALY_COLS])
    ws4.append([])
    ws4.append(["核验项目", "结果"])
    ws4[ws4.max_row][0].font = Font(bold=True)
    ws4[ws4.max_row][1].font = Font(bold=True)
    for label, val in verification:
        ws4.append([label, val])

    wb.save(g.OUTPUT_XLSX)

    bad = _reopen_check(g.OUTPUT_XLSX, len(rt), len(pd), len(pc), len(at) + len(ap))
    sha = g.sha256_of(g.OUTPUT_XLSX)

    print("=== 成果与核验报告 ===")
    print("成果路径:", g.OUTPUT_XLSX)
    print("SHA-256:", sha)
    for label, val in verification:
        print("  %s: %s" % (label, val))
    print("异常记录总数:", len(at) + len(ap))
    print("重开检查异常总数:", bad)


if __name__ == "__main__":
    main()
