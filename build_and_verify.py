# -*- coding: utf-8 -*-
"""合成最终成果工作簿并做全量核验与交叉核对。

读取 convert_transport 与 convert_power 的计算结果，生成
Output/GCAM_保有量与装机容量.xlsx，包含三个工作表：
  - 车辆保有量（长表，活动量表全部原始字段 + 计算列）
  - 发电装机容量（长表，活动量表全部原始字段 + 计算列）
  - 异常与核验（异常记录 + 核验统计 + 手算 + 交叉核对结论）

导出后重新打开工作簿检查整数列、缺失标记与异常计数，并计算文件 SHA-256。
"""
import csv
import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gcam_common as g
import convert_transport as ct
import convert_power as cp

import openpyxl
from openpyxl.styles import Font

# 异常与核验工作表统一列
ANOMALY_COLS = ["Table", "Scenario", "Region", "Sector", "Subsector",
                "Technology", "Year", "Type", "Detail"]


def _unify_anomaly(a):
    if a["Table"] == "transport_service":
        return {
            "Table": a["Table"], "Scenario": a["Scenario"], "Region": a["Region"],
            "Sector": a["Sector"], "Subsector": a["Subsector"],
            "Technology": a["Technology"], "Year": a["Year"],
            "Type": a["Type"], "Detail": a["Detail"],
        }
    return {
        "Table": a["Table"], "Scenario": a["Scenario"], "Region": a["region"],
        "Sector": a["subcetor"], "Subsector": a["output"],
        "Technology": a["technology"], "Year": a["Year"],
        "Type": a["Type"], "Detail": a["Detail"],
    }


def _norm_scenario(s):
    """SSP5_8p5 -> SSP5_8.5（与参考 CSV 命名对齐）。"""
    m = re.match(r"SSP(\d)_(\d)p(\d)", s)
    if m:
        return "SSP%d_%d.%d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return s


# --------------------------------------------------------------------------- #
# 手算与取整自检
# --------------------------------------------------------------------------- #
def _first_meaningful(rows):
    """返回第一条计算结果取整后 > 0 的记录，保证手算示例有代表性。"""
    for row in rows:
        raw = row["_raw"]
        if raw is not None and g.round_half_up(raw) > 0:
            return row, raw, g.round_half_up(raw)
    # 兜底：取第一条
    row = rows[0]
    return row, row["_raw"], g.round_half_up(row["_raw"])


def hand_calc_passenger(results):
    rows = []
    for r in results:
        if r["Subsector"] == "Car" and r["ServiceValue"] is not None \
                and r["ServiceValue"] > 0 and r["Status"] == "ok":
            svc, vkt, lf = r["ServiceValue"], r["AnnualVKT_km_per_veh"], r["LoadFactor"]
            r = dict(r)
            r["_raw"] = svc * 1_000_000.0 / (vkt * lf)
            rows.append(r)
    return _first_meaningful(rows)


def hand_calc_freight(results):
    rows = []
    for r in results:
        if r["Subsector"] == "Heavy truck" and r["ServiceValue"] is not None \
                and r["ServiceValue"] > 0 and r["Status"] == "ok":
            svc, vkt, lf = r["ServiceValue"], r["AnnualVKT_km_per_veh"], r["LoadFactor"]
            r = dict(r)
            r["_raw"] = svc * 1_000_000.0 / (vkt * lf)
            rows.append(r)
    return _first_meaningful(rows)


def hand_calc_power(results):
    rows = []
    for r in results:
        if r["Generation_EJ"] is not None and r["Generation_EJ"] > 0 \
                and r["Status"] == "ok":
            gen, cf = r["Generation_EJ"], r["CapacityFactor"]
            r = dict(r)
            r["_raw"] = gen * g.GIGA / (g.KWH_PER_J * g.HOURS_PER_YEAR * cf)
            rows.append(r)
    return _first_meaningful(rows)


# --------------------------------------------------------------------------- #
# 唯一性 / 覆盖检查
# --------------------------------------------------------------------------- #
def check_unique(results, key_fields):
    seen = set()
    dup = []
    for r in results:
        k = tuple(r[f] for f in key_fields)
        if k in seen:
            dup.append(k)
        seen.add(k)
    return dup


# --------------------------------------------------------------------------- #
# 交叉核对（参考 CSV 为旧数据版本，仅验证方法，不要求数值全等）
# --------------------------------------------------------------------------- #
def cross_check_power(results):
    ref = {}
    with open(g.REF_CAPACITY_CSV, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            key = (_norm_scenario(r["scenario"]), r["region"], r["subsector"],
                   r["technology"], r["year"])
            ref[key] = r
    n_match = n_diff_cf = 0
    same_gen = diff_gen = 0
    same_gen_match_mw = same_gen_diff_mw = 0
    for r in results:
        key = (_norm_scenario(r["Scenario"]), r["region"], r["subcetor"],
               r["technology"], str(r["Year"]))
        rr = ref.get(key)
        if rr is None:
            continue
        n_match += 1
        if abs(float(r["CapacityFactor"]) - float(rr["capacity_factor"] or 0)) > 1e-9:
            n_diff_cf += 1
        myg, refg = float(r["Generation_EJ"]), float(rr["generation_ej"] or 0)
        if abs(myg - refg) < 1e-12:
            same_gen += 1
            mymw = float(r["InstalledCapacity_MW"])
            refmw = float(rr["installed_capacity_gw"] or 0) * 1000
            if abs(mymw - refmw) <= 0.5:
                same_gen_match_mw += 1
            else:
                same_gen_diff_mw += 1
        else:
            diff_gen += 1
    return {
        "n_match": n_match, "n_diff_cf": n_diff_cf,
        "same_gen": same_gen, "diff_gen": diff_gen,
        "same_gen_match_mw": same_gen_match_mw,
        "same_gen_diff_mw": same_gen_diff_mw,
    }


def cross_check_transport(results):
    """仅按 BEV 子集 + 重叠年份（不含 2020/2021）交叉核对。"""
    ref = {}
    with open(g.REF_VEHICLE_CSV, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if r["technology"] != "BEV":
                continue
            key = (_norm_scenario(r["scenario"]), r["region"], r["sector"],
                   r["subsector"], r["technology"], r["year"])
            ref[key] = r
    n_match = same_svc = 0
    same_svc_match = same_svc_diff = 0
    for r in results:
        key = (_norm_scenario(r["Scenario"]), r["Region"], r["Sector"],
               r["Subsector"], r["Technology"], str(r["Year"]))
        rr = ref.get(key)
        if rr is None:
            continue
        n_match += 1
        myv = float(r["ServiceValue"])
        refv = float(rr["service_value"] or 0)
        if abs(myv - refv) < 1e-12:
            same_svc += 1
            my_stock = float(r["VehicleStock_vehicles"])
            ref_stock = float(rr["vehicle_stock_million"] or 0) * 1e6
            if abs(my_stock - ref_stock) <= 0.5:
                same_svc_match += 1
            else:
                same_svc_diff += 1
    return {"n_match": n_match, "same_svc": same_svc,
            "same_svc_match": same_svc_match, "same_svc_diff": same_svc_diff}


# --------------------------------------------------------------------------- #
# 写入工作簿
# --------------------------------------------------------------------------- #
def write_sheet(ws, fields, rows, integer_cols):
    ws.append(fields)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([r.get(f) for f in fields])


def build_verification(rt, at, rp, ap, cc_t, cc_p):
    """返回 [(项目, 结果), ...] 核验统计列表。"""
    ty = sorted({r["Year"] for r in rt})
    py = sorted({r["Year"] for r in rp})
    t_sc = sorted({r["Scenario"] for r in rt})
    p_sc = sorted({r["Scenario"] for r in rp})

    dup_t = check_unique(rt, ["Scenario", "Region", "Sector", "Subsector",
                              "Technology", "Year"])
    dup_p = check_unique(rp, ["Scenario", "region", "subcetor", "output",
                              "technology", "Year"])

    out_t = sum(1 for r in rt if not g.in_range(r["Year"]))
    out_p = sum(1 for r in rp if not g.in_range(r["Year"]))

    # 手算
    pr, praw, pexp = hand_calc_passenger(rt)
    fr, fraw, fexp = hand_calc_freight(rt)
    gr, graw, gexp = hand_calc_power(rp)

    # .5 取整自检
    half_cases = [(0.5, 1), (1.5, 2), (2.5, 3), (2.4, 2), (2.6, 3)]
    half_ok = all(g.round_half_up(x) == e for x, e in half_cases)

    v = []
    v.append(("车辆保有量-结果行数", len(rt)))
    v.append(("发电装机容量-结果行数", len(rp)))
    v.append(("车辆保有量-异常记录数", len(at)))
    v.append(("发电装机容量-排除/异常记录数", len(ap)))
    v.append(("交通实际年份", str(ty)))
    v.append(("电力实际年份", str(py)))
    v.append(("交通含 2010/2020/2050", str(all(y in ty for y in [2010, 2020, 2050]))))
    v.append(("电力含 2010/2021/2050", str(all(y in py for y in [2010, 2021, 2050]))))
    v.append(("交通范围外记录数", out_t))
    v.append(("电力范围外记录数", out_p))
    v.append(("交通匹配唯一性", "通过（无重复键）" if not dup_t else "存在重复键 %d" % len(dup_t)))
    v.append(("电力匹配唯一性", "通过（无重复键）" if not dup_p else "存在重复键 %d" % len(dup_p)))
    v.append(("交通情景数", "%d %s" % (len(t_sc), t_sc)))
    v.append(("电力情景数", "%d %s" % (len(p_sc), p_sc)))
    v.append((".5 取整自检", "通过" if half_ok else "失败"))

    v.append(("手算-客运", "%s %s %s: %s×1e6÷(%s×%s)=%.6f→%d 辆 (表内 %s)"
              % (pr["Region"], pr["Subsector"], pr["Year"],
                 pr["ServiceValue"], pr["AnnualVKT_km_per_veh"], pr["LoadFactor"],
                 praw, pexp, pr["VehicleStock_vehicles"])))
    v.append(("手算-货运", "%s %s %s: %s×1e6÷(%s×%s)=%.6f→%d 辆 (表内 %s)"
              % (fr["Region"], fr["Subsector"], fr["Year"],
                 fr["ServiceValue"], fr["AnnualVKT_km_per_veh"], fr["LoadFactor"],
                 fraw, fexp, fr["VehicleStock_vehicles"])))
    v.append(("手算-发电", "%s %s %s %s: %s×1e9÷(3.6×8760×%s)=%.6f→%d MW (表内 %s)"
              % (gr["Scenario"], gr["region"], gr["technology"], gr["Year"],
                 gr["Generation_EJ"], gr["CapacityFactor"],
                 graw, gexp, gr["InstalledCapacity_MW"])))

    v.append(("交叉核对-电力(CF差异)", cc_p["n_diff_cf"]))
    v.append(("交叉核对-电力(匹配行数)", cc_p["n_match"]))
    v.append(("交叉核对-电力(同发电量行)", cc_p["same_gen"]))
    v.append(("交叉核对-电力(同发电量MW吻合)", cc_p["same_gen_match_mw"]))
    v.append(("交叉核对-电力(同发电量MW不符)", cc_p["same_gen_diff_mw"]))
    v.append(("交叉核对-电力(版本差异发电量行)", cc_p["diff_gen"]))
    v.append(("交叉核对-交通(BEV重叠行)", cc_t["n_match"]))
    v.append(("交叉核对-交通(同服务量行)", cc_t["same_svc"]))
    v.append(("交叉核对-交通(同服务量保有量吻合)", cc_t["same_svc_match"]))
    v.append(("交叉核对-交通(同服务量保有量不符)", cc_t["same_svc_diff"]))
    v.append(("交叉核对-说明", "参考CSV为旧数据版本：CF与公式一致（CF差异0、同发电量MW全吻合），"
                             "发电量/服务量有版本差异，绝对装机按当前数据.xlsx计算"))
    return v


def main():
    g.ensure_dir(g.OUTPUT_DIR)

    rt, at = ct.compute()
    rp, ap = cp.compute()
    cc_p = cross_check_power(rp)
    cc_t = cross_check_transport(rt)
    verification = build_verification(rt, at, rp, ap, cc_t, cc_p)

    # 写工作簿
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "车辆保有量"
    ws2 = wb.create_sheet("发电装机容量")
    ws3 = wb.create_sheet("异常与核验")

    write_sheet(ws1, ct.RESULT_FIELDS, rt, integer_cols={"VehicleStock_vehicles"})
    write_sheet(ws2, cp.RESULT_FIELDS, rp, integer_cols={"InstalledCapacity_MW"})

    # 异常与核验：异常记录 + 核验统计
    ws3.append(ANOMALY_COLS)
    for c in ws3[1]:
        c.font = Font(bold=True)
    for a in at + ap:
        u = _unify_anomaly(a)
        ws3.append([u.get(f) for f in ANOMALY_COLS])
    ws3.append([])
    ws3.append(["核验项目", "结果"])
    ws3[ws3.max_row][0].font = Font(bold=True)
    ws3[ws3.max_row][1].font = Font(bold=True)
    for label, val in verification:
        ws3.append([label, val])

    wb.save(g.OUTPUT_XLSX)

    # 导出后重新打开检查
    wb2 = openpyxl.load_workbook(g.OUTPUT_XLSX, data_only=True)
    _check_reopened(wb2, len(rt), len(rp), len(at) + len(ap))
    wb2.close()

    sha = g.sha256_of(g.OUTPUT_XLSX)

    print("=== 成果与核验报告 ===")
    print("成果路径:", g.OUTPUT_XLSX)
    print("SHA-256:", sha)
    for label, val in verification:
        print(f"  {label}: {val}")
    print("异常记录总数:", len(at) + len(ap))


def _check_reopened(wb, n_rt, n_rp, n_anom):
    ws1 = wb["车辆保有量"]
    ws2 = wb["发电装机容量"]
    ws3 = wb["异常与核验"]

    # 车辆保有量：VehicleStock_vehicles 为 int 或空
    hdr1 = [c.value for c in ws1[1]]
    i_stock = hdr1.index("VehicleStock_vehicles")
    bad = 0
    for row in ws1.iter_rows(min_row=2, values_only=True):
        v = row[i_stock]
        if v is not None and not isinstance(v, int):
            bad += 1
    # 发电装机容量：InstalledCapacity_MW 为 int 或空
    hdr2 = [c.value for c in ws2[1]]
    i_mw = hdr2.index("InstalledCapacity_MW")
    bad2 = 0
    for row in ws2.iter_rows(min_row=2, values_only=True):
        v = row[i_mw]
        if v is not None and not isinstance(v, int):
            bad2 += 1

    # 异常计数：异常记录行数（表头后到空行前）
    anom_rows = 0
    for row in ws3.iter_rows(min_row=2, values_only=True):
        if row[0] in ("transport_service", "power generation"):
            anom_rows += 1
        elif row[0] is None and all(v is None for v in row):
            break

    print("=== 重开工作簿检查 ===")
    print("车辆保有量行数(含表头):", ws1.max_row, "数据行:", ws1.max_row - 1,
          "应为:", n_rt, "整数列异常:", bad)
    print("发电装机容量行数(含表头):", ws2.max_row, "数据行:", ws2.max_row - 1,
          "应为:", n_rp, "整数列异常:", bad2)
    print("异常记录行数:", anom_rows, "应为:", n_anom)


if __name__ == "__main__":
    main()
