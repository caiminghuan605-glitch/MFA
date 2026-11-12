# -*- coding: utf-8 -*-
"""GCAM 能源情景转换的共享工具与全局常量。

供 convert_transport.py / convert_power.py / build_and_verify.py 复用：
- 只读安全读取（被 Excel 占用时降级为共享只读临时副本）
- 年份列识别与归一（整数年份 / 文本年份 / VKT_YYYY 前缀统一）
- 传统四舍五入（.5 向上）
- 全局路径与换算口径常量
"""
import math
import os
import shutil
import tempfile

import openpyxl

# --------------------------------------------------------------------------- #
# 路径常量
# --------------------------------------------------------------------------- #
DATA_XLSX = r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx"
CF_XLSX = (
    r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力"
    r"\电力部门_容量因子.xlsx"
)
REF_VEHICLE_CSV = (
    r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\新能源电车"
    r"\新能源汽车保有量.csv"
)
REF_CAPACITY_CSV = (
    r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力"
    r"\发电装机容量.csv"
)
OUTPUT_DIR = r"C:\Users\22965\OneDrive\桌面\MFA\Output"
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "GCAM_保有量与九类装机容量_2010-2050.xlsx")

# 最终交付的双面板工作簿（车辆面板原样保留，电力面板按 tech/Subtech 重分类）
FINAL_XLSX = os.path.join(OUTPUT_DIR, "能源转型设备所有量.xlsx")
VEHICLE_PANEL_SHEET = "车辆保有量_面板"
POWER_PANEL_SHEET = "装机容量_九类_面板"
# 电力面板沿用原年份集合与顺序（保留 2021，不改为车辆的 2020）
POWER_PANEL_YEARS = [2010, 2015, 2021, 2025, 2030, 2035, 2040, 2045, 2050]

# CCS 因子补充工作簿（六种 CCS 技术缺失因子时按 Scenario+region+subcetor+technology+year 补齐）
CCS_MATCH_XLSX = (
    r"C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力"
    r"\GCAM_发电量_容量因子匹配_2010-2050.xlsx"
)

# --------------------------------------------------------------------------- #
# 全局口径常量
# --------------------------------------------------------------------------- #
YEAR_MIN = 2010          # 起始年（含）
YEAR_MAX = 2050          # 结束年（含）
BASE_YEAR = 2021         # GCAM 基准年
HOURS_PER_YEAR = 8760    # 年度换算小时数（EIA 容量因子定义口径）
KWH_PER_J = 3.6          # 1 kWh = 3.6e6 J；EJ->MW 公式中的 3.6
GIGA = 1_000_000_000.0   # 1e9，EJ -> MW 换算系数


# --------------------------------------------------------------------------- #
# 数值工具
# --------------------------------------------------------------------------- #
def round_half_up(x):
    """传统四舍五入：小数部分 .5 向上取整，返回 int；None -> None。

    实现：floor(x + 0.5)。示例：0.5->1, 1.5->2, 2.5->3, 2.4->2, 2.6->3。
    """
    if x is None:
        return None
    return int(math.floor(x + 0.5))


def to_number(v):
    """把单元格值转为 float；空 / None / 非数值 -> None。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def is_zero(v):
    """判断数值是否明确为零（用于区分『数值零』与『空白』）。"""
    return v is not None and v == 0.0


# --------------------------------------------------------------------------- #
# 年份列识别
# --------------------------------------------------------------------------- #
def parse_year_token(v):
    """把表头单元格解析为年份 int；无法识别则返回 None。

    支持三种来源：
    - 整数年份：1990 / 2010 / 2050 ...
    - 文本年份：'2005' / '2010' ...
    - VKT 前缀列：'VKT_2005' / 'VKT_2020' ...
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        # 浮点年份必须是整数（如 2010.0）
        if isinstance(v, float) and v != int(v):
            return None
        return int(v)
    if isinstance(v, str):
        s = v.strip()
        if s.upper().startswith("VKT_"):
            s = s[4:]
        if s.isdigit():
            return int(s)
    return None


def split_year_columns(header):
    """把表头拆成「元数据列」与「年份列」。

    返回 (meta, years)：
    - meta: [(列索引, 列名), ...]（非年份列，保留原始表头名）
    - years: [(列索引, 年份int), ...]（年份列，已归一为 int）
    """
    meta = []
    years = []
    for i, h in enumerate(header):
        y = parse_year_token(h)
        if y is not None:
            years.append((i, y))
        else:
            meta.append((i, h))
    return meta, years


def in_range(year):
    """判断年份是否落在 [YEAR_MIN, YEAR_MAX] 范围内。"""
    return YEAR_MIN <= year <= YEAR_MAX


# --------------------------------------------------------------------------- #
# 只读安全读取
# --------------------------------------------------------------------------- #
def _copy_shared_read_only(src, dst):
    """以共享只读方式复制文件（Windows 下用 ctypes 指定共享模式）。

    当 Excel 以允许读共享的方式占用文件时，普通 open 可能失败；
    此函数用 FILE_SHARE_READ|FILE_SHARE_WRITE 打开源文件再复制。
    若源文件被独占（不共享读），仍会抛 PermissionError。
    """
    if os.name != "nt":
        shutil.copy2(src, dst)
        return
    import ctypes
    from ctypes import wintypes

    GENERIC_READ = 0x80000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    CreateFileW = kernel32.CreateFileW
    CreateFileW.restype = wintypes.HANDLE
    CreateFileW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
    ]
    ReadFile = kernel32.ReadFile
    ReadFile.argtypes = [
        wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p,
    ]
    CloseHandle = kernel32.CloseHandle

    handle = CreateFileW(
        src, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, None,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None,
    )
    if handle == wintypes.HANDLE(-1).value or handle is None:
        raise PermissionError("无法以共享只读方式打开文件：%s" % src)
    try:
        with open(dst, "wb") as out:
            buf = ctypes.create_string_buffer(1024 * 1024)
            while True:
                nread = wintypes.DWORD(0)
                ok = ReadFile(handle, buf, len(buf), ctypes.byref(nread), None)
                if not ok or nread.value == 0:
                    break
                out.write(buf.raw[: nread.value])
    finally:
        CloseHandle(handle)


def read_workbook_safe(path):
    """只读打开工作簿；被占用时降级为共享只读临时副本再打开。"""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        tmp = os.path.join(
            tempfile.gettempdir(),
            os.path.basename(path) + ".shared_read_only.xlsx",
        )
        _copy_shared_read_only(path, tmp)
        return openpyxl.load_workbook(tmp, read_only=True, data_only=True)


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def sha256_of(path):
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()
