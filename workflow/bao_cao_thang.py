# -*- coding: utf-8 -*-
"""
TỔNG HỢP BÁO CÁO DOANH SỐ NHƯỢNG QUYỀN THEO THÁNG
=================================================
- Nhận file báo cáo tuần (.xls / .xlsx) hoặc file .zip chứa các file báo cáo tuần.
- Tự động chuyển mã font tiếng Việt cũ (TCVN3, VNI, mojibake...) sang Unicode NFC
  và tái sử dụng logic đọc/kiểm tra từ `kiem_tra_nhap_hang.py` & `viet_restore.py`.
- Đọc file báo cáo tháng gốc (mặc định `BC_TONG NQ_Tuan45 9-15.xlsx` / `templates/BC_TONG_base.xlsx`):
  + Nếu tuần được chọn là tuần đầu tiên của tháng (hoặc khác tháng/năm với file gốc):
    làm mới dữ liệu các tuần trên Dashboard để bắt đầu tháng mới.
  + Nếu cùng tháng: giữ nguyên dữ liệu các tuần trước và cập nhật/cộng dồn vào đúng tuần được chọn.
- Quản lý danh mục `Ds_nhuong_quyen` với Checkbox Form Control thật của Excel (liên kết TRUE/FALSE),
  tự điền tên/khu vực còn thiếu và tự thêm các Mã NQ mới phát sinh ("Ngoài danh mục").
- Xuất báo cáo với 16 sheet chuẩn hóa, công thức động, biểu đồ tối ưu,
  `DuLieuTho(tuan_nay)` và `LichSuBaoCao(tong_hop)`.
"""
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import os
import re
import sys
import tempfile
import unicodedata
import zipfile
from collections import OrderedDict, defaultdict
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

import openpyxl
import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawingFont,
    Paragraph,
    ParagraphProperties,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    ROOT_DIR = os.path.dirname(os.path.abspath(sys.executable))
    WORKFLOW_DIR = os.path.join(ROOT_DIR, "workflow")
else:
    WORKFLOW_DIR = THIS_DIR
    ROOT_DIR = os.path.dirname(THIS_DIR)

if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

import kiem_tra_nhap_hang as knh  # noqa: E402
from viet_restore import Restorer, detect_encoding, to_unicode  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# --------------------------------------------------------------------------------------
# CẤU HÌNH HỆ THỐNG & MÀU SẮC
# --------------------------------------------------------------------------------------
DEFAULT_OUT_DIR = os.path.join(ROOT_DIR, "output")
DEFAULT_TEMPLATE = os.path.join(ROOT_DIR, "templates", "BC_TONG_base.xlsx")
NGOAI_DM = "Ngoài danh mục"

# Định nghĩa quản lý & sheet tương ứng (giữ đúng cấu trúc 5 quản lý + 2 KV)
SUPPORTERS = [
    {"key": "Duyên", "kv": "KV Tỉnh", "dash": "5. KV Duyên", "detail": "6. Duyên", "default_target": 1_300_000_000},
    {"key": "Quỳnh", "kv": "KV Tỉnh", "dash": "7.KV Quỳnh", "detail": "8. Quỳnh", "default_target": 630_000_000},
    {"key": "Cường", "kv": "KV HCM", "dash": "9.KVCường", "detail": "10.CƯỜNG", "default_target": 1_133_000_000},
    {"key": "Thế", "kv": "KV HCM", "dash": "11.KVThế", "detail": "12.THẾ", "default_target": 1_133_000_000},
    {"key": "Tuyết", "kv": "KV HCM", "dash": "13.KVTuyết", "detail": "14.TUYẾT", "default_target": 1_133_000_000},
]
SUPPORTER_BY_KEY = {s["key"]: s for s in SUPPORTERS}

# Fallback khu vực cho vài mã chưa có trong ListNQ
FALLBACK_KV = {
    "F1-021": "Q7",
    "F2-283": "Q6",
    "F3-122": "Q9",
    "F5-187": "Q11",
    "FBDN-14": "Đà Nẵng",
}
FALLBACK_NAMES = {
    "FBDN-14": "NQ Trần Nhân Tông - Đà Nẵng",
}

# Bảng màu doanh nghiệp thống nhất
FONT_NAME = "Calibri"
FILL_NAVY = PatternFill("solid", fgColor="1F4E78", bgColor="1F4E78")
FILL_SUBHDR = PatternFill("solid", fgColor="2F5597", bgColor="2F5597")
FILL_ICE = PatternFill("solid", fgColor="D9E1F2", bgColor="D9E1F2")
FILL_ZEBRA = PatternFill("solid", fgColor="F7F9FC", bgColor="F7F9FC")
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF", bgColor="FFFFFF")
FILL_GREEN = PatternFill("solid", fgColor="E2EFDA", bgColor="E2EFDA")
FILL_AMBER = PatternFill("solid", fgColor="FFF2CC", bgColor="FFF2CC")
FILL_GREY = PatternFill("solid", fgColor="EDEDED", bgColor="EDEDED")
FILL_RED = PatternFill("solid", fgColor="FCE4D6", bgColor="FCE4D6")

FONT_TITLE = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
FONT_SECTION = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
FONT_HDR = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
FONT_SUBHDR = Font(name=FONT_NAME, size=10, bold=True, color="1F4E78")
FONT_DATA = Font(name=FONT_NAME, size=10, bold=False, color="1E293B")
FONT_BOLD = Font(name=FONT_NAME, size=10, bold=True, color="1E293B")
FONT_INPUT = Font(name=FONT_NAME, size=10, bold=False, color="0000FF")
FONT_GREEN = Font(name=FONT_NAME, size=10, bold=True, color="1E6B34")
FONT_MUTED = Font(name=FONT_NAME, size=9, italic=True, color="64748B")

THIN_SIDE = Side(style="thin", color="D1D5DB")
MED_NAVY = Side(style="medium", color="1F4E78")
DOUBLE_SIDE = Side(style="double", color="1F4E78")

BORDER_CELL = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_TOTAL = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY, bottom=DOUBLE_SIDE)

FMT_INT = "#,##0"
FMT_DEC = "#,##0.0"
FMT_PCT = "0.0%"
FMT_HIDDEN = ";;;"

CHART_COLORS = ["1F4E78", "2E75B6", "ED7D31", "70AD47", "FFC000", "5B9BD5"]

# Mã sản phẩm theo chuẩn kiểm tra nhập hàng
PROD_CODES = {
    "hamburger": "1001",
    "hotdog": "1004",
    "pita": "1003",
    "sandwich": "1002",   # số miếng; phần = / 2
    "bo_th": "1214",
    "mieng_ga": "1200",
    "mieng_tom": "1203",
    "mieng_thit": "1201",
}

HISTORY_HEADERS = [
    "Năm", "Tháng", "Tuần số", "Nhãn tuần", "Từ ngày", "Đến ngày",
    "KV Tổng", "Quản lý", "Khu vực / Tỉnh", "Mã NQ", "NHƯỢNG QUYỀN",
    "Còn hoạt động", "Ngày KD", "Tổng bánh", "Bình quân/ngày",
    "Hamburger", "Hotdog", "Pita", "Sandwich (phần)",
    "Bò thượng hạng", "Miếng gà", "Miếng tôm", "Miếng thịt",
    "Doanh thu (VNĐ)", "Thu hồi công nợ (VNĐ)",
]


# --------------------------------------------------------------------------------------
# TIỆN ÍCH CHUẨN HÓA & TUẦN LỊCH
# --------------------------------------------------------------------------------------
def norm_code(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\s*-\s*", "-", s)
    return s.upper()


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s))
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn").replace("đ", "d").replace("Đ", "D")


NGUNG_HD = "Ngưng hoạt động"


def norm_supporter(val) -> str:
    if not val:
        return NGOAI_DM
    s = strip_accents(str(val).strip()).lower()
    if s == "ngung hoat dong":
        return NGUNG_HD
    for sup in SUPPORTERS:
        if strip_accents(sup["key"]).lower() == s:
            return sup["key"]
    return str(val).strip()


def get_kv_tong(supporter_key: str) -> str:
    """Trả về KV Tổng theo tên quản lý; giữ nguyên 'Ngưng hoạt động' nếu quản lý là 'Ngưng hoạt động'."""
    if strip_accents(str(supporter_key or "").strip()).lower() == "ngung hoat dong":
        return NGUNG_HD
    return SUPPORTER_BY_KEY.get(supporter_key, {}).get("kv", NGOAI_DM)


def _week_num_vn(d: dt.date) -> int:
    """Số tuần theo chuẩn file báo cáo (Monday-based %W: 09-15/11/2026 = Tuần 45)."""
    w = int(d.strftime("%W"))
    return w if w >= 1 else 1


def get_month_week_slots(year: int, month: int) -> List[dict]:
    """Chia tháng thành các tuần (Thứ 2 -> Chủ nhật, hoặc ngày lẻ đầu/cuối tháng)."""
    last_day = calendar.monthrange(year, month)[1]
    d = dt.date(year, month, 1)
    end_month = dt.date(year, month, last_day)
    slots = []
    idx = 1
    while d <= end_month:
        # Kết thúc slot vào Chủ nhật gần nhất (weekday()==6) hoặc cuối tháng
        days_to_sun = 6 - d.weekday()
        slot_end = min(end_month, d + dt.timedelta(days=days_to_sun))
        w_num = _week_num_vn(d)
        days_cnt = (slot_end - d).days + 1
        if d == slot_end:
            label = f"Ngày {d:%d/%m/%Y}"
        else:
            label = f"Tuần {w_num:02d} {d:%d} - {slot_end:%d/%m/%Y}"
        slots.append({
            "slot_idx": idx,
            "week_num": w_num,
            "start_date": d,
            "end_date": slot_end,
            "days_count": days_cnt,
            "label": label,
            "short_key": f"{d:%d/%m}" if d == slot_end else f"{d:%d}-{slot_end:%d}/{month:02d}",
        })
        d = slot_end + dt.timedelta(days=1)
        idx += 1
    return slots


def parse_date_range_from_filename(path: str, default_year: Optional[int] = None) -> Optional[dict]:
    """
    Nhận diện khoảng ngày / tháng / năm từ tên file báo cáo tuần.
    Ví dụ:
      - '16 - 22.11.xls' -> 16/11..22/11/(default_year)
      - '09-15.11.2026.xlsx' -> 09/11..15/11/2026
      - '29.10 - 04.11.xls' -> 29/10..04/11 (chọn tháng chứa nhiều ngày hơn / ngày kết thúc)
      - '01.11.2026.xls' -> 01/11/2026
    """
    name = os.path.splitext(os.path.basename(path))[0]
    if default_year is None:
        default_year = dt.date.today().year

    # 1) Dạng dd.mm(.yyyy) - dd.mm(.yyyy)
    m = re.search(
        r"(\d{1,2})\s*[./-]\s*(\d{1,2})(?:\s*[./-]\s*(\d{4}))?\s*[-–]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})(?:\s*[./-]\s*(\d{4}))?",
        name,
    )
    if m:
        d1, m1, y1, d2, m2, y2 = m.groups()
        yr = int(y2 or y1 or default_year)
        try:
            dt1 = dt.date(int(y1 or yr), int(m1), int(d1))
            dt2 = dt.date(yr, int(m2), int(d2))
            return _match_slot_from_dates(dt1, dt2)
        except ValueError:
            pass

    # 2) Dạng dd - dd.mm(.yyyy) hoặc dd-dd/mm/yyyy
    m = re.search(
        r"(\d{1,2})\s*[-–]\s*(\d{1,2})\s*[./]\s*(\d{1,2})(?:\s*[./]\s*(\d{4}))?",
        name,
    )
    if m:
        d1, d2, mm, yy = m.groups()
        yr = int(yy or default_year)
        try:
            dt1 = dt.date(yr, int(mm), int(d1))
            dt2 = dt.date(yr, int(mm), int(d2))
            return _match_slot_from_dates(dt1, dt2)
        except ValueError:
            pass

    # 3) Dạng 1 ngày: dd.mm.yyyy hoặc dd.mm
    m = re.search(r"(\d{1,2})\s*[./]\s*(\d{1,2})(?:\s*[./]\s*(\d{4}))?", name)
    if m:
        d1, mm, yy = m.groups()
        yr = int(yy or default_year)
        try:
            dt1 = dt.date(yr, int(mm), int(d1))
            return _match_slot_from_dates(dt1, dt1)
        except ValueError:
            pass

    return None


def _match_slot_from_dates(dt1: dt.date, dt2: dt.date) -> dict:
    # Nếu trải qua 2 tháng, chọn tháng có nhiều ngày hơn trong khoảng [dt1, dt2]
    if dt1.month == dt2.month and dt1.year == dt2.year:
        target = dt2
    else:
        end_m1 = dt.date(dt1.year, dt1.month, calendar.monthrange(dt1.year, dt1.month)[1])
        days_m1 = max(0, (end_m1 - dt1).days + 1)
        start_m2 = dt.date(dt2.year, dt2.month, 1)
        days_m2 = max(0, (dt2 - start_m2).days + 1)
        target = dt1 if days_m1 > days_m2 else dt2

    slots = get_month_week_slots(target.year, target.month)
    best_slot = slots[0]
    best_overlap = -1
    for sl in slots:
        ov_start = max(dt1, sl["start_date"])
        ov_end = min(dt2, sl["end_date"])
        overlap = (ov_end - ov_start).days + 1 if ov_end >= ov_start else 0
        if overlap > best_overlap:
            best_overlap = overlap
            best_slot = sl

    ov_start = max(dt1, best_slot["start_date"])
    ov_end = min(dt2, best_slot["end_date"])
    days_in_slot = (ov_end - ov_start).days + 1 if ov_end >= ov_start else best_slot["days_count"]
    return {
        "year": target.year,
        "month": target.month,
        "slot_idx": best_slot["slot_idx"],
        "week_num": best_slot["week_num"],
        "start_date": ov_start if ov_end >= ov_start else best_slot["start_date"],
        "end_date": ov_end if ov_end >= ov_start else best_slot["end_date"],
        "days_count": max(1, days_in_slot),
        "slot": best_slot,
    }


def build_output_filename(slot: dict, year: int, month: int) -> str:
    d1, d2 = slot["start_date"], slot["end_date"]
    if d1 == d2:
        return f"BC_TONG NQ_Ngay{d1:%d}.{month:02d}.{year}.xlsx"
    return f"BC_TONG NQ_Tuan{slot['week_num']:02d} {d1:%d}-{d2:%d}.{month:02d}.{year}.xlsx"


def _match_slot_from_label(label: str, slots: List[dict]) -> Optional[int]:
    """Khớp nhãn tuần cũ (vd 'Tuần 01/11/2026', 'Tuần 44 02 - 08/11/2026', 'Ngày 30/11/2026') vào slot_idx."""
    if not label:
        return None
    s = str(label).strip()
    m = re.search(r"(\d{1,2})\s*[-–]\s*(\d{1,2})\s*/\s*(\d{1,2})", s)
    if m:
        d1, d2, mm = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for sl in slots:
            if sl["start_date"].month == mm and sl["start_date"].day == d1 and sl["end_date"].day == d2:
                return sl["slot_idx"]
    m2 = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", s)
    if m2:
        d1, mm = int(m2.group(1)), int(m2.group(2))
        for sl in slots:
            if sl["start_date"].month == mm and sl["start_date"].day == d1 and sl["end_date"].day == d1:
                return sl["slot_idx"]
    return None


# --------------------------------------------------------------------------------------
# ĐỌC FILE BÁO CÁO TUẦN (XLS / XLSX / ZIP) + CHUYỂN FONT UNICODE
# --------------------------------------------------------------------------------------
def _read_single_excel(path: str, harvest_texts: List[str]) -> pd.DataFrame:
    df, _ = knh.load_weekly_report(path)
    df["ma_kh"] = df["ma_kh"].map(norm_code)
    df["tien"] = pd.to_numeric(df["tien"], errors="coerce").fillna(0.0)
    df["ten_kh_goc"] = df["ten_kh"]
    df["ten_hh_goc"] = df["ten_hh"]

    cols = ("ten_kh", "ten_hh", "dvt")
    sample = []
    for c in cols:
        sample.extend(list(df[c].dropna().unique())[:5000])
    enc = detect_encoding(sample)
    if enc != "unicode":
        print(f"  [{os.path.basename(path)}] Phát hiện bảng mã cũ: {enc.upper()} -> chuyển sang Unicode NFC")
    for c in cols:
        df[c] = df[c].map(lambda v: to_unicode(v, enc))
    if enc in ("unicode", "vnsans"):
        rs = Restorer(harvest_texts)
        is_vns = enc == "vnsans"
        for c in cols:
            df[c] = df[c].map(lambda v: rs.restore(v, vnsans=is_vns))
    return df


def load_weekly_inputs(
    input_path: str,
    harvest_texts: List[str],
    target_year: Optional[int] = None,
    target_month: Optional[int] = None,
) -> Tuple[pd.DataFrame, dict]:
    """
    Đọc 1 file Excel (.xls/.xlsx) hoặc file .zip chứa các file Excel.
    Nếu truyền target_year & target_month và dữ liệu có cột ngày ('ngay'),
    tự động lọc bỏ các dòng không thuộc tháng được chọn (vd tuần đầu/cuối tháng vắt sang tháng khác).
    Trả về (DataFrame tổng hợp, dict pivot & thống kê theo mã NQ).
    """
    ext = os.path.splitext(input_path)[1].lower()
    frames = []
    if ext == ".zip":
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(input_path, "r") as zf:
                members = [
                    n for n in zf.namelist()
                    if not n.endswith("/")
                    and "__MACOSX" not in n
                    and not os.path.basename(n).startswith("~$")
                    and os.path.splitext(n)[1].lower() in (".xls", ".xlsx")
                ]
                if not members:
                    raise ValueError(f"Không tìm thấy file .xls / .xlsx nào trong file zip: {input_path}")
                for member in sorted(members):
                    extracted = zf.extract(member, tmpdir)
                    frames.append(_read_single_excel(extracted, harvest_texts))
    elif ext in (".xls", ".xlsx"):
        frames.append(_read_single_excel(input_path, harvest_texts))
    else:
        raise ValueError(f"Định dạng file không được hỗ trợ: {ext} (chỉ nhận .xls, .xlsx, .zip)")

    df = pd.concat(frames, ignore_index=True)

    # Lọc bỏ dữ liệu của các ngày không nằm trong tháng/năm được chọn (trường hợp tuần đầu/cuối vắt 2 tháng)
    if target_year is not None and target_month is not None and df["ngay"].notna().any():
        before_cnt = len(df)
        mask = (df["ngay"].dt.year == target_year) & (df["ngay"].dt.month == target_month)
        df = df[mask].copy()
        dropped_cnt = before_cnt - len(df)
        if dropped_cnt > 0:
            print(f"  Đã lọc bỏ {dropped_cnt} dòng ngoài tháng {target_month:02d}/{target_year}.")

    # Tổng hợp theo mã NQ
    pivot_qty: Dict[str, Dict[str, float]] = defaultdict(dict)
    for (ma, hh), qty in df.groupby(["ma_kh", "ma_hh"])["sl"].sum().items():
        pivot_qty[ma][hh] = float(qty)

    rev_by_store = {ma: float(v) for ma, v in df.groupby("ma_kh")["tien"].sum().items()}

    # Số ngày đặt hàng (nếu có cột ngày thực tế từ định dạng 11 cột)
    days_by_store: Dict[str, Optional[int]] = {}
    if df["ngay"].notna().any():
        for ma, grp in df.groupby("ma_kh"):
            n_days = grp["ngay"].dropna().dt.date.nunique()
            if n_days > 0:
                days_by_store[ma] = int(n_days)

    name_by_store = {}
    for ma, grp in df.groupby("ma_kh"):
        name_by_store[ma] = knh.best_name(grp["ten_kh"])

    # Danh mục hàng hóa có trong file tuần (mã -> tên Unicode)
    item_names = {}
    for hh, grp in df.groupby("ma_hh"):
        item_names[hh] = knh.best_name(grp["ten_hh"])

    return df, {
        "pivot": pivot_qty,
        "revenue": rev_by_store,
        "days": days_by_store,
        "names": name_by_store,
        "item_names": item_names,
    }


# --------------------------------------------------------------------------------------
# ĐỌC TRẠNG THÁI BÁO CÁO THÁNG GỐC
# --------------------------------------------------------------------------------------
def _to_float(v) -> float:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return 0.0


GROUP_FILLS = [
    PatternFill("solid", fgColor="FEF3C7"),  # Vàng nhạt
    PatternFill("solid", fgColor="DBEAFE"),  # Xanh dương nhạt
    PatternFill("solid", fgColor="D1FAE5"),  # Xanh ngọc nhạt
    PatternFill("solid", fgColor="EDE9FE"),  # Tím nhạt
    PatternFill("solid", fgColor="FFEDD5"),  # Cam nhạt
    PatternFill("solid", fgColor="FCE7F3"),  # Hồng nhạt
]


def _get_group_fill(group_id: str, group_color_map: dict) -> Optional[PatternFill]:
    if not group_id:
        return None
    if group_id not in group_color_map:
        group_color_map[group_id] = GROUP_FILLS[len(group_color_map) % len(GROUP_FILLS)]
    return group_color_map[group_id]


def _find_gom_nhom_col(ws) -> Optional[int]:
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h is not None and strip_accents(str(h).strip()).lower() == "gom nhom":
            return c
    return None


def _norm_group_id(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def _reorder_grouped_stores(stores: OrderedDict) -> OrderedDict:
    """Sắp xếp lại để các cửa hàng có cùng mã 'Gom nhóm' luôn đứng liền kề nhau."""
    by_group = defaultdict(list)
    for ma, st in stores.items():
        gid = st.get("group", "")
        if gid:
            by_group[gid].append((ma, st))
    out = OrderedDict()
    emitted_groups = set()
    for ma, st in stores.items():
        gid = st.get("group", "")
        if not gid:
            out[ma] = st
        elif gid not in emitted_groups:
            emitted_groups.add(gid)
            for g_ma, g_st in by_group[gid]:
                out[g_ma] = g_st
    return out


def _split_total_across_group(total: float, n: int) -> List[float]:
    """Chia đều `total` cho `n` phần tử; nếu có số dư thì dồn hết phần dư vào phần tử đầu tiên."""
    if n <= 1 or total == 0:
        return [total] + [0.0] * max(0, n - 1)
    if abs(total - round(total)) < 1e-6:
        t_int = int(round(total))
        q, rem = divmod(t_int, n)
        return [float(q + rem)] + [float(q)] * (n - 1)
    q = round(total / n, 2)
    first = round(total - q * (n - 1), 2)
    return [first] + [q] * (n - 1)


def _apply_group_value_split(stores: OrderedDict, weekly_info: dict) -> None:
    """Gom tổng doanh thu & số lượng hàng theo 'Gom nhóm' rồi chia đều (dư dồn về mã đầu nhóm)."""
    by_group = defaultdict(list)
    for ma, st in stores.items():
        gid = st.get("group", "")
        if gid:
            by_group[gid].append(ma)

    pivot = weekly_info["pivot"]
    revenue = weekly_info["revenue"]
    for gid, members in by_group.items():
        n = len(members)
        if n <= 1:
            continue
        tot_rev = sum(revenue.get(m, 0.0) for m in members)
        all_codes = sorted({c for m in members for c in pivot.get(m, {})})
        if tot_rev == 0 and not all_codes:
            continue
        rev_parts = _split_total_across_group(tot_rev, n)
        for m, part in zip(members, rev_parts):
            if part:
                revenue[m] = part
            else:
                revenue.pop(m, None)
        for code in all_codes:
            tot_qty = sum(pivot.get(m, {}).get(code, 0.0) for m in members)
            qty_parts = _split_total_across_group(tot_qty, n)
            for m, part in zip(members, qty_parts):
                if part:
                    pivot.setdefault(m, defaultdict(float))[code] = part
                elif m in pivot and code in pivot[m]:
                    del pivot[m][code]


def _parse_bool_active(val) -> Optional[bool]:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        s = strip_accents(val.strip()).lower()
        if s in ("true", "1", "co", "yes", "x", "dang hoat dong"):
            return True
        if s in ("false", "0", "khong", "no", "off", "ngung", "ngung hoat dong"):
            return False
    return None


def _read_legacy_store_info(wb_val: openpyxl.Workbook) -> Tuple[dict, dict, dict, dict]:
    """
    Trích xuất thông tin bổ sung (Tên, Khu vực/Tỉnh, Ngày KD, Số liệu tuần trước) từ các sheet trong file gốc.
    Trả về: (names_map, kv_map, days_map, week_store_data)
    """
    names_map: Dict[str, str] = {}
    kv_map: Dict[str, str] = {}
    days_map: Dict[str, int] = {}
    week_store_data: Dict[str, dict] = {}

    # 1) Từ ListNQ (nếu còn trong workbook gốc)
    if "ListNQ" in wb_val.sheetnames:
        ws = wb_val["ListNQ"]
        for r in range(2, ws.max_row + 1):
            ma = norm_code(ws.cell(r, 1).value)
            if not ma:
                continue
            kv = ws.cell(r, 2).value
            ten = ws.cell(r, 4).value
            if isinstance(kv, str) and kv.strip() and ma not in kv_map:
                kv_map[ma] = kv.strip()
            if isinstance(ten, str) and ten.strip() and ma not in names_map:
                names_map[ma] = ten.strip()

    # 2) Từ các sheet chi tiết cũ (6. Duyên, 8. Quỳnh, 10.CƯỜNG, 12.THẾ, 14.TUYẾT)
    detail_configs = [
        ("6. Duyên", 3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 15, 16, 17),
        ("8. Quỳnh", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14),
        ("10.CƯỜNG", 1, 2, 3, 4, 5, 6, 7, 8, None, 9, 10, 11, 12),
        ("12.THẾ", 1, 2, 3, 4, 5, 6, 7, 8, None, 9, 10, 11, 12),
        ("14.TUYẾT", 1, 2, 3, 4, 5, 6, 7, 8, None, 9, 10, 11, 12),
    ]
    for sname, c_ma, c_ten, c_day, c_avg, c_hb, c_hd, c_pt, c_sw, c_tot, c_bo, c_ga, c_tom, c_thit in detail_configs:
        if sname not in wb_val.sheetnames:
            continue
        ws = wb_val[sname]
        # Kiểm tra xem có phải format mới của tool này không (row 3 có header 'Mã NQ' ở col 2)
        is_new_fmt = (str(ws.cell(3, 2).value or "").strip() == "Mã NQ")
        if is_new_fmt:
            for r in range(5, ws.max_row + 1):
                ma = norm_code(ws.cell(r, 2).value)
                if not ma or ma.startswith("KHU VỰC") or ma.startswith("TỔNG"):
                    break
                ten = ws.cell(r, 3).value
                kv = ws.cell(r, 4).value
                if isinstance(ten, str) and ten.strip() and ma not in names_map:
                    names_map[ma] = ten.strip()
                if isinstance(kv, str) and kv.strip() and ma not in kv_map:
                    kv_map[ma] = kv.strip()
                dy = _to_float(ws.cell(r, 6).value)
                if dy > 0:
                    days_map[ma] = int(round(dy))
                week_store_data[ma] = {
                    "days": int(round(dy)),
                    "avg": _to_float(ws.cell(r, 7).value),
                    "hamburger": _to_float(ws.cell(r, 8).value),
                    "hotdog": _to_float(ws.cell(r, 9).value),
                    "pita": _to_float(ws.cell(r, 10).value),
                    "sandwich": _to_float(ws.cell(r, 11).value),
                    "total": _to_float(ws.cell(r, 12).value),
                    "bo_th": _to_float(ws.cell(r, 13).value),
                    "mieng_ga": _to_float(ws.cell(r, 14).value),
                    "mieng_tom": _to_float(ws.cell(r, 15).value),
                    "mieng_thit": _to_float(ws.cell(r, 16).value),
                    "revenue": _to_float(ws.cell(r, 17).value),
                    "debt": _to_float(ws.cell(r, 18).value),
                }
            continue

        # Format cũ
        in_top = True
        in_rev = False
        rev_c_ma, rev_c_rev, rev_c_kv = (3, 5, 7) if sname == "6. Duyên" else ((2, 4, 6) if sname == "8. Quỳnh" else (1, 3, None))
        for r in range(2, ws.max_row + 1):
            v_ma = ws.cell(r, c_ma).value
            if isinstance(v_ma, str) and v_ma.strip() == "KHU VỰC":
                in_top = False
            if isinstance(ws.cell(r, rev_c_ma).value, str) and ws.cell(r, rev_c_ma).value.strip() == "Mã NQ" and r > 20:
                in_rev = True
                continue
            if in_top and isinstance(v_ma, str) and ("-" in v_ma):
                ma = norm_code(v_ma)
                ten = ws.cell(r, c_ten).value
                if isinstance(ten, str) and ten.strip() and ma not in names_map:
                    names_map[ma] = ten.strip()
                dy = _to_float(ws.cell(r, c_day).value)
                if dy > 0 and ma not in days_map:
                    days_map[ma] = int(round(dy))
                hb = _to_float(ws.cell(r, c_hb).value)
                hd = _to_float(ws.cell(r, c_hd).value)
                pt = _to_float(ws.cell(r, c_pt).value)
                sw = _to_float(ws.cell(r, c_sw).value)
                tot = _to_float(ws.cell(r, c_tot).value) if c_tot else (hb + hd + pt + sw)
                if ma not in week_store_data:
                    week_store_data[ma] = {
                        "days": int(round(dy)),
                        "avg": _to_float(ws.cell(r, c_avg).value),
                        "hamburger": hb,
                        "hotdog": hd,
                        "pita": pt,
                        "sandwich": sw,
                        "total": tot,
                        "bo_th": _to_float(ws.cell(r, c_bo).value),
                        "mieng_ga": _to_float(ws.cell(r, c_ga).value),
                        "mieng_tom": _to_float(ws.cell(r, c_tom).value),
                        "mieng_thit": _to_float(ws.cell(r, c_thit).value),
                        "revenue": 0.0,
                        "debt": 0.0,
                    }
            if in_rev:
                rma = ws.cell(r, rev_c_ma).value
                if isinstance(rma, str) and ("-" in rma):
                    ma = norm_code(rma)
                    rv = _to_float(ws.cell(r, rev_c_rev).value)
                    if ma in week_store_data and week_store_data[ma]["revenue"] == 0:
                        week_store_data[ma]["revenue"] = rv
                    if rev_c_kv:
                        kv = ws.cell(r, rev_c_kv).value
                        if isinstance(kv, str) and kv.strip() and ma not in kv_map:
                            kv_map[ma] = kv.strip()

    for ma, kv in FALLBACK_KV.items():
        kv_map.setdefault(ma, kv)
    for ma, nm in FALLBACK_NAMES.items():
        names_map.setdefault(ma, nm)

    return names_map, kv_map, days_map, week_store_data


def _read_supporter_dashboard_state(
    wb_val: openpyxl.Workbook,
    slots: List[dict],
    target_year: int,
    target_month: int,
    wb_form: Optional[openpyxl.Workbook] = None,
) -> Tuple[dict, Optional[Tuple[int, int]]]:
    """
    Đọc dữ liệu các tuần đã có, Cùng kỳ năm trước, và Mục tiêu tháng từ 5 dashboard quản lý.
    Hỗ trợ cả format gốc (BC_TONG NQ_Tuan45 9-15.xlsx) và format chuẩn mới của tool.
    """
    state = {}
    detected_ym: Optional[Tuple[int, int]] = None

    for sup in SUPPORTERS:
        sname = sup["dash"]
        sup_state = {
            "weeks": {},          # slot_idx -> {hamburger, hotdog, pita, sandwich, avg, bo_th, mieng_ga, mieng_tom, mieng_thit, rev, debt}
            "prev_year_prod": {}, # keys like hamburger...
            "prev_year_rev": 0.0,
            "target_rev": float(sup["default_target"]),
            "prev_week_labels": [],
        }
        if sname not in wb_val.sheetnames:
            state[sup["key"]] = sup_state
            continue

        ws = wb_val[sname]
        ws_f = wb_form[sname] if (wb_form is not None and sname in wb_form.sheetnames) else None
        # Phát hiện format mới vs format cũ
        is_new_fmt = (str(ws.cell(3, 1).value or "").strip() == "TUẦN / NGÀY")
        if is_new_fmt:
            # Đọc tháng/năm từ banner A1/A2 hoặc nhãn tuần A4..A9
            for banner_txt in (str(ws.cell(1, 1).value or ""), str(ws.cell(2, 1).value or "")):
                m_ym = re.search(r"THÁNG\s*(\d{1,2})\s*/\s*(\d{4})", banner_txt, re.IGNORECASE)
                if m_ym and detected_ym is None:
                    detected_ym = (int(m_ym.group(2)), int(m_ym.group(1)))

            for r in range(4, 10):
                lbl = ws.cell(r, 1).value
                if not lbl:
                    continue
                if detected_ym is None:
                    m_ym = re.search(r"(\d{1,2})/(\d{4})", str(lbl))
                    if m_ym:
                        detected_ym = (int(m_ym.group(2)), int(m_ym.group(1)))
                sidx = _match_slot_from_label(str(lbl), slots)
                if sidx:
                    is_uncached = False
                    if ws_f is not None:
                        is_uncached = (
                            (ws.cell(r, 3).value is None and str(ws_f.cell(r, 3).value or "").startswith("="))
                            or (ws.cell(15, 1 + sidx).value is None and str(ws_f.cell(15, 1 + sidx).value or "").startswith("="))
                            or (ws.cell(18, 1 + sidx).value is None and str(ws_f.cell(18, 1 + sidx).value or "").startswith("="))
                        )
                    sup_state["weeks"][sidx] = {
                        "hamburger": _to_float(ws.cell(r, 3).value),
                        "hotdog": _to_float(ws.cell(r, 4).value),
                        "pita": _to_float(ws.cell(r, 5).value),
                        "sandwich": _to_float(ws.cell(r, 6).value),
                        "avg": _to_float(ws.cell(r, 7).value),
                        "bo_th": _to_float(ws.cell(r, 8).value),
                        "mieng_ga": _to_float(ws.cell(r, 9).value),
                        "mieng_tom": _to_float(ws.cell(r, 10).value),
                        "mieng_thit": _to_float(ws.cell(r, 11).value),
                        "rev": _to_float(ws.cell(15, 1 + sidx).value),
                        "debt": _to_float(ws.cell(18, 1 + sidx).value),
                        "_uncached": is_uncached,
                    }
            sup_state["prev_year_prod"] = {
                "hamburger": _to_float(ws.cell(11, 3).value),
                "hotdog": _to_float(ws.cell(11, 4).value),
                "pita": _to_float(ws.cell(11, 5).value),
                "sandwich": _to_float(ws.cell(11, 6).value),
                "avg": _to_float(ws.cell(11, 7).value),
                "bo_th": _to_float(ws.cell(11, 8).value),
                "mieng_ga": _to_float(ws.cell(11, 9).value),
                "mieng_tom": _to_float(ws.cell(11, 10).value),
                "mieng_thit": _to_float(ws.cell(11, 11).value),
            }
            sup_state["target_rev"] = _to_float(ws.cell(16, 9).value) or float(sup["default_target"])
            sup_state["prev_year_rev"] = _to_float(ws.cell(17, 8).value)
        else:
            # Format gốc BC_TONG NQ_Tuan45 9-15.xlsx
            c0 = 4 if sname == "5. KV Duyên" else 1
            for r in range(2, 8):
                lbl = ws.cell(r, c0).value
                if not lbl:
                    continue
                m_ym = re.search(r"(\d{1,2})/(\d{4})", str(lbl))
                if m_ym and detected_ym is None:
                    detected_ym = (int(m_ym.group(2)), int(m_ym.group(1)))
                sidx = _match_slot_from_label(str(lbl), slots)
                if sidx:
                    rev_col = c0 + sidx
                    sup_state["weeks"][sidx] = {
                        "hamburger": _to_float(ws.cell(r, c0 + 2).value),
                        "hotdog": _to_float(ws.cell(r, c0 + 3).value),
                        "pita": _to_float(ws.cell(r, c0 + 4).value),
                        "sandwich": _to_float(ws.cell(r, c0 + 5).value),
                        "avg": _to_float(ws.cell(r, c0 + 6).value),
                        "bo_th": _to_float(ws.cell(r, c0 + 7).value),
                        "mieng_ga": _to_float(ws.cell(r, c0 + 8).value),
                        "mieng_tom": _to_float(ws.cell(r, c0 + 9).value),
                        "mieng_thit": _to_float(ws.cell(r, c0 + 10).value),
                        "rev": _to_float(ws.cell(12, rev_col).value),
                        "debt": _to_float(ws.cell(14 if sname == "5. KV Duyên" else 15, rev_col).value),
                    }
            # Cùng kỳ năm trước (row 9)
            sup_state["prev_year_prod"] = {
                "hamburger": _to_float(ws.cell(9, c0 + 2).value),
                "hotdog": _to_float(ws.cell(9, c0 + 3).value),
                "pita": _to_float(ws.cell(9, c0 + 4).value),
                "sandwich": _to_float(ws.cell(9, c0 + 5).value),
                "avg": _to_float(ws.cell(9, c0 + 6).value),
                "bo_th": _to_float(ws.cell(9, c0 + 7).value),
                "mieng_ga": _to_float(ws.cell(9, c0 + 8).value),
                "mieng_tom": _to_float(ws.cell(9, c0 + 9).value),
                "mieng_thit": _to_float(ws.cell(9, c0 + 10).value),
            }
            sup_state["target_rev"] = _to_float(ws.cell(13, c0).value) or float(sup["default_target"])
            if sname == "5. KV Duyên":
                sup_state["prev_year_rev"] = _to_float(ws.cell(15, c0).value)
            elif sname == "7.KV Quỳnh":
                sup_state["prev_year_rev"] = _to_float(ws.cell(15, c0).value)
            else:
                sup_state["prev_year_rev"] = _to_float(ws.cell(14, c0).value)

        state[sup["key"]] = sup_state

    return state, detected_ym


def _read_existing_history(wb_val: openpyxl.Workbook) -> List[list]:
    if "LichSuBaoCao(tong_hop)" not in wb_val.sheetnames:
        return []
    ws = wb_val["LichSuBaoCao(tong_hop)"]
    if ws.max_row < 2:
        return []
    hdr = [ws.cell(1, c).value for c in range(1, len(HISTORY_HEADERS) + 1)]
    if hdr[0] != "Năm":
        return []
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, len(HISTORY_HEADERS) + 1)]
        if vals[0] is not None and vals[9] is not None:
            rows.append(vals)
    return rows


def _aggregate_history_for_dash(
    history_rows: List[list],
    slots: List[dict],
    target_year: int,
    target_month: int,
) -> dict:
    """
    Tổng hợp số liệu các tuần thuộc (target_year, target_month) từ LichSuBaoCao(tong_hop)
    theo từng quản lý để khôi phục đầy đủ các tuần đã báo cáo ngay cả khi file trước chưa lưu cache công thức từ Excel.
    """
    agg: Dict[str, Dict[int, dict]] = {}
    avg_lists: Dict[Tuple[str, int], List[float]] = defaultdict(list)

    for rw in history_rows:
        if len(rw) < 25:
            continue
        try:
            ry = int(round(_to_float(rw[0])))
            rm = int(round(_to_float(rw[1])))
        except Exception:
            continue
        if ry != target_year or rm != target_month:
            continue

        sidx = _match_slot_from_label(str(rw[3] or ""), slots)
        sup_key = norm_supporter(rw[7])
        if not sidx or sup_key not in SUPPORTER_BY_KEY:
            continue

        sup_weeks = agg.setdefault(sup_key, {})
        bucket = sup_weeks.setdefault(
            sidx,
            {
                "hamburger": 0.0,
                "hotdog": 0.0,
                "pita": 0.0,
                "sandwich": 0.0,
                "avg": 0.0,
                "bo_th": 0.0,
                "mieng_ga": 0.0,
                "mieng_tom": 0.0,
                "mieng_thit": 0.0,
                "rev": 0.0,
                "debt": 0.0,
            },
        )
        bucket["hamburger"] += _to_float(rw[15])
        bucket["hotdog"] += _to_float(rw[16])
        bucket["pita"] += _to_float(rw[17])
        bucket["sandwich"] += _to_float(rw[18])
        bucket["bo_th"] += _to_float(rw[19])
        bucket["mieng_ga"] += _to_float(rw[20])
        bucket["mieng_tom"] += _to_float(rw[21])
        bucket["mieng_thit"] += _to_float(rw[22])
        bucket["rev"] += _to_float(rw[23])
        bucket["debt"] += _to_float(rw[24])

        ngay_kd = _to_float(rw[12])
        tot_banh = _to_float(rw[13])
        row_avg = (tot_banh / ngay_kd) if (ngay_kd > 0 and tot_banh > 0) else _to_float(rw[14])
        if row_avg > 0:
            avg_lists[(sup_key, sidx)].append(row_avg)

    for (sup_key, sidx), vals in avg_lists.items():
        if vals:
            agg[sup_key][sidx]["avg"] = sum(vals) / len(vals)

    return agg


def load_base_workbook_state(base_path: str, year: int, month: int, weekly_info: dict) -> dict:
    """Đọc toàn bộ trạng thái từ file báo cáo tháng gốc để chuẩn bị gộp/làm mới."""
    wb_val = openpyxl.load_workbook(base_path, data_only=True)
    wb_form = openpyxl.load_workbook(base_path, data_only=False)
    slots = get_month_week_slots(year, month)

    legacy_names, legacy_kv, legacy_days, legacy_week_data = _read_legacy_store_info(wb_val)
    dash_state, detected_ym = _read_supporter_dashboard_state(wb_val, slots, year, month, wb_form=wb_form)
    history_rows = _read_existing_history(wb_val)

    # Khôi phục các tuần đã có trong LichSuBaoCao(tong_hop) khi ô công thức trên dashboard chưa có cached value
    hist_dash = _aggregate_history_for_dash(history_rows, slots, year, month)
    for sup in SUPPORTERS:
        sk = sup["key"]
        for sidx, h_data in hist_dash.get(sk, {}).items():
            cur = dash_state[sk]["weeks"].get(sidx)
            is_empty = (
                cur is None
                or cur.get("_uncached", False)
                or (
                    cur.get("hamburger", 0.0) == 0.0
                    and cur.get("hotdog", 0.0) == 0.0
                    and cur.get("pita", 0.0) == 0.0
                    and cur.get("sandwich", 0.0) == 0.0
                    and cur.get("rev", 0.0) == 0.0
                    and cur.get("debt", 0.0) == 0.0
                )
            )
            if is_empty:
                dash_state[sk]["weeks"][sidx] = dict(h_data)

    for st in dash_state.values():
        for wk in st["weeks"].values():
            wk.pop("_uncached", None)

    # Đọc Ds_nhuong_quyen
    stores: "OrderedDict[str, dict]" = OrderedDict()
    ds_name = "Ds_nhuong_quyen" if "Ds_nhuong_quyen" in wb_val.sheetnames else ("DSNQ" if "DSNQ" in wb_val.sheetnames else None)
    if ds_name:
        ws_ds = wb_val[ds_name]
        gn_col = _find_gom_nhom_col(ws_ds)
        for r in range(2, ws_ds.max_row + 1):
            ql_raw = ws_ds.cell(r, 1).value
            ma_raw = ws_ds.cell(r, 2).value
            ten_raw = ws_ds.cell(r, 3).value
            act_raw = ws_ds.cell(r, 4).value
            # Col F nếu đã được tool trước ghi Khu vực/Tỉnh
            kv_raw = ws_ds.cell(r, 6).value if ws_ds.max_column >= 6 else None
            gn_raw = ws_ds.cell(r, gn_col).value if gn_col else None
            if not ma_raw:
                continue
            ma = norm_code(ma_raw)
            ql = norm_supporter(ql_raw)
            ten = (str(ten_raw).strip() if ten_raw else "") or weekly_info["names"].get(ma, "") or legacy_names.get(ma, "")
            kv = (str(kv_raw).strip() if kv_raw else "") or legacy_kv.get(ma, "")
            act = _parse_bool_active(act_raw)
            if act is None:
                # Lần đầu chuyển từ cột D trống: Active nếu có đơn tuần này hoặc tuần trước > 0
                has_curr = ma in weekly_info["pivot"]
                has_prev = legacy_week_data.get(ma, {}).get("total", 0) > 0
                act = bool(has_curr or has_prev)
            stores[ma] = {
                "ql": ql,
                "ma": ma,
                "ten": ten,
                "active": act,
                "kv": kv or "Khác",
                "group": _norm_group_id(gn_raw),
                "is_new": False,
            }

    # Luôn đồng bộ trạng thái hoạt động, quản lý & Gom nhóm từ BC_TONG_base.xlsx (nếu base_path là file tuần khác)
    if os.path.exists(DEFAULT_TEMPLATE) and os.path.abspath(base_path) != os.path.abspath(DEFAULT_TEMPLATE):
        wb_tpl = openpyxl.load_workbook(DEFAULT_TEMPLATE, read_only=True, data_only=True)
        tpl_ds_name = "Ds_nhuong_quyen" if "Ds_nhuong_quyen" in wb_tpl.sheetnames else ("DSNQ" if "DSNQ" in wb_tpl.sheetnames else None)
        if tpl_ds_name:
            ws_tpl = wb_tpl[tpl_ds_name]
            tpl_gn_col = _find_gom_nhom_col(ws_tpl)
            for r in range(2, ws_tpl.max_row + 1):
                ql_raw = ws_tpl.cell(r, 1).value
                ma_raw = ws_tpl.cell(r, 2).value
                ten_raw = ws_tpl.cell(r, 3).value
                act_raw = ws_tpl.cell(r, 4).value
                gn_raw = ws_tpl.cell(r, tpl_gn_col).value if tpl_gn_col else None
                if not ma_raw:
                    continue
                ma = norm_code(ma_raw)
                act = _parse_bool_active(act_raw)
                gid = _norm_group_id(gn_raw)
                if ma in stores:
                    if act is not None:
                        stores[ma]["active"] = act
                    if ql_raw:
                        stores[ma]["ql"] = norm_supporter(ql_raw)
                    if tpl_gn_col:
                        stores[ma]["group"] = gid
                else:
                    ten = (str(ten_raw).strip() if ten_raw else "") or weekly_info["names"].get(ma, "") or legacy_names.get(ma, "")
                    stores[ma] = {
                        "ql": norm_supporter(ql_raw),
                        "ma": ma,
                        "ten": ten,
                        "active": bool(act) if act is not None else False,
                        "kv": legacy_kv.get(ma, "Khác"),
                        "group": gid,
                        "is_new": False,
                    }
        wb_tpl.close()

    # Bổ sung mã NQ có trong file tuần nhưng chưa có trong Ds_nhuong_quyen
    for ma in sorted(weekly_info["pivot"].keys()):
        if ma not in stores:
            stores[ma] = {
                "ql": NGOAI_DM,
                "ma": ma,
                "ten": weekly_info["names"].get(ma, "") or legacy_names.get(ma, "") or ma,
                "active": True,
                "kv": legacy_kv.get(ma, "Khác"),
                "group": "",
                "is_new": True,
            }

    stores = _reorder_grouped_stores(stores)

    # Nếu LichSuBaoCao(tong_hop) đang rỗng (lần đầu nâng cấp từ file gốc Tuan45),
    # tự động seed dữ liệu Tuần 45 (nếu file gốc thuộc cùng tháng/năm) để tra cứu Thu hồi công nợ
    if not history_rows and detected_ym == (year, month) and legacy_week_data:
        # Tìm tuần có dữ liệu gần nhất trước tuần hiện tại trong dash_state
        populated_slots = [sidx for sidx in sorted({s for st in dash_state.values() for s in st["weeks"]})
                           if any(dash_state[sup["key"]]["weeks"].get(sidx, {}).get("hamburger", 0) > 0 for sup in SUPPORTERS)]
        if populated_slots:
            last_sidx = populated_slots[-1]
            sl = slots[last_sidx - 1]
            for ma, st in stores.items():
                wd = legacy_week_data.get(ma)
                if not wd:
                    continue
                kv_tong = get_kv_tong(st["ql"])
                history_rows.append([
                    year, month, sl["week_num"], sl["label"],
                    sl["start_date"].strftime("%Y-%m-%d"), sl["end_date"].strftime("%Y-%m-%d"),
                    kv_tong, st["ql"], st["kv"], ma, st["ten"],
                    "Có" if st["active"] else "Không",
                    wd["days"], wd["total"], round(wd["avg"], 2),
                    wd["hamburger"], wd["hotdog"], wd["pita"], wd["sandwich"],
                    wd["bo_th"], wd["mieng_ga"], wd["mieng_tom"], wd["mieng_thit"],
                    wd["revenue"], wd["debt"],
                ])

    wb_form.close()
    wb_val.close()
    return {
        "stores": stores,
        "dash_state": dash_state,
        "detected_ym": detected_ym,
        "history_rows": history_rows,
        "legacy_days": legacy_days,
    }


# --------------------------------------------------------------------------------------
# GHI CÁC SHEET EXCEL (CLEAN, FORMULA-DRIVEN, PROFESSIONAL)
# --------------------------------------------------------------------------------------
def _style_cell(cell, font=FONT_DATA, fill=None, border=BORDER_CELL, align=None, num_fmt=None):
    cell.font = font
    if fill is not None:
        cell.fill = fill
    if border is not None:
        cell.border = border
    if align is not None:
        cell.alignment = align
    if num_fmt is not None:
        cell.number_format = num_fmt


def _make_banner(ws, max_col: int, title: str, subtitle: str = ""):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c = ws.cell(1, 1, title)
    _style_cell(c, font=FONT_TITLE, fill=FILL_NAVY, border=None,
                align=Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[1].height = 30
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        s = ws.cell(2, 1, subtitle)
        _style_cell(s, font=FONT_MUTED, fill=FILL_ZEBRA, border=None,
                    align=Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[2].height = 18


def _chart_text_props(size_pt: float = 9.0, bold: bool = False, color: str = "404040") -> RichText:
    """Tạo thuộc tính font DrawingML đồng bộ cho tiêu đề, trục, nhãn dữ liệu và chú giải biểu đồ."""
    sz = int(round(size_pt * 100))
    return RichText(
        p=[
            Paragraph(
                pPr=ParagraphProperties(
                    defRPr=CharacterProperties(
                        sz=sz,
                        b=bold,
                        solidFill=color,
                        latin=DrawingFont(typeface=FONT_NAME),
                    )
                ),
                endParaRPr=CharacterProperties(sz=sz),
            )
        ]
    )


def _apply_title(chart_or_axis, text: str, size_pt: float = 11.0, bold: bool = True, color: str = "1F4E78"):
    """Đặt tiêu đề biểu đồ/trục với overlay=False để không đè lên vùng vẽ hay nhãn."""
    chart_or_axis.title = text
    chart_or_axis.title.overlay = False
    sz = int(round(size_pt * 100))
    chart_or_axis.title.tx.rich.p[0].pPr = ParagraphProperties(
        defRPr=CharacterProperties(
            sz=sz,
            b=bold,
            solidFill=color,
            latin=DrawingFont(typeface=FONT_NAME),
        )
    )


def _make_data_labels(
    show_val: bool = True,
    pos: Optional[str] = None,
    num_fmt: Optional[str] = None,
    size_pt: float = 8.5,
    color: str = "1F4E78",
) -> DataLabelList:
    """Tạo DataLabelList với đầy đủ cờ tường minh tránh Excel hiển thị tên series/ô vuông legend."""
    dl = DataLabelList()
    dl.showVal = show_val
    dl.showLegendKey = False
    dl.showCatName = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showBubbleSize = False
    dl.showLeaderLines = False
    dl.txPr = _chart_text_props(size_pt=size_pt, bold=True, color=color)
    if pos:
        dl.dLblPos = pos
    if num_fmt:
        dl.numFmt = num_fmt
    return dl


def _style_common_chart_frame(chart, height: float, width: float):
    """Chuẩn hóa khung, nền trắng và viền mảnh hiện đại cho biểu đồ."""
    chart.height = height
    chart.width = width
    chart.roundedCorners = False
    chart.graphical_properties = GraphicalProperties(
        solidFill="FFFFFF",
        ln=LineProperties(solidFill="D9E1F2", w=9525),
    )
    chart.plot_area.graphicalProperties = GraphicalProperties(
        solidFill="FFFFFF",
        ln=LineProperties(noFill=True),
    )


def _style_bar_chart(
    chart: BarChart,
    title: str,
    y_title: str = "",
    stacked: bool = False,
    horizontal: bool = False,
    show_labels: bool = True,
    height: float = 10.0,
    width: float = 16.0,
    y_num_fmt: Optional[str] = None,
    label_num_fmt: Optional[str] = None,
    show_legend: bool = True,
):
    chart.type = "bar" if horizontal else "col"
    chart.style = 10
    _apply_title(chart, title, size_pt=11.0, bold=True, color="1F4E78")
    chart.grouping = "stacked" if stacked else "clustered"
    chart.overlap = 100 if stacked else -15
    chart.gapWidth = 60 if horizontal else 75
    _style_common_chart_frame(chart, height, width)

    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if horizontal:
        chart.x_axis.axPos = "l"
        chart.y_axis.axPos = "b"
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"
    else:
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.tickLblPos = "nextTo"

    chart.y_axis.scaling.min = 0
    chart.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill="D9E1F2", w=6350)
    )
    chart.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="BFBFBF", w=6350))
    chart.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.x_axis.txPr = _chart_text_props(size_pt=8.5, color="404040")
    chart.y_axis.txPr = _chart_text_props(size_pt=8.5, color="404040")

    if y_title:
        _apply_title(chart.y_axis, y_title, size_pt=9.0, bold=True, color="595959")
    if y_num_fmt:
        chart.y_axis.number_format = y_num_fmt
        chart.y_axis.sourceLinked = False

    if show_legend:
        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = _chart_text_props(size_pt=8.5, color="404040")
    else:
        chart.legend = None

    if show_labels and not stacked:
        chart.dataLabels = _make_data_labels(
            show_val=True,
            pos="outEnd",
            num_fmt=label_num_fmt,
            size_pt=8.5,
            color="1F4E78",
        )
    else:
        chart.dataLabels = None

    for idx, ser in enumerate(chart.series):
        col = CHART_COLORS[idx % len(CHART_COLORS)]
        ser.graphicalProperties.solidFill = col
        ser.graphicalProperties.line.solidFill = col


def _style_line_chart(
    chart: LineChart,
    title: str,
    y_title: str = "",
    height: float = 10.0,
    width: float = 16.0,
    label_num_fmt: str = '0.0;;""',
    show_legend: bool = False,
):
    chart.style = 10
    chart.display_blanks = "gap"
    chart.varyColors = False
    _apply_title(chart, title, size_pt=11.0, bold=True, color="1F4E78")
    _style_common_chart_frame(chart, height, width)

    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.scaling.min = 0
    chart.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill="D9E1F2", w=6350)
    )
    chart.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="BFBFBF", w=6350))
    chart.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.x_axis.txPr = _chart_text_props(size_pt=8.5, color="404040")
    chart.y_axis.txPr = _chart_text_props(size_pt=8.5, color="404040")

    if y_title:
        _apply_title(chart.y_axis, y_title, size_pt=9.0, bold=True, color="595959")

    if show_legend:
        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = _chart_text_props(size_pt=8.5, color="404040")
    else:
        chart.legend = None

    chart.dataLabels = _make_data_labels(
        show_val=True,
        pos="t",
        num_fmt=label_num_fmt,
        size_pt=8.5,
        color="1F4E78",
    )
    for idx, ser in enumerate(chart.series):
        col = CHART_COLORS[idx % len(CHART_COLORS)]
        ser.smooth = False
        ser.graphicalProperties.line.solidFill = col
        ser.graphicalProperties.line.width = 25400
        ser.marker.symbol = "circle"
        ser.marker.size = 7
        ser.marker.graphicalProperties.solidFill = col
        ser.marker.graphicalProperties.line.solidFill = "FFFFFF"


def write_ds_nhuong_quyen(wb: openpyxl.Workbook, stores: OrderedDict) -> int:
    ws = wb.create_sheet("Ds_nhuong_quyen")
    ws.views.sheetView[0].showGridLines = True
    headers = ["Quản lý", "Mã NQ", "NHƯỢNG QUYỀN", "Còn hoạt động?", "Trạng thái", "Khu vực / Tỉnh", "KV Tổng", "Gom nhóm"]
    widths = [18, 12, 62, 16, 18, 18, 16, 12]
    for idx, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, idx, h)
        _style_cell(c, font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.row_dimensions[1].height = 28

    group_color_map = {}
    for row_idx, (ma, st) in enumerate(stores.items(), start=2):
        kv_tong = get_kv_tong(st["ql"])
        grp_fill = _get_group_fill(st.get("group", ""), group_color_map)
        row_fill = grp_fill or (FILL_AMBER if st.get("is_new") else (FILL_ZEBRA if row_idx % 2 == 0 else FILL_WHITE))

        _style_cell(ws.cell(row_idx, 1, st["ql"]), font=FONT_BOLD if st.get("is_new") else FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(row_idx, 2, ma), font=FONT_BOLD, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(row_idx, 3, st["ten"]), font=FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="left", vertical="center"))

        # Col D lưu boolean TRUE/FALSE liên kết trực tiếp với Form Control CheckBox (ẩn chữ bằng ';;;')
        c_act = ws.cell(row_idx, 4, bool(st["active"]))
        _style_cell(c_act, font=FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"), num_fmt=FMT_HIDDEN)

        c_status = ws.cell(row_idx, 5, f'=IF(D{row_idx},"Đang hoạt động","Ngưng / Off")')
        _style_cell(c_status, font=FONT_GREEN if st["active"] else FONT_MUTED,
                    fill=FILL_GREEN if st["active"] else FILL_GREY,
                    align=Alignment(horizontal="center", vertical="center"))

        _style_cell(ws.cell(row_idx, 6, st["kv"]), font=FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(row_idx, 7, kv_tong), font=FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(row_idx, 8, st.get("group") or None), font=FONT_DATA, fill=row_fill,
                    align=Alignment(horizontal="center", vertical="center"))
        ws.row_dimensions[row_idx].height = 20

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{1 + len(stores)}"
    return 1 + len(stores)


def write_du_lieu_tho(wb: openpyxl.Workbook, weekly_info: dict, stores: OrderedDict) -> Tuple[int, dict]:
    ws = wb.create_sheet("DuLieuTho(tuan_nay)")
    pivot = weekly_info["pivot"]
    rev = weekly_info["revenue"]
    item_names = weekly_info["item_names"]

    # Các cột chuẩn ở đầu để detail sheets INDEX/MATCH cố định, đáng tin cậy
    fixed_codes = [
        ("1001", "Hamburger (1001)"),
        ("1004", "Hotdog (1004)"),
        ("1003", "Pita (1003)"),
        ("1002", "Sandwich miếng (1002)"),
        # Col J (10) là Sandwich phần = I/2
        ("1214", "Bò thượng hạng (1214)"),
        ("1200", "Miếng gà (1200)"),
        ("1203", "Miếng tôm (1203)"),
        ("1201", "Miếng thịt (1201)"),
    ]
    fixed_set = {c for c, _ in fixed_codes}
    other_codes = sorted(c for c in item_names.keys() if c not in fixed_set)

    headers = [
        "Mã NQ", "NHƯỢNG QUYỀN", "Quản lý", "Khu vực / Tỉnh",
        "Tổng tiền gốc (VNĐ)", "Doanh thu (+4%) (VNĐ)",
        "Hamburger (1001)", "Hotdog (1004)", "Pita (1003)", "Sandwich miếng (1002)", "Sandwich phần (÷2)",
        "Bò thượng hạng (1214)", "Miếng gà (1200)", "Miếng tôm (1203)", "Miếng thịt (1201)",
    ] + [f"{item_names[c]} ({c})" for c in other_codes]

    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(1, col_idx, h)
        _style_cell(c, font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[get_column_letter(col_idx)].width = (
            62 if col_idx == 2 else (18 if col_idx in (5, 6) else (16 if col_idx <= 15 else 18))
        )
    ws.row_dimensions[1].height = 36

    ordered_codes = sorted(pivot.keys())
    for idx, ma in enumerate(ordered_codes, start=2):
        st = stores.get(ma, {})
        ten = st.get("ten") or weekly_info["names"].get(ma, ma)
        ql = st.get("ql", NGOAI_DM)
        kv = st.get("kv", "Khác")
        fill = FILL_ZEBRA if idx % 2 == 0 else FILL_WHITE
        p_row = pivot[ma]

        _style_cell(ws.cell(idx, 1, ma), font=FONT_BOLD, fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(idx, 2, ten), font=FONT_DATA, fill=fill, align=Alignment(horizontal="left", vertical="center"))
        _style_cell(ws.cell(idx, 3, ql), font=FONT_DATA, fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(idx, 4, kv), font=FONT_DATA, fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(idx, 5, rev.get(ma, 0.0)), font=FONT_DATA, fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 6, f"=ROUND(E{idx}*1.04,0)"), font=FONT_BOLD, fill=fill, num_fmt=FMT_INT)

        _style_cell(ws.cell(idx, 7, p_row.get("1001", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 8, p_row.get("1004", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 9, p_row.get("1003", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 10, p_row.get("1002", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 11, f"=J{idx}/2"), font=FONT_BOLD, fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(idx, 12, p_row.get("1214", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 13, p_row.get("1200", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 14, p_row.get("1203", 0.0)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(idx, 15, p_row.get("1201", 0.0)), fill=fill, num_fmt=FMT_INT)

        for j, code in enumerate(other_codes, start=16):
            val = p_row.get(code, 0.0)
            _style_cell(ws.cell(idx, j, val if val else 0.0), fill=fill, num_fmt=FMT_DEC if val % 1 else FMT_INT)
        ws.row_dimensions[idx].height = 19

    # Dòng TỔNG CỘNG
    tot_row = 2 + len(ordered_codes)
    _style_cell(ws.cell(tot_row, 1, "TỔNG CỘNG"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL,
                align=Alignment(horizontal="center", vertical="center"))
    _style_cell(ws.cell(tot_row, 2, f'=COUNTA(A2:A{tot_row - 1})&" cửa hàng có đơn"'), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(tot_row, 3, ""), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(tot_row, 4, ""), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    for col_idx in range(5, len(headers) + 1):
        cl = get_column_letter(col_idx)
        _style_cell(ws.cell(tot_row, col_idx, f"=SUM({cl}2:{cl}{tot_row - 1})"),
                    font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_INT)
    ws.row_dimensions[tot_row].height = 22
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{tot_row - 1}"

    col_map = {
        "rev_raw": "E",
        "rev": "F",
        "hamburger": "G",
        "hotdog": "H",
        "pita": "I",
        "sandwich_phan": "K",
        "bo_th": "L",
        "mieng_ga": "M",
        "mieng_tom": "N",
        "mieng_thit": "O",
    }
    return tot_row - 1, col_map


def write_supporter_detail_sheet(
    wb: openpyxl.Workbook,
    sup: dict,
    stores: OrderedDict,
    weekly_info: dict,
    raw_col_map: dict,
    slot: dict,
    year: int,
    month: int,
    prev_week_label: Optional[str],
    legacy_days: dict,
    ds_last_row: int,
    raw_last_data_row: int,
    hist_last_row: int,
    prev_banh_map: Optional[dict] = None,
) -> dict:
    """Tạo sheet chi tiết tuần của từng quản lý với công thức động nối thẳng DuLieuTho, Ds_nhuong_quyen và LichSuBaoCao."""
    ws = wb.create_sheet(sup["detail"])
    ws.views.sheetView[0].showGridLines = True

    sup_stores = [st for st in stores.values() if st["ql"] == sup["key"]]
    # Sắp xếp theo tổng bánh tuần này giảm dần; các cửa hàng cùng Gom nhóm luôn đứng liền kề nhau
    piv = weekly_info["pivot"]
    def _week_tot(ma: str) -> float:
        p = piv.get(ma, {})
        return p.get("1001", 0) + p.get("1004", 0) + p.get("1003", 0) + p.get("1002", 0) / 2.0
    store_order = {ma: idx for idx, ma in enumerate(stores.keys())}
    group_max_tot = defaultdict(float)
    group_first_idx = {}
    for s in sup_stores:
        gid = s.get("group", "")
        if gid:
            group_max_tot[gid] = max(group_max_tot[gid], _week_tot(s["ma"]))
            if gid not in group_first_idx:
                group_first_idx[gid] = store_order.get(s["ma"], 0)

    def _sort_key(s: dict):
        gid = s.get("group", "")
        if gid:
            return (-group_max_tot[gid], group_first_idx[gid], store_order.get(s["ma"], 0))
        return (-_week_tot(s["ma"]), store_order.get(s["ma"], 0), s["ma"])
    sup_stores.sort(key=_sort_key)

    _make_banner(
        ws, 18,
        f"CHI TIẾT DOANH SỐ TUẦN - QUẢN LÝ: {sup['key'].upper()} ({sup['kv']})",
        f"{slot['label']}  |  Tháng {month:02d}/{year}  |  Dữ liệu tự động liên kết DuLieuTho(tuan_nay) & Ds_nhuong_quyen",
    )

    headers = [
        "STT", "Mã NQ", "NHƯỢNG QUYỀN", "Khu vực / Tỉnh", "Hoạt động?",
        "NGÀY KD", "BÌNH QUÂN/NGÀY", "HAMBURGER", "HOTDOG", "PITA",
        "SANDWICH (÷2)", "TỔNG BÁNH", "BÒ THƯỢNG HẠNG", "MIẾNG GÀ",
        "MIẾNG TÔM", "MIẾNG THỊT", "DOANH THU (VNĐ)", "THU HỒI CÔNG NỢ (VNĐ)",
    ]
    widths = [6, 15, 62, 16, 16, 10, 15, 13, 12, 16, 14, 13, 15, 12, 12, 12, 18, 20]
    for col_idx, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(3, col_idx, h)
        _style_cell(c, font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[3].height = 34

    n_stores = len(sup_stores)
    start_row = 5
    end_row = start_row + max(1, n_stores) - 1

    # Row 4: TỔNG CỘNG
    _style_cell(ws.cell(4, 1, "TỔNG"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL,
                align=Alignment(horizontal="center", vertical="center"))
    _style_cell(ws.cell(4, 2, f'=COUNTIF(E5:E{end_row},"Đang hoạt động")&"/"&COUNTA(B5:B{end_row})'),
                font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, align=Alignment(horizontal="center", vertical="center"))
    _style_cell(ws.cell(4, 3, "Tổng hợp cửa hàng thuộc quản lý"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(4, 4, ""), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(4, 5, ""), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(4, 6, f"=IFERROR(AVERAGEIF(F5:F{end_row},\">0\"),0)"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_DEC)
    _style_cell(ws.cell(4, 7, f"=IFERROR(AVERAGEIF(G5:G{end_row},\">0\"),0)"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_DEC)
    for col_idx in range(8, 19):
        cl = get_column_letter(col_idx)
        _style_cell(ws.cell(4, col_idx, f"=SUM({cl}5:{cl}{end_row})"),
                    font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_INT)
    ws.row_dimensions[4].height = 24

    ds_rng_ma = f"Ds_nhuong_quyen!$B$2:$B${ds_last_row}"
    ds_rng_ten = f"Ds_nhuong_quyen!$C$2:$C${ds_last_row}"
    ds_rng_act = f"Ds_nhuong_quyen!$D$2:$D${ds_last_row}"
    ds_rng_kv = f"Ds_nhuong_quyen!$F$2:$F${ds_last_row}"
    raw_rng_ma = f"'DuLieuTho(tuan_nay)'!$A$2:$A${raw_last_data_row}"

    def _raw_lookup(col_letter: str, row_num: int) -> str:
        return f"=IFERROR(INDEX('DuLieuTho(tuan_nay)'!${col_letter}$2:${col_letter}${raw_last_data_row},MATCH(B{row_num},{raw_rng_ma},0)),0)"

    group_color_map = {}
    for idx, st in enumerate(sup_stores):
        r = start_row + idx
        ma = st["ma"]
        fill = _get_group_fill(st.get("group", ""), group_color_map) or (FILL_ZEBRA if r % 2 == 0 else FILL_WHITE)
        has_order = ma in piv
        # Số ngày KD: từ cột ngày thực tế (nếu có), else số ngày của slot tuần được chọn (nếu có đơn), else 0
        if has_order:
            ngay_kd = weekly_info["days"].get(ma) or slot["days_count"]
        else:
            ngay_kd = 0

        _style_cell(ws.cell(r, 1, idx + 1), fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(r, 2, ma), font=FONT_BOLD, fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(r, 3, f'=IFERROR(INDEX({ds_rng_ten},MATCH(B{r},{ds_rng_ma},0)),"")'), fill=fill)
        _style_cell(ws.cell(r, 4, f'=IFERROR(INDEX({ds_rng_kv},MATCH(B{r},{ds_rng_ma},0)),"")'), fill=fill,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(r, 5, f'=IF(IFERROR(INDEX({ds_rng_act},MATCH(B{r},{ds_rng_ma},0)),FALSE),"Đang hoạt động","Ngưng / Off")'),
                    fill=fill, align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(r, 6, ngay_kd), font=FONT_INPUT, fill=fill, num_fmt=FMT_INT,
                    align=Alignment(horizontal="center", vertical="center"))
        _style_cell(ws.cell(r, 7, f"=IF(F{r}>0,L{r}/F{r},0)"), font=FONT_BOLD, fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(r, 8, _raw_lookup(raw_col_map["hamburger"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 9, _raw_lookup(raw_col_map["hotdog"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 10, _raw_lookup(raw_col_map["pita"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 11, _raw_lookup(raw_col_map["sandwich_phan"], r)), fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(r, 12, f"=SUM(H{r}:K{r})"), font=FONT_BOLD, fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(r, 13, _raw_lookup(raw_col_map["bo_th"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 14, _raw_lookup(raw_col_map["mieng_ga"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 15, _raw_lookup(raw_col_map["mieng_tom"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 16, _raw_lookup(raw_col_map["mieng_thit"], r)), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(r, 17, _raw_lookup(raw_col_map["rev"], r)), font=FONT_BOLD, fill=fill, num_fmt=FMT_INT)

        if prev_week_label and hist_last_row >= 2:
            debt_f = (
                f"=ROUND(SUMIFS('LichSuBaoCao(tong_hop)'!$X$2:$X${hist_last_row},"
                f"'LichSuBaoCao(tong_hop)'!$J$2:$J${hist_last_row},B{r},"
                f"'LichSuBaoCao(tong_hop)'!$D$2:$D${hist_last_row},\"{prev_week_label}\")*0.96,0)"
            )
        else:
            debt_f = "=0"
        _style_cell(ws.cell(r, 18, debt_f), fill=fill, num_fmt=FMT_INT)
        ws.row_dimensions[r].height = 19

    # Block Tóm tắt KPI bên phải (Col T..U)
    ws.column_dimensions["S"].width = 3
    ws.column_dimensions["T"].width = 26
    ws.column_dimensions["U"].width = 68

    _style_cell(ws.cell(3, 20, "TÓM TẮT CHỈ SỐ TUẦN"), font=FONT_HDR, fill=FILL_SUBHDR,
                align=Alignment(horizontal="center", vertical="center"))
    _style_cell(ws.cell(3, 21, "GIÁ TRỊ"), font=FONT_HDR, fill=FILL_SUBHDR,
                align=Alignment(horizontal="center", vertical="center"))

    # Tính NQ BQ thấp nhất (>0) trực tiếp bằng Python để không lỗi công thức mảng trên Excel cũ
    min_bq_val = None
    min_bq_name = ""
    for st in sup_stores:
        ma = st["ma"]
        if ma not in piv:
            continue
        p = piv[ma]
        tot_b = p.get("1001", 0) + p.get("1004", 0) + p.get("1003", 0) + p.get("1002", 0) / 2.0
        ngay = weekly_info["days"].get(ma) or slot["days_count"]
        if ngay > 0 and tot_b > 0:
            bq = tot_b / ngay
            if min_bq_val is None or bq < min_bq_val:
                min_bq_val = bq
                min_bq_name = st.get("ten") or ma
    min_bq_display = f"{min_bq_name} ({round(min_bq_val)} phần)" if min_bq_val is not None else ""

    kpis = [
        ("Tổng số cửa hàng", f'=COUNTIF(Ds_nhuong_quyen!$A$2:$A${ds_last_row},"{sup["key"]}")', FMT_INT),
        ("Số lượng NQ hoạt động", f'=COUNTIFS(Ds_nhuong_quyen!$A$2:$A${ds_last_row},"{sup["key"]}",Ds_nhuong_quyen!$D$2:$D${ds_last_row},TRUE)', FMT_INT),
        ("Số lượng NQ đang off", f'=COUNTIFS(Ds_nhuong_quyen!$A$2:$A${ds_last_row},"{sup["key"]}",Ds_nhuong_quyen!$D$2:$D${ds_last_row},FALSE)', FMT_INT),
        ("Số NQ có đơn tuần này", f'=COUNTIF(L{start_row}:L{end_row},">0")', FMT_INT),
        ("NQ BQ cao nhất", f'=IFERROR(INDEX(C{start_row}:C{end_row},MATCH(MAX(G{start_row}:G{end_row}),G{start_row}:G{end_row},0))&" ("&ROUND(MAX(G{start_row}:G{end_row}),0)&" phần)","")', None),
        ("NQ BQ thấp nhất (>0)", min_bq_display, None),
        ("BQ trung bình (>0)", f"=G4", FMT_DEC),
    ]
    for idx, (lbl, formula, fmt) in enumerate(kpis, start=4):
        fill = FILL_ZEBRA if idx % 2 == 0 else FILL_WHITE
        _style_cell(ws.cell(idx, 20, lbl), font=FONT_BOLD, fill=fill)
        _style_cell(ws.cell(idx, 21, formula), font=FONT_BOLD, fill=fill, num_fmt=fmt,
                    align=Alignment(horizontal="right" if fmt else "left", vertical="center"))

    # Bảng Tổng hợp theo Khu vực / Tỉnh bên dưới
    kv_list = sorted({st["kv"] for st in sup_stores if st["kv"]})
    kv_hdr_row = end_row + 3
    _style_cell(ws.cell(kv_hdr_row - 1, 2, "TỔNG HỢP THEO KHU VỰC / TỈNH"), font=FONT_SECTION, border=None)
    ws.merge_cells(start_row=kv_hdr_row - 1, start_column=2, end_row=kv_hdr_row - 1, end_column=10)
    kv_headers = ["Khu vực / Tỉnh", "Số NQ", "BQ trung bình", "Hamburger", "Hotdog", "Pita", "Sandwich (÷2)", "Tổng bánh", "Doanh thu (VNĐ)"]
    for j, h in enumerate(kv_headers, start=2):
        _style_cell(ws.cell(kv_hdr_row, j, h), font=FONT_HDR, fill=FILL_SUBHDR,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[kv_hdr_row].height = 26

    kv_start = kv_hdr_row + 1
    for i, kv_name in enumerate(kv_list):
        kr = kv_start + i
        fill = FILL_ZEBRA if kr % 2 == 0 else FILL_WHITE
        _style_cell(ws.cell(kr, 2, kv_name), font=FONT_BOLD, fill=fill)
        _style_cell(ws.cell(kr, 3, f'=COUNTIF($D${start_row}:$D${end_row},B{kr})'), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(kr, 4, f'=IFERROR(AVERAGEIFS($G${start_row}:$G${end_row},$D${start_row}:$D${end_row},B{kr},$G${start_row}:$G${end_row},">0"),0)'), fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(kr, 5, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$H${start_row}:$H${end_row})'), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(kr, 6, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$I${start_row}:$I${end_row})'), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(kr, 7, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$J${start_row}:$J${end_row})'), fill=fill, num_fmt=FMT_INT)
        _style_cell(ws.cell(kr, 8, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$K${start_row}:$K${end_row})'), fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(kr, 9, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$L${start_row}:$L${end_row})'), font=FONT_BOLD, fill=fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(kr, 10, f'=SUMIF($D${start_row}:$D${end_row},B{kr},$Q${start_row}:$Q${end_row})'), font=FONT_BOLD, fill=fill, num_fmt=FMT_INT)

    kv_end = kv_start + max(1, len(kv_list)) - 1
    kv_tot_row = kv_end + 1
    _style_cell(ws.cell(kv_tot_row, 2, "TỔNG"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    _style_cell(ws.cell(kv_tot_row, 3, f"=SUM(C{kv_start}:C{kv_end})"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_INT)
    _style_cell(ws.cell(kv_tot_row, 4, "=G4"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_DEC)
    for j in range(5, 11):
        cl = get_column_letter(j)
        _style_cell(ws.cell(kv_tot_row, j, f"=SUM({cl}{kv_start}:{cl}{kv_end})"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_INT)

    # 2 Biểu đồ sạch, dễ đọc trên sheet chi tiết
    # Chart 1: Horizontal bar Top 12 cửa hàng theo Doanh thu tuần (bỏ dòng 4 TỔNG)
    top_end = min(end_row, start_row + 11)
    ch1 = BarChart()
    vals1 = Reference(ws, min_col=17, min_row=start_row, max_row=top_end)
    cats1 = Reference(ws, min_col=3, min_row=start_row, max_row=top_end)
    s1 = Series(vals1, title="Doanh thu tuần")
    ch1.series.append(s1)
    ch1.set_categories(cats1)
    _style_bar_chart(
        ch1,
        f"Top 12 Nhượng Quyền theo Doanh thu tuần ({sup['key']})",
        y_title="Triệu VNĐ",
        horizontal=True,
        show_labels=True,
        height=14.0,
        width=18.5,
        y_num_fmt='#,##0,,',
        label_num_fmt='#,##0.0,," Tr";;""',
        show_legend=False,
    )
    ws.add_chart(ch1, "T12")

    # Chart 2: Stacked Column theo Khu vực / Tỉnh
    ch2 = BarChart()
    data2 = Reference(ws, min_col=5, max_col=8, min_row=kv_hdr_row, max_row=kv_end)
    cats2 = Reference(ws, min_col=2, min_row=kv_start, max_row=kv_end)
    ch2.add_data(data2, titles_from_data=True)
    ch2.set_categories(cats2)
    _style_bar_chart(
        ch2,
        f"Cơ cấu sản lượng bánh theo Khu vực ({sup['key']})",
        y_title="Số phần bánh",
        stacked=True,
        show_labels=False,
        height=11.0,
        width=18.0,
        y_num_fmt='#,##0',
        show_legend=True,
    )
    ws.add_chart(ch2, f"B{kv_tot_row + 2}")

    # Block So sánh Tổng bánh Tuần này vs Tuần trước (Col Z..AC, cách biểu đồ T..U qua vùng đệm V..Y)
    for sp_col in ("V", "W", "X", "Y"):
        ws.column_dimensions[sp_col].width = 4
    cmp_start_col = 26  # Z
    cmp_headers = ["TUẦN TRƯỚC (BÁNH)", "TUẦN NÀY (BÁNH)", "CHÊNH LỆCH (+/-)", "TĂNG / GIẢM"]
    cmp_widths = [18, 17, 17, 20]
    for offset, (ch, cw) in enumerate(zip(cmp_headers, cmp_widths)):
        col_i = cmp_start_col + offset
        _style_cell(ws.cell(3, col_i, ch), font=FONT_HDR, fill=FILL_SUBHDR,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[get_column_letter(col_i)].width = cw

    _style_cell(ws.cell(4, 26, f"=SUM(Z{start_row}:Z{end_row})"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_DEC)
    _style_cell(ws.cell(4, 27, f"=SUM(AA{start_row}:AA{end_row})"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=FMT_DEC)
    _style_cell(ws.cell(4, 28, "=AA4-Z4"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt="+#,##0;-#,##0;0")
    _style_cell(ws.cell(4, 29, '=IF(AB4>0,"Tăng (+"&TEXT(AB4,"#,##0")&")",IF(AB4<0,"Giảm ("&TEXT(AB4,"#,##0")&")","Không đổi"))'),
                font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, align=Alignment(horizontal="center", vertical="center"))

    last_active_row = start_row
    for idx, st in enumerate(sup_stores):
        r = start_row + idx
        prev_b = (prev_banh_map or {}).get(st["ma"], 0.0)
        if st["ma"] in piv or prev_b > 0:
            last_active_row = r
        r_fill = _get_group_fill(st.get("group", ""), group_color_map) or (FILL_ZEBRA if r % 2 == 0 else FILL_WHITE)
        if prev_week_label and hist_last_row >= 2:
            prev_banh_f = (
                f"=SUMIFS('LichSuBaoCao(tong_hop)'!$N$2:$N${hist_last_row},"
                f"'LichSuBaoCao(tong_hop)'!$J$2:$J${hist_last_row},B{r},"
                f"'LichSuBaoCao(tong_hop)'!$D$2:$D${hist_last_row},\"{prev_week_label}\")"
            )
        else:
            prev_banh_f = "=0"
        _style_cell(ws.cell(r, 26, prev_banh_f), fill=r_fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(r, 27, f"=L{r}"), font=FONT_BOLD, fill=r_fill, num_fmt=FMT_DEC)
        _style_cell(ws.cell(r, 28, f"=AA{r}-Z{r}"), font=FONT_BOLD, fill=r_fill, num_fmt="+#,##0;-#,##0;0")
        _style_cell(
            ws.cell(r, 29, f'=IF(AB{r}>0,"Tăng (+"&TEXT(AB{r},"#,##0")&")",IF(AB{r}<0,"Giảm ("&TEXT(AB{r},"#,##0")&")","Không đổi"))'),
            fill=r_fill, align=Alignment(horizontal="center", vertical="center"),
        )
        # Cột phụ ẩn AD/AE tách riêng Tăng (>=0) và Giảm (<=0) để Excel tô xanh/đỏ chắc chắn 100%
        ws.cell(r, 30, f"=IF(AB{r}>0,AB{r},0)")
        ws.cell(r, 31, f"=IF(AB{r}<0,AB{r},0)")
    ws.column_dimensions["AD"].hidden = True
    ws.column_dimensions["AE"].hidden = True

    # Chart 3: Biểu đồ chênh lệch Tổng bánh (Tăng xanh lá / Giảm đỏ) giữa Tuần này và Tuần trước
    chart_end_row = min(end_row, max(start_row + 11, last_active_row))
    chart_n = chart_end_row - start_row + 1
    ch3 = BarChart()
    cats3 = Reference(ws, min_col=3, min_row=start_row, max_row=chart_end_row)
    s_inc = Series(Reference(ws, min_col=30, min_row=start_row, max_row=chart_end_row), title="Tăng (+)")
    s_dec = Series(Reference(ws, min_col=31, min_row=start_row, max_row=chart_end_row), title="Giảm (-)")
    ch3.series.extend([s_inc, s_dec])
    ch3.set_categories(cats3)
    _style_bar_chart(
        ch3,
        f"Chênh lệch Tổng bánh: Tuần này vs Tuần trước ({sup['key']})",
        y_title="Số phần bánh (+/-)",
        stacked=False,
        horizontal=True,
        show_labels=True,
        height=max(13.0, chart_n * 0.62),
        width=18.5,
        y_num_fmt='+#,##0;-#,##0;0',
        label_num_fmt='+#,##0;-#,##0;""',
        show_legend=True,
    )
    # Đặt lại màu & overlap sau _style_bar_chart: Tăng = xanh lá (#16A34A), Giảm = đỏ (#DC2626)
    s_inc.graphicalProperties.solidFill = "16A34A"
    s_inc.graphicalProperties.line.solidFill = "16A34A"
    s_dec.graphicalProperties.solidFill = "DC2626"
    s_dec.graphicalProperties.line.solidFill = "DC2626"
    s_dec.invertIfNegative = False
    ch3.overlap = 100
    ch3.visible_cells_only = False
    ch3.gapWidth = 45
    ch3.y_axis.scaling.min = None
    ch3.x_axis.crosses = "autoZero"
    ch3.y_axis.crosses = "max"
    ch3.x_axis.tickLblPos = "low"
    ws.add_chart(ch3, "T41")

    ws.freeze_panes = "D5"
    return {"start_row": start_row, "end_row": end_row, "total_row": 4}


def _write_dashboard_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    title: str,
    slots: List[dict],
    current_slot_idx: int,
    year: int,
    month: int,
    level: str,  # "supporter" | "region" | "total"
    sup_key: Optional[str] = None,
    child_sheets: Optional[List[str]] = None,
    sup_dash_state: Optional[dict] = None,
    reset_month: bool = False,
    detail_sheet_name: Optional[str] = None,
    prev_month_last_rev: float = 0.0,
    ngoai_dm_info: Optional[dict] = None,
    has_curr_orders: bool = True,
):
    ws = wb.create_sheet(sheet_name)
    ws.views.sheetView[0].showGridLines = False

    n_slots = len(slots)
    _make_banner(
        ws, 11,
        title,
        f"Tháng {month:02d}/{year}  |  Tuần đang chọn: {slots[current_slot_idx - 1]['label']}",
    )

    widths = [38, 17, 17, 17, 17, 17, 17, 19, 19, 14, 14]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # ---- BẢNG 1: SẢN LƯỢNG BÁNH & NVL (Row 3..12) ----
    prod_headers = [
        "TUẦN / NGÀY", "TỔNG SL", "HAMBURGER", "HOTDOG", "PITA", "SANDWICH",
        "BÌNH QUÂN", "BÒ THƯỢNG HẠNG", "MIẾNG GÀ", "MIẾNG TÔM", "MIẾNG THỊT",
    ]
    for col_idx, h in enumerate(prod_headers, 1):
        _style_cell(ws.cell(3, col_idx, h), font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[3].height = 30

    prod_keys = ["hamburger", "hotdog", "pita", "sandwich", "avg", "bo_th", "mieng_ga", "mieng_tom", "mieng_thit"]
    # Ánh xạ từ cột bảng 1 sang cột detail sheet: H(8)..K(11), G(7), M(13)..P(16)
    detail_col_for_prod = ["H", "I", "J", "K", "G", "M", "N", "O", "P"]

    for i in range(6):
        r = 4 + i
        if i < n_slots:
            sl = slots[i]
            sidx = sl["slot_idx"]
            is_curr = (sidx == current_slot_idx)
            row_fill = FILL_GREEN if is_curr else (FILL_ZEBRA if r % 2 == 0 else FILL_WHITE)
            _style_cell(ws.cell(r, 1, sl["label"]), font=FONT_BOLD if is_curr else FONT_DATA, fill=row_fill)
            _style_cell(ws.cell(r, 2, f"=SUM(C{r}:F{r})"), font=FONT_BOLD, fill=row_fill, num_fmt=FMT_INT)

            if level == "supporter":
                if is_curr:
                    for j, d_col in enumerate(detail_col_for_prod):
                        fmt = FMT_DEC if prod_keys[j] == "avg" else FMT_INT
                        val_f = None if (prod_keys[j] == "avg" and not has_curr_orders) else f"='{detail_sheet_name}'!{d_col}4"
                        _style_cell(ws.cell(r, 3 + j, val_f),
                                    font=FONT_BOLD, fill=row_fill, num_fmt=fmt)
                else:
                    saved = {} if reset_month else (sup_dash_state or {}).get("weeks", {}).get(sidx, {})
                    for j, pk in enumerate(prod_keys):
                        val = saved.get(pk, 0.0)
                        if pk == "avg":
                            val = val if val > 0 else None
                        fmt = FMT_DEC if pk == "avg" else FMT_INT
                        _style_cell(ws.cell(r, 3 + j, val), font=FONT_DATA, fill=row_fill, num_fmt=fmt)
            else:
                # Region / Total: sum or averageIF(>0) từ các sheet con
                for j, pk in enumerate(prod_keys):
                    col_l = get_column_letter(3 + j)
                    if pk == "avg":
                        has_child_avg = any(wb[cs].cell(r, 7).value is not None for cs in child_sheets)
                        if not has_child_avg:
                            f_str = None
                        else:
                            refs = ",".join(f"'{cs}'!{col_l}{r}" for cs in child_sheets)
                            nonz = "+".join(f"IF('{cs}'!{col_l}{r}>0,1,0)" for cs in child_sheets)
                            f_str = f"=IFERROR(SUM({refs})/({nonz}),0)"
                        fmt = FMT_DEC
                    else:
                        f_str = "=" + "+".join(f"'{cs}'!{col_l}{r}" for cs in child_sheets)
                        fmt = FMT_INT
                    _style_cell(ws.cell(r, 3 + j, f_str), font=FONT_BOLD if is_curr else FONT_DATA, fill=row_fill, num_fmt=fmt)
        else:
            _style_cell(ws.cell(r, 1, ""), fill=FILL_GREY)
            for col_idx in range(2, 12):
                _style_cell(ws.cell(r, col_idx, None), fill=FILL_GREY)
        ws.row_dimensions[r].height = 20

    # Row 10: TỔNG THÁNG
    _style_cell(ws.cell(10, 1, "TỔNG THÁNG"), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL)
    for col_idx in range(2, 12):
        cl = get_column_letter(col_idx)
        if col_idx == 7:
            f_str = f'=IFERROR(AVERAGEIF({cl}4:{cl}9,">0"),0)'
            fmt = FMT_DEC
        else:
            f_str = f"=SUM({cl}4:{cl}9)"
            fmt = FMT_INT
        _style_cell(ws.cell(10, col_idx, f_str), font=FONT_BOLD, fill=FILL_ICE, border=BORDER_TOTAL, num_fmt=fmt)
    ws.row_dimensions[10].height = 22

    # Row 11: CÙNG KỲ NĂM TRƯỚC
    _style_cell(ws.cell(11, 1, f"THÁNG {month:02d}/{year - 1} (Cùng kỳ)"), font=FONT_BOLD, fill=FILL_WHITE)
    _style_cell(ws.cell(11, 2, "=SUM(C11:F11)"), font=FONT_BOLD, fill=FILL_WHITE, num_fmt=FMT_INT)
    if level == "supporter":
        py = (sup_dash_state or {}).get("prev_year_prod", {})
        for j, pk in enumerate(prod_keys):
            _style_cell(ws.cell(11, 3 + j, py.get(pk, 0.0)), font=FONT_INPUT, fill=FILL_WHITE,
                        num_fmt=FMT_DEC if pk == "avg" else FMT_INT)
    else:
        for j, pk in enumerate(prod_keys):
            cl = get_column_letter(3 + j)
            if pk == "avg":
                refs = ",".join(f"'{cs}'!{cl}11" for cs in child_sheets)
                nonz = "+".join(f"IF('{cs}'!{cl}11>0,1,0)" for cs in child_sheets)
                f_str = f"=IFERROR(SUM({refs})/({nonz}),0)"
            else:
                f_str = "=" + "+".join(f"'{cs}'!{cl}11" for cs in child_sheets)
            _style_cell(ws.cell(11, 3 + j, f_str), font=FONT_DATA, fill=FILL_WHITE,
                        num_fmt=FMT_DEC if pk == "avg" else FMT_INT)

    # Row 12: TỶ LỆ SO CÙNG KỲ
    _style_cell(ws.cell(12, 1, "TỶ LỆ SO CÙNG KỲ (%)"), font=FONT_SUBHDR, fill=FILL_ZEBRA)
    for col_idx in range(2, 12):
        cl = get_column_letter(col_idx)
        _style_cell(ws.cell(12, col_idx, f"=IFERROR({cl}10/{cl}11,0)"), font=FONT_BOLD, fill=FILL_ZEBRA, num_fmt=FMT_PCT)

    # ---- BẢNG 2: DOANH THU, MỤC TIÊU & THU HỒI CÔNG NỢ (Row 14..18) ----
    _style_cell(ws.cell(14, 1, "CHỈ TIÊU TÀI CHÍNH"), font=FONT_HDR, fill=FILL_SUBHDR,
                align=Alignment(horizontal="center", vertical="center"))
    for i in range(6):
        col_idx = 2 + i
        lbl = f"=A{4 + i}" if i < n_slots else ""
        _style_cell(ws.cell(14, col_idx, lbl), font=FONT_HDR, fill=FILL_SUBHDR,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    _style_cell(ws.cell(14, 8, "LŨY KẾ THÁNG"), font=FONT_HDR, fill=FILL_SUBHDR,
                align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    _style_cell(ws.cell(14, 9, "MỤC TIÊU / TỶ LỆ"), font=FONT_HDR, fill=FILL_SUBHDR,
                align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[14].height = 34

    labels_fin = [
        (15, "Doanh thu (VNĐ)"),
        (16, "% Đạt mục tiêu tháng"),
        (17, f"Doanh thu T{month:02d}/{year - 1} (VNĐ)"),
        (18, "Thu hồi công nợ (VNĐ)"),
    ]
    for r, txt in labels_fin:
        _style_cell(ws.cell(r, 1, txt), font=FONT_BOLD, fill=FILL_ZEBRA if r % 2 == 0 else FILL_WHITE)
        ws.row_dimensions[r].height = 21

    for i in range(6):
        col_idx = 2 + i
        cl = get_column_letter(col_idx)
        if i < n_slots:
            sidx = slots[i]["slot_idx"]
            is_curr = (sidx == current_slot_idx)
            # Row 15: Doanh thu tuần
            if level == "supporter":
                if is_curr:
                    _style_cell(ws.cell(15, col_idx, f"='{detail_sheet_name}'!Q4"), font=FONT_BOLD, fill=FILL_GREEN, num_fmt=FMT_INT)
                else:
                    saved_w = {} if reset_month else (sup_dash_state or {}).get("weeks", {}).get(sidx, {})
                    _style_cell(ws.cell(15, col_idx, saved_w.get("rev", 0.0)), font=FONT_DATA, num_fmt=FMT_INT)
            else:
                f_rev = "=" + "+".join(f"'{cs}'!{cl}15" for cs in child_sheets)
                _style_cell(ws.cell(15, col_idx, f_rev), font=FONT_BOLD if is_curr else FONT_DATA,
                            fill=FILL_GREEN if is_curr else FILL_WHITE, num_fmt=FMT_INT)

            # Row 16: % Đạt mục tiêu
            _style_cell(ws.cell(16, col_idx, f"=IFERROR({cl}15/$I$16,0)"), font=FONT_DATA, num_fmt=FMT_PCT)

            # Row 17: Cùng kỳ theo tuần (nếu có)
            _style_cell(ws.cell(17, col_idx, None), fill=FILL_GREY)

            # Row 18: Thu hồi công nợ (= tổng Doanh thu tuần trước của từng NQ từ LichSuBaoCao qua sheet chi tiết!R4)
            if level == "supporter":
                if is_curr:
                    _style_cell(ws.cell(18, col_idx, f"='{detail_sheet_name}'!R4"),
                                font=FONT_BOLD, fill=FILL_GREEN, num_fmt=FMT_INT)
                else:
                    saved_w = {} if reset_month else (sup_dash_state or {}).get("weeks", {}).get(sidx, {})
                    _style_cell(ws.cell(18, col_idx, saved_w.get("debt", 0.0)), font=FONT_DATA, num_fmt=FMT_INT)
            else:
                f_debt = "=" + "+".join(f"'{cs}'!{cl}18" for cs in child_sheets)
                _style_cell(ws.cell(18, col_idx, f_debt), font=FONT_BOLD if is_curr else FONT_DATA,
                            fill=FILL_GREEN if is_curr else FILL_WHITE, num_fmt=FMT_INT)
        else:
            for r in (15, 16, 17, 18):
                _style_cell(ws.cell(r, col_idx, None), fill=FILL_GREY)

    # Col H (Lũy kế) & Col I (Mục tiêu / Tỷ lệ)
    _style_cell(ws.cell(15, 8, "=SUM(B15:G15)"), font=FONT_BOLD, fill=FILL_ICE, num_fmt=FMT_INT)
    _style_cell(ws.cell(15, 9, "=IFERROR(H15/I16,0)"), font=FONT_BOLD, fill=FILL_ICE, num_fmt=FMT_PCT)

    _style_cell(ws.cell(16, 8, "=IFERROR(H15/I16,0)"), font=FONT_BOLD, fill=FILL_ZEBRA, num_fmt=FMT_PCT)
    if level == "supporter":
        t_rev = (sup_dash_state or {}).get("target_rev", 1_000_000_000)
        _style_cell(ws.cell(16, 9, t_rev), font=FONT_INPUT, fill=FILL_ZEBRA, num_fmt=FMT_INT)
        _style_cell(ws.cell(17, 8, (sup_dash_state or {}).get("prev_year_rev", 0.0)), font=FONT_INPUT, num_fmt=FMT_INT)
    else:
        _style_cell(ws.cell(16, 9, "=" + "+".join(f"'{cs}'!I16" for cs in child_sheets)), font=FONT_BOLD, fill=FILL_ZEBRA, num_fmt=FMT_INT)
        _style_cell(ws.cell(17, 8, "=" + "+".join(f"'{cs}'!H17" for cs in child_sheets)), font=FONT_BOLD, num_fmt=FMT_INT)
    _style_cell(ws.cell(17, 9, "=IFERROR(H15/H17,0)"), font=FONT_BOLD, num_fmt=FMT_PCT)
    _style_cell(ws.cell(18, 8, "=SUM(B18:G18)"), font=FONT_BOLD, fill=FILL_ICE, num_fmt=FMT_INT)
    _style_cell(ws.cell(18, 9, "=IFERROR(H18/H15,0)"), font=FONT_BOLD, fill=FILL_ICE, num_fmt=FMT_PCT)

    # ---- BẢNG 3: KPI CỬA HÀNG (LIÊN KẾT CHECKBOX Ds_nhuong_quyen) (Row 20..27) ----
    _style_cell(ws.cell(20, 1, "TRẠNG THÁI CỬA HÀNG (THEO Ds_nhuong_quyen)"), font=FONT_HDR, fill=FILL_NAVY,
                align=Alignment(horizontal="left", vertical="center", wrap_text=True))
    for mc in range(2, 6):
        _style_cell(ws.cell(20, mc, "GIÁ TRỊ" if mc == 2 else None), font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center"))
    ws.merge_cells("B20:E20")
    for mc in range(6, 10):
        _style_cell(ws.cell(20, mc, "GHI CHÚ" if mc == 6 else None), font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center"))
    ws.merge_cells("F20:I20")
    ws.row_dimensions[20].height = 28

    ds_end = ws.parent["Ds_nhuong_quyen"].max_row
    ds_ql = f"Ds_nhuong_quyen!$A$2:$A${ds_end}"
    ds_act = f"Ds_nhuong_quyen!$D$2:$D${ds_end}"
    ds_kv = f"Ds_nhuong_quyen!$G$2:$G${ds_end}"

    if level == "supporter":
        kpi_rows = [
            ("Tổng số cửa hàng", f'=COUNTIF({ds_ql},"{sup_key}")', FMT_INT, "Tính theo Ds_nhuong_quyen"),
            ("Số lượng NQ hoạt động", f'=COUNTIFS({ds_ql},"{sup_key}",{ds_act},TRUE)', FMT_INT, "Tự động cập nhật khi tick Checkbox"),
            ("Số lượng NQ đang off", f'=COUNTIFS({ds_ql},"{sup_key}",{ds_act},FALSE)', FMT_INT, "Tự động cập nhật khi bỏ tick Checkbox"),
            ("Số NQ có đơn tuần này", f"='{detail_sheet_name}'!U7", FMT_INT, "Theo dữ liệu tuần đang chọn"),
            ("NQ BQ cao nhất tuần", f"='{detail_sheet_name}'!U8", None, "Theo sheet chi tiết"),
            ("NQ BQ thấp nhất (>0)", f"='{detail_sheet_name}'!U9", None, "Theo sheet chi tiết"),
            ("BQ trung bình tuần (>0)", f"='{detail_sheet_name}'!U10", FMT_DEC, "Theo sheet chi tiết"),
        ]
    elif level == "region":
        kv_name = "KV HCM" if "HCM" in sheet_name else "KV Tỉnh"
        kpi_rows = [
            ("Tổng số cửa hàng", f'=COUNTIF({ds_kv},"{kv_name}")', FMT_INT, "Tính theo Ds_nhuong_quyen"),
            ("Số lượng NQ hoạt động", f'=COUNTIFS({ds_kv},"{kv_name}",{ds_act},TRUE)', FMT_INT, "Liên kết Checkbox Ds_nhuong_quyen"),
            ("Số lượng NQ đang off", f'=COUNTIFS({ds_kv},"{kv_name}",{ds_act},FALSE)', FMT_INT, "Liên kết Checkbox Ds_nhuong_quyen"),
            ("Số NQ có đơn tuần này", "=" + "+".join(f"'{cs}'!B24" for cs in child_sheets), FMT_INT, "Cộng từ các quản lý trong KV"),
            ("BQ trung bình tuần", f"=G{3 + current_slot_idx}", FMT_DEC, "Trung bình cộng các quản lý"),
        ]
    else:
        # Dashboard Tổng
        kpi_rows = [
            ("Tổng số cửa hàng (trong DM)", f'=COUNTIFS({ds_ql},"<>",{ds_ql},"<>Ngoài danh mục")', FMT_INT, "Tổng HCM + Tỉnh"),
            ("Số lượng NQ hoạt động", f'=COUNTIFS({ds_ql},"<>",{ds_ql},"<>Ngoài danh mục",{ds_act},TRUE)', FMT_INT, "Liên kết Checkbox Ds_nhuong_quyen"),
            ("Số lượng NQ đang off", f'=COUNTIFS({ds_ql},"<>",{ds_ql},"<>Ngoài danh mục",{ds_act},FALSE)', FMT_INT, "Liên kết Checkbox Ds_nhuong_quyen"),
            ("Số NQ Ngoài danh mục", f'=COUNTIF({ds_ql},"Ngoài danh mục")', FMT_INT, "Cần gán Quản lý ở Ds_nhuong_quyen"),
            ("Số NQ có đơn tuần này", "=" + "+".join(f"'{cs}'!B24" for cs in child_sheets), FMT_INT, "Trong danh mục 2 KV"),
            ("Doanh thu Ngoài danh mục (tuần)", ngoai_dm_info.get("rev", 0.0) if ngoai_dm_info else 0.0, FMT_INT,
             f"Mã: {', '.join(ngoai_dm_info.get('codes', []))}" if ngoai_dm_info and ngoai_dm_info.get("codes") else "Không phát sinh"),
            ("Tổng doanh thu toàn hệ thống (tuần)", f"=INDEX(B15:G15,1,{current_slot_idx})+B26", FMT_INT, "Khớp với DuLieuTho(tuan_nay)!F (đã +4%)"),
        ]

    for idx, (lbl, val, fmt, note) in enumerate(kpi_rows, start=21):
        fill = FILL_AMBER if ("Ngoài danh mục" in lbl and ngoai_dm_info and ngoai_dm_info.get("codes")) else (FILL_ZEBRA if idx % 2 == 0 else FILL_WHITE)
        _style_cell(ws.cell(idx, 1, lbl), font=FONT_BOLD, fill=fill,
                    align=Alignment(horizontal="left", vertical="center"))
        val_align = Alignment(horizontal="right" if fmt else "left", vertical="center", wrap_text=True)
        for mc in range(2, 6):
            _style_cell(ws.cell(idx, mc, val if mc == 2 else None), font=FONT_BOLD, fill=fill,
                        num_fmt=fmt if mc == 2 else None, align=val_align)
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=5)
        note_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for mc in range(6, 10):
            _style_cell(ws.cell(idx, mc, note if mc == 6 else None), font=FONT_MUTED, fill=fill, align=note_align)
        ws.merge_cells(start_row=idx, start_column=6, end_row=idx, end_column=9)
        ws.row_dimensions[idx].height = 22

    # ---- 4 BIỂU ĐỒ TỐI ƯU (Bố trí 2x2 ngay ngắn, không chồng lấn) ----
    last_slot_row = 3 + n_slots
    last_slot_col = 1 + n_slots

    # Chart 1: Tổng sản lượng bánh theo tuần
    c1 = BarChart()
    c1.add_data(Reference(ws, min_col=2, min_row=3, max_row=last_slot_row), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=4, max_row=last_slot_row))
    _style_bar_chart(
        c1,
        "Tổng sản lượng bánh theo tuần",
        y_title="Số phần bánh",
        show_labels=True,
        height=10.2,
        width=15.8,
        y_num_fmt='#,##0',
        label_num_fmt='#,##0;;""',
        show_legend=False,
    )
    ws.add_chart(c1, "A29")

    # Chart 2: Cơ cấu sản lượng theo loại bánh (Stacked)
    c2 = BarChart()
    c2.add_data(Reference(ws, min_col=3, max_col=6, min_row=3, max_row=last_slot_row), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=4, max_row=last_slot_row))
    _style_bar_chart(
        c2,
        "Cơ cấu sản lượng theo loại bánh",
        y_title="Số phần bánh",
        stacked=True,
        show_labels=False,
        height=10.2,
        width=15.8,
        y_num_fmt='#,##0',
        show_legend=True,
    )
    ws.add_chart(c2, "E29")

    # Chart 3: Xu hướng Bình quân phần/ngày
    c3 = LineChart()
    c3.add_data(Reference(ws, min_col=7, min_row=3, max_row=last_slot_row), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=1, min_row=4, max_row=last_slot_row))
    _style_line_chart(
        c3,
        "Xu hướng Bình quân phần/ngày theo tuần",
        y_title="Phần / ngày",
        height=10.2,
        width=15.8,
        label_num_fmt='0.0;;""',
        show_legend=False,
    )
    ws.add_chart(c3, "A50")

    # Chart 4: Doanh thu tuần vs Thu hồi công nợ (từ rows 15 & 18)
    c4 = BarChart()
    cats4 = Reference(ws, min_col=2, max_col=last_slot_col, min_row=14, max_row=14)
    c4.add_data(Reference(ws, min_col=1, max_col=last_slot_col, min_row=15, max_row=15), from_rows=True, titles_from_data=True)
    c4.add_data(Reference(ws, min_col=1, max_col=last_slot_col, min_row=18, max_row=18), from_rows=True, titles_from_data=True)
    c4.set_categories(cats4)
    _style_bar_chart(
        c4,
        "Doanh thu tuần vs Thu hồi công nợ (Triệu VNĐ)",
        y_title="Triệu VNĐ",
        show_labels=False,
        height=10.2,
        width=15.8,
        y_num_fmt='#,##0,,',
        show_legend=True,
    )
    ws.add_chart(c4, "E50")


def write_lich_su_bao_cao(
    wb: openpyxl.Workbook,
    existing_history: List[list],
    stores: OrderedDict,
    weekly_info: dict,
    slot: dict,
    year: int,
    month: int,
    prev_week_label: Optional[str],
) -> int:
    """Ghi sheet LichSuBaoCao(tong_hop): bảo toàn dữ liệu cũ và cập nhật/append tuần đang xử lý."""
    ws = wb.create_sheet("LichSuBaoCao(tong_hop)")
    widths = [8, 8, 9, 24, 12, 12, 12, 14, 16, 11, 36, 13, 10, 13, 14, 12, 11, 11, 14, 14, 12, 12, 12, 17, 19]
    for col_idx, (h, w) in enumerate(zip(HISTORY_HEADERS, widths), 1):
        _style_cell(ws.cell(1, col_idx, h), font=FONT_HDR, fill=FILL_NAVY,
                    align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[1].height = 32

    # Lập index doanh thu tuần trước theo mã NQ để tính Thu hồi công nợ (= Doanh thu tuần trước - 4%)
    prev_rev_by_code = {}
    if prev_week_label:
        for rw in existing_history:
            if len(rw) >= 24 and str(rw[3]) == prev_week_label:
                prev_rev_by_code[norm_code(rw[9])] = _to_float(rw[23])

    # Xóa các dòng trùng (Năm, Tháng, Nhãn tuần) nếu chạy lại cùng 1 tuần -> idempotent
    filtered_history = [
        rw for rw in existing_history
        if not (rw[0] == year and rw[1] == month and str(rw[3]) == slot["label"])
    ]

    piv = weekly_info["pivot"]
    revs = weekly_info["revenue"]
    days_map = weekly_info["days"]

    new_rows = []
    for ma, st in stores.items():
        p = piv.get(ma, {})
        has_order = ma in piv
        ngay_kd = (days_map.get(ma) or slot["days_count"]) if has_order else 0
        hb = float(p.get("1001", 0.0))
        hd = float(p.get("1004", 0.0))
        pt = float(p.get("1003", 0.0))
        sw = float(p.get("1002", 0.0)) / 2.0
        tot = hb + hd + pt + sw
        avg = tot / ngay_kd if ngay_kd > 0 else 0.0
        bo = float(p.get("1214", 0.0))
        ga = float(p.get("1200", 0.0))
        tom = float(p.get("1203", 0.0))
        thit = float(p.get("1201", 0.0))
        rv = float(round(revs.get(ma, 0.0) * 1.04))
        debt = float(round(prev_rev_by_code.get(ma, 0.0) * 0.96))
        kv_tong = get_kv_tong(st["ql"])

        new_rows.append([
            year, month, slot["week_num"], slot["label"],
            slot["start_date"].strftime("%Y-%m-%d"), slot["end_date"].strftime("%Y-%m-%d"),
            kv_tong, st["ql"], st["kv"], ma, st["ten"],
            "Có" if st["active"] else "Không",
            ngay_kd, tot, round(avg, 2),
            hb, hd, pt, sw, bo, ga, tom, thit, rv, debt,
        ])

    all_rows = filtered_history + new_rows
    for idx, rw in enumerate(all_rows, start=2):
        fill = FILL_ZEBRA if idx % 2 == 0 else FILL_WHITE
        for col_idx, val in enumerate(rw, 1):
            fmt = None
            if col_idx in (13, 14, 16, 17, 18, 20, 21, 22, 23, 24, 25):
                fmt = FMT_INT
            elif col_idx in (15, 19):
                fmt = FMT_DEC
            _style_cell(ws.cell(idx, col_idx, val), fill=fill, num_fmt=fmt)
        ws.row_dimensions[idx].height = 19

    ws.freeze_panes = "L2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HISTORY_HEADERS))}{1 + len(all_rows)}"
    return 1 + len(all_rows)


# --------------------------------------------------------------------------------------
# BƠM EXCEL FORM CONTROL CHECKBOX VÀO Ds_nhuong_quyen (CHUẨN OOXML EXCEL 16)
# --------------------------------------------------------------------------------------
def inject_form_checkboxes(xlsx_path: str, sheet_name: str, row_states: List[Tuple[int, bool]]):
    """
    Chèn các Excel Form Control CheckBox thật vào cột D của sheet `Ds_nhuong_quyen`,
    mỗi checkbox liên kết trực tiếp với ô `$D$<row>` (TRUE/FALSE).
    Cấu trúc XML bám sát 100% file chuẩn xuất từ Microsoft Excel 16.
    """
    if not row_states:
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            zin.extractall(tmpdir)

        # 1. Tìm sheet xml của `Ds_nhuong_quyen`
        import xml.etree.ElementTree as ET
        wb_xml = os.path.join(tmpdir, "xl", "workbook.xml")
        wb_rels_xml = os.path.join(tmpdir, "xl", "_rels", "workbook.xml.rels")
        wb_tree = ET.parse(wb_xml)
        rels_tree = ET.parse(wb_rels_xml)

        ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        target_rid = None
        for sh in wb_tree.getroot().findall(f".//{{{ns_main}}}sheet"):
            if sh.attrib.get("name") == sheet_name:
                target_rid = sh.attrib.get(f"{{{ns_r}}}id")
                break
        if not target_rid:
            return

        sheet_rel_target = None
        for rel in rels_tree.getroot():
            if rel.attrib.get("Id") == target_rid:
                sheet_rel_target = rel.attrib.get("Target").lstrip("/")
                if not sheet_rel_target.startswith("xl/"):
                    sheet_rel_target = "xl/" + sheet_rel_target
                break
        if not sheet_rel_target:
            return

        sheet_filename = os.path.basename(sheet_rel_target)
        sheet_xml_path = os.path.join(tmpdir, sheet_rel_target)
        sheet_rels_dir = os.path.join(os.path.dirname(sheet_xml_path), "_rels")
        os.makedirs(sheet_rels_dir, exist_ok=True)
        sheet_rels_path = os.path.join(sheet_rels_dir, f"{sheet_filename}.rels")

        # 2. Tạo xl/drawings/vmlDrawing_ds.vml & xl/ctrlProps/ctrlProp{i}.xml
        drawings_dir = os.path.join(tmpdir, "xl", "drawings")
        ctrl_dir = os.path.join(tmpdir, "xl", "ctrlProps")
        os.makedirs(drawings_dir, exist_ok=True)
        os.makedirs(ctrl_dir, exist_ok=True)

        vml_name = "vmlDrawing_ds.vml"
        vml_path = os.path.join(drawings_dir, vml_name)

        # Tính tọa độ cột D (index 0-based = 3): A(18) + B(12) + C(52) = ~82 chars; ở trong ô D (col 3)
        vml_shapes = []
        ctrl_items = []
        ctrl_rels = []

        for idx, (row_num, is_checked) in enumerate(row_states, start=1):
            shape_id = 1024 + idx
            row0 = row_num - 1  # 0-based
            top_pt = 21 + (row_num - 2) * 15
            checked_attr = ' checked="Checked"' if is_checked else ""
            ctrl_filename = f"ctrlProp_ds_{idx}.xml"
            ctrl_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
                f'objectType="CheckBox"{checked_attr} fmlaLink="$D${row_num}" lockText="1" noThreeD="1"/>'
            )
            with open(os.path.join(ctrl_dir, ctrl_filename), "w", encoding="utf-8") as cf:
                cf.write(ctrl_xml)

            checked_vml = "<x:Checked>1</x:Checked>" if is_checked else ""
            # Anchor: LeftCol=3, LeftOff=45px, TopRow=row0, TopOff=2px, RightCol=3, RightOff=75px, BotRow=row0, BotOff=18px
            vml_shapes.append(
                f'<v:shape id="_x0000_s{shape_id}" type="#_x0000_t201" '
                f"style='position:absolute;margin-left:440pt;margin-top:{top_pt}pt;width:16pt;height:14pt;z-index:{idx};mso-wrap-style:tight' "
                f'filled="f" fillcolor="windowText [64]" stroked="f" strokecolor="window [65]" strokeweight="3e-5mm" o:insetmode="auto">\n'
                f'  <v:fill color2="window [65]"/>\n'
                f'  <v:path shadowok="t" strokeok="t" fillok="t"/>\n'
                f'  <o:lock v:ext="edit" rotation="t"/>\n'
                f'  <v:textbox style=\'mso-direction-alt:auto\' o:singleclick="f"><div style=\'text-align:left\'></div></v:textbox>\n'
                f'  <x:ClientData ObjectType="Checkbox">\n'
                f'   <x:SizeWithCells/>\n'
                f'   <x:Anchor>3, 42, {row0}, 2, 3, 62, {row0}, 18</x:Anchor>\n'
                f'   <x:AutoFill>False</x:AutoFill>\n'
                f'   <x:AutoLine>False</x:AutoLine>\n'
                f'   <x:TextVAlign>Center</x:TextVAlign>\n'
                f'   {checked_vml}\n'
                f'   <x:FmlaLink>$D${row_num}</x:FmlaLink>\n'
                f'   <x:NoThreeD/>\n'
                f'  </x:ClientData>\n'
                f'</v:shape>'
            )
            ctrl_rels.append((idx, shape_id, row0, ctrl_filename))

        vml_doc = (
            '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel">\n'
            ' <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>\n'
            ' <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">\n'
            '  <v:stroke joinstyle="miter"/>\n'
            '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\n'
            '  <o:lock v:ext="edit" shapetype="t"/>\n'
            ' </v:shapetype>\n'
            + "\n".join(vml_shapes)
            + "\n</xml>"
        )
        with open(vml_path, "w", encoding="utf-8") as vf:
            vf.write(vml_doc)

        # 3. Cập nhật sheet rels
        rel_lines = []
        vml_rid = "rId1"
        rel_lines.append(
            f'<Relationship Id="{vml_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/{vml_name}"/>'
        )
        for idx, shape_id, row0, ctrl_filename in ctrl_rels:
            rid = f"rId{idx + 1}"
            rel_lines.append(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/{ctrl_filename}"/>'
            )
            ctrl_items.append(
                f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                f'<mc:Choice Requires="x14">'
                f'<control shapeId="{shape_id}" r:id="{rid}" name="Check Box {idx}">'
                f'<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0">'
                f'<anchor moveWithCells="1">'
                f'<from><xdr:col>3</xdr:col><xdr:colOff>400050</xdr:colOff><xdr:row>{row0}</xdr:row><xdr:rowOff>19050</xdr:rowOff></from>'
                f'<to><xdr:col>3</xdr:col><xdr:colOff>590550</xdr:colOff><xdr:row>{row0}</xdr:row><xdr:rowOff>171450</xdr:rowOff></to>'
                f'</anchor></controlPr></control></mc:Choice></mc:AlternateContent>'
            )

        rels_doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            + "\n".join(rel_lines)
            + "\n</Relationships>"
        )
        with open(sheet_rels_path, "w", encoding="utf-8") as rf:
            rf.write(rels_doc)

        # 4. Chèn namespaces + <legacyDrawing> + <controls> vào sheet XML
        with open(sheet_xml_path, "r", encoding="utf-8") as sf:
            s_txt = sf.read()

        # Đảm bảo khai báo namespaces trên <worksheet ...>
        extra_ns = ""
        if "xmlns:r=" not in s_txt:
            extra_ns += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        if "xmlns:xdr=" not in s_txt:
            extra_ns += ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        if "xmlns:x14=" not in s_txt:
            extra_ns += ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
        if "xmlns:mc=" not in s_txt:
            extra_ns += ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14"'
        if extra_ns:
            s_txt = s_txt.replace("<worksheet ", f"<worksheet {extra_ns.strip()} ", 1)

        controls_block = (
            f'<legacyDrawing r:id="{vml_rid}"/>'
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<mc:Choice Requires="x14"><controls>'
            + "".join(ctrl_items)
            + '</controls></mc:Choice></mc:AlternateContent>'
        )
        s_txt = s_txt.replace("</worksheet>", controls_block + "</worksheet>")
        with open(sheet_xml_path, "w", encoding="utf-8") as sf:
            sf.write(s_txt)

        # 5. Cập nhật [Content_Types].xml cho vml và ctrlProp*
        ct_path = os.path.join(tmpdir, "[Content_Types].xml")
        with open(ct_path, "r", encoding="utf-8") as cf:
            ct_txt = cf.read()

        add_ct = []
        if 'Extension="vml"' not in ct_txt:
            add_ct.append('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')
        for idx, *_ in ctrl_rels:
            part = f"/xl/ctrlProps/ctrlProp_ds_{idx}.xml"
            if part not in ct_txt:
                add_ct.append(f'<Override PartName="{part}" ContentType="application/vnd.ms-excel.controlproperties+xml"/>')
        if add_ct:
            ct_txt = ct_txt.replace("</Types>", "".join(add_ct) + "</Types>")
            with open(ct_path, "w", encoding="utf-8") as cf:
                cf.write(ct_txt)

        # 5b. Đảm bảo <numFmt> trong biểu đồ có sourceLinked="0" để Excel áp dụng định dạng nhãn dữ liệu (ẩn số 0, hiển thị "Tr")
        charts_dir = os.path.join(tmpdir, "xl", "charts")
        if os.path.isdir(charts_dir):
            for cfn in os.listdir(charts_dir):
                if cfn.startswith("chart") and cfn.endswith(".xml"):
                    cp = os.path.join(charts_dir, cfn)
                    with open(cp, "r", encoding="utf-8") as cf:
                        ctxt = cf.read()
                    new_ctxt = re.sub(
                        r'<((?:\w+:)?numFmt)\b(?![^>]*sourceLinked=)',
                        r'<\1 sourceLinked="0"',
                        ctxt,
                    )
                    if new_ctxt != ctxt:
                        with open(cp, "w", encoding="utf-8") as cf:
                            cf.write(new_ctxt)

        # 6. Đóng gói lại file .xlsx
        tmp_out = xlsx_path + ".tmp"
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(tmpdir):
                for fn in files:
                    full = os.path.join(root_dir, fn)
                    arc = os.path.relpath(full, tmpdir).replace("\\", "/")
                    zout.write(full, arc)
        os.replace(tmp_out, xlsx_path)


# --------------------------------------------------------------------------------------
# PIPELINE CHÍNH
# --------------------------------------------------------------------------------------
def find_default_base_report() -> str:
    """Tìm file báo cáo tháng gốc gần nhất trong output/, root, templates/, hoặc Downloads."""
    candidates = []
    for folder in (DEFAULT_OUT_DIR, ROOT_DIR, os.path.join(os.path.expanduser("~"), "Downloads")):
        if os.path.isdir(folder):
            for fn in os.listdir(folder):
                if fn.startswith("BC_TONG") and fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    full = os.path.join(folder, fn)
                    candidates.append((os.path.getmtime(full), full))
    if candidates:
        candidates.sort(reverse=True)
        # Ưu tiên BC_TONG NQ_Tuan45 9-15.xlsx nếu có, hoặc file mới nhất trong output
        for _, p in candidates:
            if os.path.dirname(p) == DEFAULT_OUT_DIR:
                return p
        for _, p in candidates:
            if os.path.basename(p) == "BC_TONG NQ_Tuan45 9-15.xlsx":
                return p
        return candidates[0][1]
    if os.path.exists(DEFAULT_TEMPLATE):
        return DEFAULT_TEMPLATE
    raise FileNotFoundError("Không tìm thấy file báo cáo tháng gốc (BC_TONG_base.xlsx / BC_TONG NQ_*.xlsx).")


def generate_monthly_report(
    weekly_path: str,
    base_path: Optional[str] = None,
    year: Optional[int] = None,
    month: Optional[int] = None,
    slot_idx: Optional[int] = None,
    operating_days_override: Optional[int] = None,
    out_path: Optional[str] = None,
) -> dict:
    """Chạy toàn bộ quy trình tổng hợp báo cáo tháng và trả về metadata + đường dẫn output."""
    if not base_path:
        base_path = find_default_base_report()

    parsed = parse_date_range_from_filename(weekly_path)
    if year is None:
        year = parsed["year"] if parsed else dt.date.today().year
    if month is None:
        month = parsed["month"] if parsed else dt.date.today().month
    slots = get_month_week_slots(year, month)
    if slot_idx is None:
        slot_idx = parsed["slot_idx"] if parsed else 1
    slot_idx = max(1, min(slot_idx, len(slots)))
    slot = slots[slot_idx - 1]
    if operating_days_override is not None and operating_days_override > 0:
        slot = dict(slot)
        slot["days_count"] = int(operating_days_override)

    # Thu thập từ vựng từ file gốc để khôi phục tên tiếng Việt nếu file nguồn bị mất dấu '?'
    harvest_texts: List[str] = []
    if os.path.exists(base_path):
        wb_h = openpyxl.load_workbook(base_path, read_only=True, data_only=True)
        for sname in ("Ds_nhuong_quyen", "ListNQ", "6. Duyên", "8. Quỳnh", "10.CƯỜNG", "12.THẾ", "14.TUYẾT"):
            if sname in wb_h.sheetnames:
                for row in wb_h[sname].iter_rows(max_row=200, values_only=True):
                    harvest_texts.extend(v for v in row if isinstance(v, str))
        wb_h.close()

    df, weekly_info = load_weekly_inputs(weekly_path, harvest_texts, target_year=year, target_month=month)
    base_state = load_base_workbook_state(base_path, year, month, weekly_info)
    stores = base_state["stores"]
    _apply_group_value_split(stores, weekly_info)
    dash_state = base_state["dash_state"]
    detected_ym = base_state["detected_ym"]
    existing_history = base_state["history_rows"]

    # Quy tắc reset đầu tháng: nếu tuần 1 của tháng HOẶC khác tháng/năm với báo cáo gốc -> xóa sạch dữ liệu tuần cũ
    reset_month = (slot_idx == 1) or (detected_ym is not None and detected_ym != (year, month))

    # Nhãn tuần trước (trong cùng tháng hoặc cuối tháng trước trong LichSuBaoCao)
    prev_week_label: Optional[str] = None
    prev_month_rev_by_sup = defaultdict(float)
    if slot_idx > 1:
        prev_week_label = slots[slot_idx - 2]["label"]
    elif existing_history:
        # Tìm tuần gần nhất trong LichSuBaoCao trước (year, month)
        prior = sorted({(int(r[0]), int(r[1]), int(r[2]), str(r[3])) for r in existing_history
                        if (int(r[0]), int(r[1])) < (year, month)})
        if prior:
            prev_week_label = prior[-1][3]
            for r in existing_history:
                if str(r[3]) == prev_week_label:
                    prev_month_rev_by_sup[norm_supporter(r[7])] += _to_float(r[23])

    # Xây dựng workbook mới sạch 16 sheet
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. Ghi Ds_nhuong_quyen, DuLieuTho(tuan_nay), LichSuBaoCao(tong_hop) trước (hoặc sắp xếp lại thứ tự tab sau)
    ds_last_row = write_ds_nhuong_quyen(wb, stores)
    raw_last_data_row, raw_col_map = write_du_lieu_tho(wb, weekly_info, stores)
    hist_last_row = write_lich_su_bao_cao(wb, existing_history, stores, weekly_info, slot, year, month, prev_week_label)

    # 2. Ghi 5 sheet chi tiết quản lý
    prev_banh_map = {}
    if prev_week_label:
        for rw in existing_history:
            if len(rw) >= 14 and str(rw[3]) == prev_week_label:
                prev_banh_map[norm_code(rw[9])] = _to_float(rw[13])

    for sup in SUPPORTERS:
        write_supporter_detail_sheet(
            wb=wb,
            sup=sup,
            stores=stores,
            weekly_info=weekly_info,
            raw_col_map=raw_col_map,
            slot=slot,
            year=year,
            month=month,
            prev_week_label=prev_week_label,
            legacy_days=base_state["legacy_days"],
            ds_last_row=ds_last_row,
            raw_last_data_row=raw_last_data_row,
            hist_last_row=hist_last_row,
            prev_banh_map=prev_banh_map,
        )

    # 3. Ghi 5 Dashboard quản lý
    for sup in SUPPORTERS:
        sup_has_orders = any(ma in weekly_info["pivot"] for ma, st in stores.items() if st["ql"] == sup["key"])
        _write_dashboard_sheet(
            wb=wb,
            sheet_name=sup["dash"],
            title=f"DASHBOARD DOANH SỐ - {sup['key'].upper()} ({sup['kv']})",
            slots=slots,
            current_slot_idx=slot_idx,
            year=year,
            month=month,
            level="supporter",
            sup_key=sup["key"],
            sup_dash_state=dash_state.get(sup["key"]),
            reset_month=reset_month,
            detail_sheet_name=sup["detail"],
            prev_month_last_rev=prev_month_rev_by_sup.get(sup["key"], 0.0),
            has_curr_orders=sup_has_orders,
        )

    # 4. Ghi Dashboard KV HCM & KV Tỉnh
    hcm_dashes = [s["dash"] for s in SUPPORTERS if s["kv"] == "KV HCM"]
    tinh_dashes = [s["dash"] for s in SUPPORTERS if s["kv"] == "KV Tỉnh"]
    _write_dashboard_sheet(
        wb=wb,
        sheet_name="2.Dashboard KV HCM",
        title="DASHBOARD KHU VỰC TP.HCM",
        slots=slots,
        current_slot_idx=slot_idx,
        year=year,
        month=month,
        level="region",
        child_sheets=hcm_dashes,
        reset_month=reset_month,
    )
    _write_dashboard_sheet(
        wb=wb,
        sheet_name="3.Dashboard KV Tỉnh",
        title="DASHBOARD KHU VỰC TỈNH",
        slots=slots,
        current_slot_idx=slot_idx,
        year=year,
        month=month,
        level="region",
        child_sheets=tinh_dashes,
        reset_month=reset_month,
    )

    # 5. Ghi Dashboard Tổng
    ngoai_codes = [ma for ma, st in stores.items() if st["ql"] == NGOAI_DM and ma in weekly_info["pivot"]]
    ngoai_rev = sum(round(weekly_info["revenue"].get(ma, 0.0) * 1.04) for ma in ngoai_codes)
    _write_dashboard_sheet(
        wb=wb,
        sheet_name="1.Dashboard Tổng",
        title="BÁO CÁO TỔNG HỢP NHƯỢNG QUYỀN TOÀN QUỐC",
        slots=slots,
        current_slot_idx=slot_idx,
        year=year,
        month=month,
        level="total",
        child_sheets=["2.Dashboard KV HCM", "3.Dashboard KV Tỉnh"],
        reset_month=reset_month,
        ngoai_dm_info={"codes": ngoai_codes, "rev": ngoai_rev},
    )

    # Sắp xếp đúng thứ tự 16 sheet chuẩn
    desired_order = [
        "1.Dashboard Tổng", "2.Dashboard KV HCM", "3.Dashboard KV Tỉnh",
        "5. KV Duyên", "6. Duyên",
        "7.KV Quỳnh", "8. Quỳnh",
        "9.KVCường", "10.CƯỜNG",
        "11.KVThế", "12.THẾ",
        "13.KVTuyết", "14.TUYẾT",
        "Ds_nhuong_quyen", "DuLieuTho(tuan_nay)", "LichSuBaoCao(tong_hop)",
    ]
    wb._sheets = [wb[name] for name in desired_order]

    # Lưu file + chèn Form Control CheckBox thật cho Ds_nhuong_quyen
    out_file = out_path or os.path.join(DEFAULT_OUT_DIR, build_output_filename(slot, year, month))
    out_file = knh._safe_save_path(out_file)
    os.makedirs(os.path.dirname(os.path.abspath(out_file)), exist_ok=True)
    wb.save(out_file)

    row_states = [(idx, bool(st["active"])) for idx, st in enumerate(stores.values(), start=2)]
    inject_form_checkboxes(out_file, "Ds_nhuong_quyen", row_states)

    summary = {
        "out_file": out_file,
        "base_path": base_path,
        "year": year,
        "month": month,
        "slot_idx": slot_idx,
        "slot_label": slot["label"],
        "reset_month": reset_month,
        "rows_read": len(df),
        "stores_with_orders": len(weekly_info["pivot"]),
        "total_stores": len(stores),
        "new_stores": [ma for ma, st in stores.items() if st.get("is_new")],
        "total_revenue_raw": sum(weekly_info["revenue"].values()),
        "total_revenue": sum(round(v * 1.04) for v in weekly_info["revenue"].values()),
    }
    print(f"Tuần: {slot['label']} (Tháng {month:02d}/{year}) | Reset tháng mới: {'Có' if reset_month else 'Không'}")
    print(
        f"Đọc {len(df)} dòng hàng từ {len(weekly_info['pivot'])} cửa hàng | "
        f"Tổng tiền gốc: {summary['total_revenue_raw']:,.0f} VNĐ | "
        f"Doanh thu (+4%): {summary['total_revenue']:,.0f} VNĐ"
    )
    if summary["new_stores"]:
        print(f"Đã thêm {len(summary['new_stores'])} mã NQ mới vào Ds_nhuong_quyen: {', '.join(summary['new_stores'])}")
    print("Đã xuất:", out_file)
    return summary


def main():
    ap = argparse.ArgumentParser(description="Tổng hợp báo cáo doanh số Nhượng Quyền theo tuần / tháng")
    ap.add_argument("report", help="File báo cáo tuần (.xls, .xlsx hoặc .zip)")
    ap.add_argument("--base", default=None, help="File báo cáo tháng gốc để cộng dồn")
    ap.add_argument("--year", type=int, default=None, help="Năm báo cáo")
    ap.add_argument("--month", type=int, default=None, help="Tháng báo cáo (1-12)")
    ap.add_argument("--week-idx", type=int, default=None, help="Thứ tự tuần trong tháng (1..6)")
    ap.add_argument("--days", type=int, default=None, help="Số ngày kinh doanh mặc định trong tuần")
    ap.add_argument("--out", default=None, help="Đường dẫn file Excel xuất ra")
    a = ap.parse_args()

    generate_monthly_report(
        weekly_path=a.report,
        base_path=a.base,
        year=a.year,
        month=a.month,
        slot_idx=a.week_idx,
        operating_days_override=a.days,
        out_path=a.out,
    )


if __name__ == "__main__":
    main()
