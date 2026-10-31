# -*- coding: utf-8 -*-
"""
KIỂM TRA NHẬP HÀNG NHƯỢNG QUYỀN THEO TUẦN
==========================================
Thay thế quy trình thủ công:  báo cáo thô (.xls) -> Pivot Table -> CHECK_2026.xlsx
(ChiTiet / SOP_Banh / SOP_Nuoc / QuyDoi_DonVi / KIỂM TRA).

Cách dùng:
    python workflow/kiem_tra_nhap_hang.py "19 - 25.10.xls"
    python workflow/kiem_tra_nhap_hang.py "19 - 25.10.xls" --per-nq
    python workflow/kiem_tra_nhap_hang.py --validate        # đối chiếu với kết quả có sẵn trong CHECK_2026.xlsx

Quy ước % (giống file CHECK_2026):  % = thực nhập * 100 / định mức SOP - 100
    Âm = nhập thiếu, dương = nhập dư.
Ngưỡng cảnh báo từng nguyên liệu lấy đúng theo Conditional Formatting của sheet KIỂM TRA.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from viet_restore import Restorer, looks_like_tcvn3  # noqa: E402
from collections import OrderedDict, defaultdict

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# --------------------------------------------------------------------------------------
# CẤU HÌNH
# --------------------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
    HERE = os.path.join(ROOT, "workflow")
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
    ROOT = os.path.dirname(HERE)
DEFAULT_MASTER = os.path.join(ROOT, "CHECK_2026.xlsx")
DEFAULT_OUT_DIR = os.path.join(ROOT, "output")

# Ngưỡng cảnh báo thiếu (%), lấy từ Conditional Formatting sheet "KIỂM TRA" của CHECK_2026.xlsx
TH_NUOC_LY = -30  # Tổng ly nước quy đổi từ NVL so với ly theo SOP (sheet SOP_Nuoc, cột "Chênh lệch")

# Mã hàng trong báo cáo thô (cột "Mã HH-VT")
BANH = dict(hamburger="1001", hotdog="1004", pita="1003", sandwich="1002")
MA = dict(
    thit_hamburger="1201", thit_ga="1200", thit_tom="1203", thit_bo="1214",
    xuc_xich="1205", pate="1234", cha_bong="1204", jambon="1208",
    sua_bot="1403", tra_den="1405", cacao="1401", pepsi="1209", bot_chanh="1421", duong="1402",
    ca="1102", mayo="1101", dau_an="1104",
    giay_goi="2008", hop_burger="2017", hop_pita="2021", hop_sandwich="2064",
    tui_pita="2019", tui_hotdog="2027", ly="2068", mang_ep="2012", ong_hut="2066", tui_dung_ly="2067",
)

# Hệ số quy đổi (sheet QuyDoi_DonVi / SOP_Nuoc)
PATE_G_PER_UNIT = 500          # Pate 0.5kg/bịch -> 500g
GRAM_PER_KG = 1000
JAMBON_PIECE_PER_KG = 1000 / 250 * 16   # 1 bịch 250g làm 16 bánh
TUI_DUNG_LY_PER_KG = 192
MANG_EP_LY_PER_ROLL = 3600
ONG_HUT_PER_BAO = 250
LY_TRA_SUA_PER_KG_SUA = (1 / 1.1) * 12 * 6   # 1.1kg sữa bột -> 12 lít trà sữa -> 6 ly/lít
LY_CACAO_PER_KG = (1 / 1.5) * 6 * 8
LY_PEPSI_PER_CHAI = 9
LY_CHANH_PER_KG = 1 * 10 * 7
TRA_DEN_KG_PER_KG_SUA = 200 / 1100            # 200g trà đen / 1.1kg sữa bột
DUONG_KG_PER_KG_SUA = 1100 / 1100             # 1.1kg đường / 1.1kg sữa bột


# --------------------------------------------------------------------------------------
# ĐỌC DỮ LIỆU
# --------------------------------------------------------------------------------------
def load_weekly_report(path):
    """Đọc báo cáo thô .xls -> (DataFrame chi tiết, pivot {mã KH: {mã HH: tổng SL}}).
    Bỏ các dòng 'Total'. Tên cột đặt theo vị trí vì file .xls hay lỗi mã hóa tiếng Việt."""
    df = pd.read_excel(path, header=0)
    if df.shape[1] < 11:
        raise ValueError("File báo cáo tuần phải có >= 11 cột (Ngày CT, Số CT, Mã KH, Tên KH, Địa chỉ, "
                         "Mã HH, Tên HH, ĐVT, Số lượng, Giá, Số tiền)")
    df = df.iloc[:, :11]
    df.columns = ["ngay", "so_ct", "ma_kh", "ten_kh", "dia_chi", "ma_hh", "ten_hh", "dvt", "sl", "gia", "tien"]
    df = df[df["ma_kh"].notna() & df["ma_hh"].notna()].copy()
    df["ma_kh"] = df["ma_kh"].astype(str).str.strip()
    df["ma_hh"] = df["ma_hh"].apply(lambda v: str(int(v)) if isinstance(v, (int, float)) else str(v).strip())
    df["sl"] = pd.to_numeric(df["sl"], errors="coerce").fillna(0)
    df["ngay"] = pd.to_datetime(df["ngay"], errors="coerce")
    pv = df.groupby(["ma_kh", "ma_hh"])["sl"].sum()
    pivot = defaultdict(dict)
    for (kh, hh), v in pv.items():
        pivot[kh][hh] = float(v)
    return df, pivot


def load_master(path):
    """Danh mục nhượng quyền (Quản lý, Mã NQ, Tên NQ) từ sheet 'KIỂM TRA' của CHECK_2026.xlsx."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["KIỂM TRA"]
    master = OrderedDict()
    for row in ws.iter_rows(min_row=3, max_col=3, values_only=True):
        ql, ma, ten = row
        if ma:
            master[str(ma).strip()] = (ql or "", ten or "")
    wb.close()
    return master


NGOAI_DM = "Ngoài danh mục"


def fix_vietnamese(df, master_path, master):
    """Sửa chữ Việt bị mất dấu ('?') trong báo cáo thô (xem viet_restore.py). Giữ tên gốc ở cột *_goc."""
    texts = list(df["ten_kh"].dropna().unique())[:5000]
    if looks_like_tcvn3(texts):
        print("CẢNH BÁO: dữ liệu có vẻ là TCVN3 - cần bảng chuyển mã riêng (chưa hỗ trợ).")
    harvest = [t for _, t in master.values()]
    wb = openpyxl.load_workbook(master_path, read_only=True, data_only=True)
    for name in ("Dữ liệu thô", "ChiTiet", "SOP_Banh", "SOP_Nuoc", "QuyDoi_DonVi", "Tổng hợp T5"):
        if name in wb.sheetnames:
            for row in wb[name].iter_rows(max_row=8, values_only=True):
                harvest += [v for v in row if isinstance(v, str)]
    wb.close()
    rs = Restorer(harvest)
    df["ten_kh_goc"] = df["ten_kh"]
    df["ten_kh"] = df["ten_kh"].map(rs.restore)
    df["ten_hh_goc"] = df["ten_hh"]
    df["ten_hh"] = df["ten_hh"].map(rs.restore)
    return df




def best_name(series):
    """Tên ít ký tự '?' nhất; nếu bằng nhau thì tên xuất hiện nhiều nhất."""
    vals = [str(v).strip() for v in series if isinstance(v, str)]
    if not vals:
        return ""
    cnt = pd.Series(vals).value_counts()
    return sorted(cnt.index, key=lambda n: (n.count("?"), -cnt[n]))[0]


def extend_master(master, df):
    """Thêm các KH có trong báo cáo tuần nhưng chưa có trong danh mục (Quản lý = 'Ngoài danh mục')."""
    full = OrderedDict(master)
    for ma in sorted(df["ma_kh"].unique()):
        if ma not in full:
            full[ma] = (NGOAI_DM, best_name(df.loc[df["ma_kh"] == ma, "ten_kh"]))
    return full


# --------------------------------------------------------------------------------------
# TÍNH TOÁN SOP VÀ SO SÁNH
# --------------------------------------------------------------------------------------
# (khóa, nhóm, tên hiển thị, đơn vị so sánh, ngưỡng thiếu %)
BANH_ITEMS = [
    ("thit", "Bánh", "Thịt / Gà / Tôm / Bò hỗn hợp", "miếng", -10),
    ("ca", "Bánh", "Tương cà", "g", -20),
    ("mayo", "Bánh", "Mayonaise", "g", -20),
    ("dau", "Bánh", "Dầu ăn", "lít", -20),
    ("hop", "Bánh", "Hộp giấy Pita / Hamburger", "cái", -10),
    ("tui_hd", "Bánh", "Túi giấy Hotdog", "cái", -10),
    ("hop_sw", "Bánh", "Hộp giấy Sandwich", "cái", -10),
    ("pate", "Bánh", "Pate", "g", -20),
    ("jam", "Bánh", "Jambon", "miếng", -20),
    ("chabong", "Bánh", "Chà bông", "g", -20),
    ("xucxich", "Bánh", "Xúc xích", "cây", -10),
    ("bao", "Nước", "Túi dựng ly (bao)", "cái", -20),
    ("ly", "Nước", "Ly giấy", "cái", -30),
    ("mang", "Nước", "Màng ép ly", "ly", -30),
    ("ong", "Nước", "Ống hút", "cây", -30),
]
ITEM_META = {k: (g, n, u, th) for k, g, n, u, th in BANH_ITEMS}


def compute_nq(pivot_row):
    """Trả về dict kết quả cho 1 nhượng quyền từ {mã HH: SL}."""
    q = lambda code: float(pivot_row.get(code, 0) or 0)
    E, F, G, H = q(BANH["hamburger"]), q(BANH["hotdog"]), q(BANH["pita"]), q(BANH["sandwich"])

    # ---- SOP theo số lượng bánh nhập (sheet SOP_Banh) ----
    ly_sop = E + G + F + H / 2
    sop = {
        "thit": E + G,
        "ca": (E + F + G) * 5,
        "mayo": (E + G) * 5 + H / 2 * 5,
        "dau": (q(MA["thit_ga"]) + q(MA["thit_tom"])) / 80,
        "hop": E + G,
        "tui_hd": F,
        "hop_sw": H / 2,
        "pate": H / 2 * 5,
        "jam": H / 2,  # = (H/2/16 bịch) * 16
        "chabong": H / 2 * 5,
        "xucxich": F,
        "bao": ly_sop,
        "ly": ly_sop,
        "mang": ly_sop,
        "ong": ly_sop,
    }
    # ---- Thực nhập đã quy đổi (sheet QuyDoi_DonVi + ChiTiet) ----
    act = {
        "thit": q(MA["thit_hamburger"]) + q(MA["thit_ga"]) + q(MA["thit_tom"]) + q(MA["thit_bo"]),
        "ca": q(MA["ca"]) * GRAM_PER_KG,
        "mayo": q(MA["mayo"]) * GRAM_PER_KG,
        "dau": q(MA["dau_an"]),
        "hop": q(MA["hop_burger"]) + q(MA["hop_pita"]),
        "tui_hd": q(MA["tui_hotdog"]),
        "hop_sw": q(MA["hop_sandwich"]),
        "pate": q(MA["pate"]) * PATE_G_PER_UNIT,
        "jam": q(MA["jambon"]) * JAMBON_PIECE_PER_KG,
        "chabong": q(MA["cha_bong"]) * GRAM_PER_KG,
        "xucxich": q(MA["xuc_xich"]),
        "bao": q(MA["tui_dung_ly"]) * TUI_DUNG_LY_PER_KG,
        "ly": q(MA["ly"]),
        "mang": q(MA["mang_ep"]) * MANG_EP_LY_PER_ROLL,
        "ong": q(MA["ong_hut"]) * ONG_HUT_PER_BAO,
    }
    pct = {}
    for k in sop:
        pct[k] = (act[k] * 100 / sop[k] - 100) if sop[k] > 0 else None

    # ---- SOP nước (sheet SOP_Nuoc) ----
    sua, tra, cacao, pepsi, chanh, duong = (q(MA[x]) for x in ("sua_bot", "tra_den", "cacao", "pepsi", "bot_chanh", "duong"))
    ly_nvl = {
        "Trà sữa": sua * LY_TRA_SUA_PER_KG_SUA,
        "Cacao": cacao * LY_CACAO_PER_KG,
        "Pepsi": pepsi * LY_PEPSI_PER_CHAI,
        "Hồng trà chanh": chanh * LY_CHANH_PER_KG,
    }
    ly_nvl_tong = sum(ly_nvl.values())
    nuoc = {
        "sua": sua, "tra": tra, "duong": duong, "cacao": cacao, "pepsi": pepsi, "chanh": chanh,
        "ly_nvl": ly_nvl, "ly_nvl_tong": ly_nvl_tong, "ly_sop": ly_sop,
        "pct_ly": (ly_nvl_tong * 100 / ly_sop - 100) if ly_sop > 0 else None,
        "tra_sop": sua * TRA_DEN_KG_PER_KG_SUA,
        "duong_sop": sua * DUONG_KG_PER_KG_SUA,
        "pct_tra": (tra * 100 / (sua * TRA_DEN_KG_PER_KG_SUA) - 100) if sua > 0 else 0,
    }
    return dict(banh=dict(hamburger=E, hotdog=F, pita=G, sandwich=H), sop=sop, act=act, pct=pct, nuoc=nuoc)


def build_findings(res):
    """Danh sách cảnh báo [(nhóm, nguyên liệu, mô tả, % , mức)] cho 1 NQ. mức: 'THIẾU' / 'KHÔNG NHẬP'."""
    out = []
    for k, g, name, unit, th in BANH_ITEMS:
        p = res["pct"][k]
        if p is None:
            continue
        if res["act"][k] == 0:
            out.append((g, name, "KHÔNG NHẬP", p, th))
        elif p < th:
            out.append((g, name, "THIẾU", p, th))
    n = res["nuoc"]
    for key, label in (("sua", "Sữa bột"), ("duong", "Đường"), ("cacao", "Bột kem sữa cacao"), ("pepsi", "Nước ngọt Pepsi")):
        if n[key] == 0:
            out.append(("Nước", label, "KHÔNG NHẬP", -100.0, 0))
    if n["sua"] > 0 and n["pct_tra"] < -20:
        out.append(("Nước", "Trà đen (so với lượng sữa bột)", "THIẾU", n["pct_tra"], -20))
    if n["pct_ly"] is not None and n["pct_ly"] < TH_NUOC_LY:
        out.append(("Nước", "Tổng ly nước quy đổi từ NVL (so với ly SOP)", "THIẾU", n["pct_ly"], TH_NUOC_LY))
    return out


def run_check(pivot, master):
    results = OrderedDict()
    for ma, (ql, ten) in master.items():
        row = pivot.get(ma, {})
        res = compute_nq(row)
        res["ql"], res["ten"], res["co_don"] = ql, ten, bool(row)
        res["findings"] = build_findings(res) if row else []
        total_banh = sum(res["banh"].values())
        if not row:
            res["status"] = "KHÔNG CÓ ĐƠN NHẬP"
        elif total_banh == 0:
            res["status"] = "KHÔNG NHẬP BÁNH"
        elif res["findings"]:
            res["status"] = "THIẾU HÀNG"
        else:
            res["status"] = "ĐỦ HÀNG"
        results[ma] = res
    return results


# --------------------------------------------------------------------------------------
# XUẤT EXCEL
# --------------------------------------------------------------------------------------
RED = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")
GREEN = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
YELLOW = PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")
GREY = PatternFill("solid", bgColor="D9D9D9", fgColor="D9D9D9")
HEAD = PatternFill("solid", bgColor="1F4E78", fgColor="1F4E78")
REDFONT = Font(color="9C0006", bold=True)
HFONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILL = {"ĐỦ HÀNG": GREEN, "THIẾU HÀNG": RED, "KHÔNG CÓ ĐƠN NHẬP": GREY, "KHÔNG NHẬP BÁNH": YELLOW}


def _header(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.fill, c.font, c.border = HEAD, HFONT, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[row].height = 45
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def _fmt_pct(p):
    return "" if p is None else f"{p:+.0f}%"


def findings_text(res):
    parts = []
    for g, name, lvl, p, th in res["findings"]:
        short = name.split(" (")[0]
        parts.append(f"{short}: không nhập" if lvl == "KHÔNG NHẬP" else f"{short} {p:+.0f}%")
    return "; ".join(parts)


def findings_bullets(res):
    """Mỗi cảnh báo 1 dòng, bắt đầu bằng dấu •."""
    lines = []
    for g, name, lvl, p, th in res["findings"]:
        short = name.split(" (")[0]
        lines.append(f"• {short}: không nhập" if lvl == "KHÔNG NHẬP" else f"• {short}: {p:+.0f}%")
    return "\n".join(lines)


def sheet_summary(wb, results, title_info):
    ws = wb.active
    ws.title = "Tổng quan"
    ws["A1"] = f"BÁO CÁO KIỂM TRA NHẬP HÀNG NHƯỢNG QUYỀN - {title_info}"
    ws["A1"].font = Font(bold=True, size=14)
    n_ok = sum(r["status"] == "ĐỦ HÀNG" for r in results.values())
    n_th = sum(r["status"] == "THIẾU HÀNG" for r in results.values())
    n_other = len(results) - n_ok - n_th
    ws["A2"] = (f"Tổng {len(results)} NQ  |  Đủ hàng: {n_ok}  |  Thiếu hàng: {n_th}  |  "
                f"Không có đơn / không nhập bánh: {n_other}")
    ws["A3"] = "% = thực nhập*100/định mức SOP - 100 (âm = thiếu, dương = dư). Ngưỡng cảnh báo xem sheet 'Tham số'."
    hdr = ["STT", "Quản lý", "Mã NQ", "Nhượng quyền", "Tình trạng", "Số mục thiếu", "Nội dung cảnh báo",
           "Hamburger", "Hotdog", "Pita", "Sandwich", "Tổng ly SOP"]
    _header(ws, 5, hdr, [5, 10, 9, 46, 20, 9, 80, 11, 9, 9, 10, 11])
    order = {"THIẾU HÀNG": 0, "KHÔNG CÓ ĐƠN NHẬP": 1, "KHÔNG NHẬP BÁNH": 2, "ĐỦ HÀNG": 3}
    rows = sorted(results.items(), key=lambda kv: (order[kv[1]["status"]], -len(kv[1]["findings"]), kv[0]))
    for i, (ma, r) in enumerate(rows, 1):
        vals = [i, r["ql"], ma, r["ten"], r["status"], len(r["findings"]), findings_bullets(r),
                r["banh"]["hamburger"], r["banh"]["hotdog"], r["banh"]["pita"], r["banh"]["sandwich"],
                round(r["nuoc"]["ly_sop"])]
        for j, v in enumerate(vals, 1):
            c = ws.cell(5 + i, j, v)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[5 + i].height = max(15, 15 * len(r["findings"]))
        ws.cell(5 + i, 5).fill = STATUS_FILL[r["status"]]
    ws.freeze_panes = "E6"
    ws.auto_filter.ref = f"A5:{get_column_letter(len(hdr))}{5 + len(rows)}"


def sheet_matrix(wb, results):
    ws = wb.create_sheet("Ma trận %")
    ws["A1"] = "% nhập so với SOP bánh/nước (âm = thiếu, dương = dư). Ô đỏ = vượt ngưỡng cảnh báo."
    ws["A1"].font = Font(bold=True)
    hdr = ["Quản lý", "Mã NQ", "Nhượng quyền"] + [n for _, _, n, _, _ in BANH_ITEMS] + \
          ["Sữa bột (kg)", "Trà đen %", "Đường (kg)", "Cacao (kg)", "Pepsi (chai)", "Ly NVL quy đổi vs SOP %"]
    _header(ws, 3, hdr, [10, 9, 44] + [13] * (len(hdr) - 3))
    ths = [t for *_, t in BANH_ITEMS]
    for i, (ma, r) in enumerate(results.items()):
        row = 4 + i
        ws.cell(row, 1, r["ql"]); ws.cell(row, 2, ma); ws.cell(row, 3, r["ten"])
        for j, (k, *_x) in enumerate(BANH_ITEMS):
            p = r["pct"][k]
            c = ws.cell(row, 4 + j)
            if p is None or not r["co_don"]:
                c.value = None
            else:
                c.value = round(p, 1)
                c.number_format = '0.0"%"'
                if p < ths[j]:
                    c.fill, c.font = RED, REDFONT
        n = r["nuoc"]
        base = 4 + len(BANH_ITEMS)
        vals = [n["sua"], n["pct_tra"] if n["sua"] > 0 else None, n["duong"], n["cacao"], n["pepsi"], n["pct_ly"]]
        for j, v in enumerate(vals):
            c = ws.cell(row, base + j, None if v is None or not r["co_don"] else round(v, 1))
            if j == 1 or j == 5:
                c.number_format = '0.0"%"'
            bad = False
            if r["co_don"]:
                if j in (0, 2, 3, 4) and v == 0:
                    bad = True
                if j == 1 and v is not None and v < -20:
                    bad = True
                if j == 5 and v is not None and v < TH_NUOC_LY:
                    bad = True
            if bad:
                c.fill, c.font = RED, REDFONT
    ws.freeze_panes = "D4"


def detail_rows(ma, r):
    rows = []
    for k, g, name, unit, th in BANH_ITEMS:
        p = r["pct"][k]
        if p is None:
            st = "Không có SOP (không nhập bánh)"
        elif r["act"][k] == 0:
            st = "KHÔNG NHẬP"
        elif p < th:
            st = "THIẾU"
        else:
            st = "Đủ"
        rows.append([r["ql"], ma, r["ten"], g, name, unit, round(r["act"][k], 2), round(r["sop"][k], 2),
                     None if p is None else round(p, 1), th, st])
    n = r["nuoc"]
    def add(name, unit, a, s, p, th, flag_zero=False):
        if flag_zero and a == 0:
            st = "KHÔNG NHẬP"
        elif p is not None and th is not None and p < th:
            st = "THIẾU"
        else:
            st = "Đủ"
        rows.append([r["ql"], ma, r["ten"], "Nước", name, unit, round(a, 2), None if s is None else round(s, 2),
                     None if p is None else round(p, 1), th, st])
    add("Sữa bột", "kg", n["sua"], None, None, 0, True)
    add("Trà đen (so với lượng sữa bột)", "kg", n["tra"], n["tra_sop"], n["pct_tra"] if n["sua"] > 0 else None, -20)
    add("Đường", "kg", n["duong"], n["duong_sop"], None, 0, True)
    add("Bột kem sữa cacao", "kg", n["cacao"], None, None, 0, True)
    add("Nước ngọt Pepsi", "chai", n["pepsi"], None, None, 0, True)
    add("Bột chanh", "kg", n["chanh"], None, None, None)
    add("Tổng ly nước quy đổi từ NVL (so với ly SOP)", "ly", n["ly_nvl_tong"], n["ly_sop"], n["pct_ly"], TH_NUOC_LY)
    return rows


DETAIL_HDR = ["Quản lý", "Mã NQ", "Nhượng quyền", "Nhóm", "Nguyên liệu", "ĐVT so sánh", "Thực nhập (quy đổi)",
              "Định mức SOP", "% chênh", "Ngưỡng cảnh báo %", "Kết quả"]


def _write_detail(ws, start_row, rows):
    _header(ws, start_row, DETAIL_HDR, [10, 9, 44, 8, 44, 11, 14, 13, 10, 12, 28])
    for i, rw in enumerate(rows):
        for j, v in enumerate(rw, 1):
            c = ws.cell(start_row + 1 + i, j, v)
            c.border = BORDER
        st = rw[-1]
        c = ws.cell(start_row + 1 + i, len(rw))
        if st in ("THIẾU", "KHÔNG NHẬP"):
            c.fill, c.font = RED, REDFONT
        elif st == "Đủ":
            c.fill = GREEN
        else:
            c.fill = GREY


def sheet_detail(wb, results):
    ws = wb.create_sheet("Chi tiết từng NQ")
    rows = []
    for ma, r in results.items():
        if r["co_don"]:
            rows += detail_rows(ma, r)
    _write_detail(ws, 1, rows)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(DETAIL_HDR))}{1 + len(rows)}"


def sheet_nuoc(wb, results):
    ws = wb.create_sheet("SOP nước")
    hdr = ["Quản lý", "Mã NQ", "Nhượng quyền", "Ly theo SOP (từ bánh)", "Ly Trà sữa (từ sữa bột)", "Ly Cacao",
           "Ly Pepsi", "Ly Hồng trà chanh", "Tổng ly từ NVL", "% chênh so SOP", "Sữa bột (kg)", "Trà đen (kg)",
           "Trà đen theo SOP (kg)", "Đường (kg)", "Đường theo SOP (kg)", "Kết quả"]
    _header(ws, 1, hdr, [10, 9, 44] + [13] * 12 + [16])
    i = 0
    for ma, r in results.items():
        if not r["co_don"]:
            continue
        n = r["nuoc"]
        lv = n["ly_nvl"]
        ok = n["pct_ly"] is None or n["pct_ly"] >= TH_NUOC_LY
        vals = [r["ql"], ma, r["ten"], round(n["ly_sop"]), round(lv["Trà sữa"]), round(lv["Cacao"]), round(lv["Pepsi"]),
                round(lv["Hồng trà chanh"]), round(n["ly_nvl_tong"]),
                None if n["pct_ly"] is None else round(n["pct_ly"], 1),
                n["sua"], n["tra"], round(n["tra_sop"], 2), n["duong"], round(n["duong_sop"], 2),
                "Đủ" if ok else "THIẾU"]
        for j, v in enumerate(vals, 1):
            ws.cell(2 + i, j, v).border = BORDER
        ws.cell(2 + i, 16).fill = GREEN if ok else RED
        i += 1
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:P{1 + i}"


def sheet_unknown(wb, df, master):
    ws = wb.create_sheet("KH ngoài danh mục")
    unk = df[~df["ma_kh"].isin(master.keys())]
    ws["A1"] = ("Khách hàng có trong báo cáo tuần nhưng KHÔNG có trong danh mục NQ (sheet KIỂM TRA của CHECK_2026.xlsx). "
                "Đã chèn vào các báo cáo với Quản lý = 'Ngoài danh mục'. Nếu là NQ mới -> bổ sung vào danh mục.")
    ws["A1"].font = Font(bold=True)
    _header(ws, 3, ["Mã KH", "Tên khách hàng (đã khôi phục)", "Số dòng hàng", "Tổng tiền",
                    "Tên gốc trong file .xls", "Ghi chú"], [12, 60, 14, 16, 60, 55])
    g = unk.groupby("ma_kh").agg(n=("ma_hh", "count"), tien=("tien", "sum")).reset_index()
    for i, r in g.iterrows():
        goc = best_name(unk.loc[unk["ma_kh"] == r.ma_kh, "ten_kh_goc"])
        ten = best_name(unk.loc[unk["ma_kh"] == r.ma_kh, "ten_kh"])
        if "?" in ten:
            note = "Vẫn còn ký tự '?' - bổ sung từ vào workflow/tu_dien_phuc_hoi.txt"
        elif "?" in goc:
            note = "Tên khôi phục tự động từ file mất dấu - cần kiểm tra lại"
        else:
            note = ""
        for j, v in enumerate([r.ma_kh, ten, int(r.n), float(r.tien) if pd.notna(r.tien) else 0, goc, note], 1):
            ws.cell(4 + i, j, v)
    return len(g)


def _text(ws, row, col, v):
    """Ghi giá trị; chuỗi bắt đầu bằng '=' được lưu dạng text (như gõ ' trước dấu =) để không thành công thức."""
    c = ws.cell(row, col, v)
    if isinstance(v, str) and v.startswith("="):
        c.data_type = "s"
        c.quotePrefix = True
    return c


def sheet_params(wb):
    ws = wb.create_sheet("Tham số")
    _header(ws, 1, ["Nguyên liệu", "Nhóm", "Ngưỡng cảnh báo thiếu (%)", "Công thức SOP"], [48, 10, 22, 90])
    formulas = {
        "thit": "Thịt = Hamburger + Pita (miếng). Thực nhập = thịt hamburger + gà + tôm + bò hỗn hợp",
        "ca": "(Hamburger+Hotdog+Pita) x 5g",
        "mayo": "(Hamburger+Pita) x 5g + Sandwich/2 x 5g",
        "dau": "(Thịt gà + thịt tôm) / 80  (lít)",
        "hop": "Hamburger + Pita (cái). Thực nhập = Hộp Burger Gà + Hộp Pita Gà",
        "tui_hd": "= số bánh Hotdog",
        "hop_sw": "= Sandwich / 2",
        "pate": "Sandwich/2 x 5g. Pate 0.5kg/bịch -> x500g",
        "jam": "Sandwich/2 miếng. Jambon 1kg nhập = 1000/250*16 = 64 miếng",
        "chabong": "Sandwich/2 x 5g. Nhập kg x1000",
        "xucxich": "= số bánh Hotdog (cây)",
        "bao": "Hamburger+Hotdog+Pita+Sandwich/2 (ly). Túi dựng ly: 1kg = 192 cái",
        "ly": "Hamburger+Hotdog+Pita+Sandwich/2",
        "mang": "như ly. 1 cuộn = 3600 ly",
        "ong": "như ly. 1 bao = 250 cây",
    }
    for i, (k, g, n, u, th) in enumerate(BANH_ITEMS):
        for j, v in enumerate([n, g, th, formulas[k]], 1):
            _text(ws, 2 + i, j, v)
    r0 = 3 + len(BANH_ITEMS)
    extra = [
        ("Sữa bột / Đường / Cacao / Pepsi", "Nước", "= 0 là KHÔNG NHẬP", "Cảnh báo khi không nhập"),
        ("Trà đen (vs sữa bột)", "Nước", -20, "Trà đen SOP = sữa bột x 200/1100 (kg); chỉ xét khi có nhập sữa bột"),
        ("Tổng ly nước quy đổi từ NVL", "Nước", TH_NUOC_LY,
         "Trà sữa: 1kg sữa = 65.45 ly; Cacao: 1kg = 32 ly; Pepsi: 1 chai = 9 ly; Chanh: 1kg = 70 ly. So với ly SOP"),
    ]
    for i, rw in enumerate(extra):
        for j, v in enumerate(rw, 1):
            _text(ws, r0 + i, j, v)


def write_report(path, results, df, master, title_info):
    wb = openpyxl.Workbook()
    sheet_summary(wb, results, title_info)
    sheet_matrix(wb, results)
    sheet_detail(wb, results)
    sheet_nuoc(wb, results)
    n_unk = sheet_unknown(wb, df, master)
    sheet_params(wb)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return n_unk


def write_per_nq(folder, results, title_info):
    os.makedirs(folder, exist_ok=True)
    for ma, r in results.items():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = ma[:31]
        ws["A1"] = f"{r['ten']} ({ma}) - QL: {r['ql']} - {title_info}"
        ws["A1"].font = Font(bold=True, size=13)
        ws["A2"] = "Tình trạng:"
        ws["B2"] = r["status"]
        ws["B2"].fill = STATUS_FILL[r["status"]]
        ws["B2"].font = Font(bold=True)
        ws["A3"] = "Cảnh báo:"
        ws["B3"] = findings_bullets(r) or "Không có"
        ws["B3"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[3].height = max(15, 15 * len(r["findings"]))
        ws["A4"] = (f"Bánh nhập: Hamburger {r['banh']['hamburger']:.0f} | Hotdog {r['banh']['hotdog']:.0f} | "
                    f"Pita {r['banh']['pita']:.0f} | Sandwich {r['banh']['sandwich']:.0f}")
        if r["co_don"]:
            _write_detail(ws, 6, detail_rows(ma, r))
        safe = re.sub(r'[\\/:*?"<>|]', "", f"{ma} {r['ten']}")[:90]
        wb.save(os.path.join(folder, f"{safe}.xlsx"))


# --------------------------------------------------------------------------------------
# ĐỐI CHIẾU VỚI CHECK_2026.xlsx (kiểm chứng logic)
# --------------------------------------------------------------------------------------
def validate(master_path):
    """Dựng lại pivot từ sheet 'Dữ liệu thô' của CHECK_2026 rồi so với giá trị đã tính sẵn ở sheet KIỂM TRA."""
    wbf = openpyxl.load_workbook(master_path, data_only=False)
    wbv = openpyxl.load_workbook(master_path, data_only=True)
    ct, raw = wbf["ChiTiet"], wbf["Dữ liệu thô"]
    code_col = {}
    for c in range(6, 42):
        code, idx = ct.cell(1, c).value, ct.cell(2, c).value
        if code and idx:
            code_col[str(int(code))] = int(idx)
    pivot = {}
    for r in range(3, raw.max_row + 1):
        kh = raw.cell(r, 1).value
        if not kh or kh == "(trống)":
            continue
        pivot[kh] = {code: (raw.cell(r, col).value or 0) for code, col in code_col.items()}
    master = load_master(master_path)
    results = run_check(pivot, master)
    kt = wbv["KIỂM TRA"]
    cols = {"thit": 4, "ca": 5, "mayo": 6, "dau": 7, "hop": 8, "tui_hd": 9, "hop_sw": 10, "pate": 11, "jam": 12,
            "chabong": 13, "xucxich": 14, "bao": 15, "ly": 16, "mang": 17, "ong": 18}
    bad = n = 0
    for r in range(3, 97):
        ma = kt.cell(r, 2).value
        if ma not in results:
            continue
        for k, c in cols.items():
            exp, got = kt.cell(r, c).value, results[ma]["pct"][k]
            if got is None or not isinstance(exp, (int, float)):
                continue
            n += 1
            if abs(exp - got) > 0.01:
                bad += 1
                print(f"LỆCH {ma} {k}: Excel={exp:.3f} Python={got:.3f}")
    print(f"Đối chiếu {n} ô: lệch {bad}")
    return bad == 0


def _safe_save_path(path):
    """Nếu file đang mở (PermissionError), tự động tăng hậu tố (v2), (v3)... để không lỗi."""
    if not os.path.exists(path):
        return path
    try:
        with open(path, "ab"):
            return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        for v in range(2, 100):
            cand = f"{base} (v{v}){ext}"
            if not os.path.exists(cand):
                return cand
            try:
                with open(cand, "ab"):
                    return cand
            except PermissionError:
                continue
        return path


def process_report(report_path, master_path=DEFAULT_MASTER, out_path=None, per_nq=False):
    """Chạy toàn bộ quy trình kiểm tra từ file báo cáo tuần và trả về đường dẫn file output."""
    df, pivot = load_weekly_report(report_path)
    master_goc = load_master(master_path)
    df = fix_vietnamese(df, master_path, master_goc)
    master = extend_master(master_goc, df)
    results = run_check(pivot, master)

    d0, d1 = df["ngay"].min(), df["ngay"].max()
    title = f"Tuần {d0:%d/%m} - {d1:%d/%m/%Y}" if pd.notna(d0) else os.path.splitext(os.path.basename(report_path))[0]
    base = os.path.splitext(os.path.basename(report_path))[0]
    out = out_path or os.path.join(DEFAULT_OUT_DIR, f"BaoCao_NhapHang_{base}.xlsx")
    out = _safe_save_path(out)
    n_unk = write_report(out, results, df, master_goc, title)
    if per_nq:
        write_per_nq(os.path.splitext(out)[0] + "_NQ", results, title)

    cnt = defaultdict(int)
    for r in results.values():
        cnt[r["status"]] += 1
    print(f"{title}: {len(df)} dòng hàng, {len(pivot)} khách hàng, {len(master_goc)} NQ trong danh mục (+{len(master) - len(master_goc)} ngoài danh mục)")
    for k, v in cnt.items():
        print(f"  {k}: {v}")
    if n_unk:
        print(f"  Có {n_unk} KH ngoài danh mục, đã chèn vào các báo cáo (Quản lý = '{NGOAI_DM}')")
    print("Đã xuất:", out)
    return out


# --------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Kiểm tra nhập hàng nhượng quyền theo tuần")
    ap.add_argument("report", nargs="?", help="File báo cáo tuần (.xls/.xlsx)")
    ap.add_argument("--master", default=DEFAULT_MASTER, help="File CHECK_2026.xlsx (danh mục NQ)")
    ap.add_argument("--out", default=None, help="File Excel kết quả")
    ap.add_argument("--per-nq", action="store_true", help="Xuất thêm 1 file Excel riêng cho từng NQ")
    ap.add_argument("--validate", action="store_true", help="Đối chiếu logic với CHECK_2026.xlsx")
    a = ap.parse_args()

    if a.validate:
        sys.exit(0 if validate(a.master) else 1)
    if not a.report:
        ap.error("Cần truyền file báo cáo tuần")

    process_report(a.report, a.master, a.out, a.per_nq)


if __name__ == "__main__":
    main()


