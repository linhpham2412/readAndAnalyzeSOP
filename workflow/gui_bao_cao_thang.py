# -*- coding: utf-8 -*-
"""
Giao diện kéo-thả / chọn file báo cáo tuần (.xls / .xlsx / .zip) để tổng hợp
báo cáo doanh số Nhượng Quyền theo tháng (`BaoCaoThangNQ.exe`).
- Cho phép kéo thả hoặc bấm chọn file tuần (giống `KiemTraNhapHang.exe`).
- Tự động nhận diện ngày/tháng/năm/tuần từ tên file (vd `16 - 22.11.xls`);
  nếu không nhận diện được thì nhắc người dùng chọn Năm / Tháng / Tuần trên giao diện.
- Cho phép chọn/thay đổi file báo cáo tháng gốc để cộng dồn (mặc định tự tìm bản mới nhất).
- Sau khi xuất báo cáo -> tự động mở file Excel kết quả và đóng ứng dụng.
"""
import datetime as dt
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False

if getattr(sys, "frozen", False):
    ROOT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.join(ROOT_DIR, "workflow"))
if os.path.dirname(os.path.abspath(__file__)) not in sys.path:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import bao_cao_thang as core  # noqa: E402


def parse_dropped_files(raw_data: str, root: tk.Tk = None) -> list:
    """Trích xuất danh sách đường dẫn từ chuỗi sự kiện kéo thả của TkDND."""
    raw = (raw_data or "").strip()
    if not raw:
        return []
    if root is not None:
        try:
            items = list(root.tk.splitlist(raw))
            if items:
                return [s.strip().strip('"') for s in items if s and s.strip()]
        except Exception:
            pass
    # Fallback regex cho chuỗi dạng {path 1} {path 2} hoặc path1 path2
    tokens = re.findall(r"\{([^}]+)\}|(\S+)", raw)
    return [a or b for a, b in tokens if (a or b)]


def parse_dropped_file(raw_data: str) -> str:
    """Trích xuất đường dẫn đầu tiên từ chuỗi sự kiện kéo thả của TkDND."""
    files = parse_dropped_files(raw_data)
    return files[0] if files else ""


class MonthlyReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Tổng hợp Doanh số Nhượng Quyền theo Tháng - 1 Phút 30 Giây")
        self.root.geometry("660x530")
        self.root.resizable(False, False)
        self.root.configure(bg="#F5F7FA")

        today = dt.date.today()
        self.selected_files: list = []
        self.selected_file = tk.StringVar(value="")
        try:
            default_base = core.find_default_base_report()
        except Exception:
            default_base = ""
        self.base_file = tk.StringVar(value=default_base)

        self.year_var = tk.IntVar(value=today.year)
        self.month_var = tk.IntVar(value=today.month)
        self.days_var = tk.IntVar(value=7)
        self.week_choices = []
        self.week_var = tk.StringVar(value="")
        self.mode_badge_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Sẵn sàng. Kéo thả hoặc chọn file báo cáo tuần (.xls / .xlsx / .zip).")

        self._build_ui()
        self._refresh_week_slots(keep_idx=1)
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1F4E78", height=54)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="TỔNG HỢP BÁO CÁO DOANH SỐ NHƯỢNG QUYỀN THEO THÁNG",
            bg="#1F4E78",
            fg="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
        ).pack(expand=True)

        body = tk.Frame(self.root, bg="#F5F7FA", padx=22, pady=14)
        body.pack(fill="both", expand=True)

        # 1. Khung kéo thả file tuần
        self.drop_frame = tk.Frame(
            body,
            bg="#FFFFFF",
            highlightbackground="#4A90E2",
            highlightcolor="#1F4E78",
            highlightthickness=2,
            cursor="hand2",
            height=115,
        )
        self.drop_frame.pack(fill="x", pady=(0, 10))
        self.drop_frame.pack_propagate(False)

        self.drop_title = tk.Label(
            self.drop_frame,
            text="📥  Kéo và thả 1 hoặc nhiều file báo cáo tuần (.xls / .xlsx / .zip) vào đây\nhoặc bấm vào khung này để chọn file",
            bg="#FFFFFF",
            fg="#1F4E78",
            font=("Segoe UI", 10, "bold"),
            justify="center",
        )
        self.drop_title.pack(expand=True)
        for w in (self.drop_frame, self.drop_title):
            w.bind("<Button-1>", lambda _e: self.browse_weekly_file())

        if HAS_DND:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self.on_drop)
            self.drop_frame.dnd_bind("<<DragEnter>>", lambda _e: self.drop_frame.configure(bg="#EBF3FC") or self.drop_title.configure(bg="#EBF3FC"))
            self.drop_frame.dnd_bind("<<DragLeave>>", lambda _e: self.drop_frame.configure(bg="#FFFFFF") or self.drop_title.configure(bg="#FFFFFF"))

        # Đường dẫn file tuần
        path_row = tk.Frame(body, bg="#F5F7FA")
        path_row.pack(fill="x", pady=(0, 10))
        tk.Label(path_row, text="File tuần:", bg="#F5F7FA", fg="#334155", font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
        tk.Entry(
            path_row, textvariable=self.selected_file, font=("Segoe UI", 9),
            state="readonly", readonlybackground="#FFFFFF", fg="#1E293B", relief="solid", bd=1,
        ).pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))
        self.browse_multi_btn = tk.Button(
            path_row, text="Chọn nhiều file...", command=self.browse_weekly_files,
            bg="#DBEAFE", fg="#1E3A8A", activebackground="#BFDBFE",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2",
        )
        self.browse_multi_btn.pack(side="right", padx=(6, 0))
        self.browse_btn = tk.Button(
            path_row, text="Chọn file...", command=self.browse_weekly_file,
            bg="#E2E8F0", fg="#1E293B", activebackground="#CBD5E1",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=4, cursor="hand2",
        )
        self.browse_btn.pack(side="right")

        # File tháng gốc để cộng dồn
        base_row = tk.Frame(body, bg="#F5F7FA")
        base_row.pack(fill="x", pady=(0, 12))
        tk.Label(base_row, text="File tháng gốc:", bg="#F5F7FA", fg="#334155", font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
        tk.Entry(
            base_row, textvariable=self.base_file, font=("Segoe UI", 9),
            state="readonly", readonlybackground="#FFFFFF", fg="#1E293B", relief="solid", bd=1,
        ).pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))
        self.browse_base_btn = tk.Button(
            base_row, text="Đổi file gốc...", command=self.browse_base_file,
            bg="#E2E8F0", fg="#1E293B", activebackground="#CBD5E1",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2",
        )
        self.browse_base_btn.pack(side="right")

        # 2. Khung chọn Năm / Tháng / Tuần / Số ngày KD
        period_box = tk.LabelFrame(
            body, text="  Thời gian báo cáo (tự động nhận diện từ tên file hoặc chọn thủ công)  ",
            bg="#FFFFFF", fg="#1F4E78", font=("Segoe UI", 9, "bold"), padx=12, pady=10,
        )
        period_box.pack(fill="x", pady=(0, 12))

        row1 = tk.Frame(period_box, bg="#FFFFFF")
        row1.pack(fill="x", pady=(0, 8))

        tk.Label(row1, text="Năm:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left")
        year_cb = ttk.Combobox(row1, textvariable=self.year_var, values=list(range(2024, 2031)), width=6, state="readonly")
        year_cb.pack(side="left", padx=(6, 16))
        year_cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_week_slots())

        tk.Label(row1, text="Tháng:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left")
        month_cb = ttk.Combobox(row1, textvariable=self.month_var, values=list(range(1, 13)), width=4, state="readonly")
        month_cb.pack(side="left", padx=(6, 16))
        month_cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_week_slots())

        tk.Label(row1, text="Số ngày KD mặc định:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Spinbox(row1, from_=1, to=7, textvariable=self.days_var, width=4, font=("Segoe UI", 9)).pack(side="left", padx=(6, 0))

        row2 = tk.Frame(period_box, bg="#FFFFFF")
        row2.pack(fill="x", pady=(0, 6))
        tk.Label(row2, text="Tuần trong tháng:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.week_cb = ttk.Combobox(row2, textvariable=self.week_var, width=44, state="readonly")
        self.week_cb.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.week_cb.bind("<<ComboboxSelected>>", lambda _e: self._update_mode_badge())

        self.mode_lbl = tk.Label(
            period_box, textvariable=self.mode_badge_var,
            bg="#EFF6FF", fg="#1E40AF", font=("Segoe UI", 9, "bold"), anchor="w", padx=8, pady=4,
        )
        self.mode_lbl.pack(fill="x", pady=(2, 0))

        # 3. Nút xử lý
        self.run_btn = tk.Button(
            body,
            text="▶   TỔNG HỢP VÀ MỞ BÁO CÁO THÁNG",
            command=self.start_processing,
            bg="#1F4E78",
            fg="#FFFFFF",
            activebackground="#163A59",
            activeforeground="#FFFFFF",
            disabledforeground="#CBD5E1",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            pady=10,
            cursor="hand2",
            state="disabled",
        )
        self.run_btn.pack(fill="x", pady=(0, 8))

        tk.Label(
            body, textvariable=self.status_var,
            bg="#F5F7FA", fg="#475569", font=("Segoe UI", 9), anchor="w",
        ).pack(fill="x")

    def _refresh_week_slots(self, keep_idx=None):
        try:
            y = int(self.year_var.get())
            m = int(self.month_var.get())
        except Exception:
            return
        self.week_choices = core.get_month_week_slots(y, m)
        labels = [f"Tuần thứ {s['slot_idx']} trong tháng  —  {s['label']} ({s['days_count']} ngày)" for s in self.week_choices]
        self.week_cb["values"] = labels
        idx = (keep_idx - 1) if (keep_idx and 1 <= keep_idx <= len(labels)) else 0
        self.week_var.set(labels[idx])
        self.days_var.set(self.week_choices[idx]["days_count"])
        self._update_mode_badge()

    def _selected_slot_idx(self) -> int:
        cur = self.week_var.get()
        vals = list(self.week_cb["values"])
        if cur in vals:
            return vals.index(cur) + 1
        return 1

    def _update_mode_badge(self):
        sidx = self._selected_slot_idx()
        if sidx == 1:
            self.mode_badge_var.set("🔄 Tuần đầu tiên của tháng: sẽ làm mới (clean up) dữ liệu các tuần trên Dashboard.")
            self.mode_lbl.configure(bg="#FEF3C7", fg="#92400E")
        else:
            self.mode_badge_var.set(f"➕ Cộng dồn: giữ nguyên các tuần trước và cập nhật vào vị trí tuần thứ {sidx} trong tháng.")
            self.mode_lbl.configure(bg="#DCFCE7", fg="#166534")
        if 1 <= sidx <= len(self.week_choices):
            self.days_var.set(self.week_choices[sidx - 1]["days_count"])

    def set_file(self, path: str):
        self.set_files([path])

    def set_files(self, paths: list):
        valid_paths = []
        for raw_p in paths:
            p = os.path.abspath(str(raw_p).strip().strip('"'))
            if not os.path.isfile(p):
                messagebox.showwarning("File không hợp lệ", f"Không tìm thấy file:\n{p}")
                return
            ext = os.path.splitext(p)[1].lower()
            if ext not in (".xls", ".xlsx", ".zip"):
                messagebox.showwarning("Sai định dạng", f"File không đúng định dạng (.xls, .xlsx, .zip):\n{os.path.basename(p)}")
                return
            if p not in valid_paths:
                valid_paths.append(p)

        if not valid_paths:
            return

        self.selected_files = valid_paths
        if len(valid_paths) == 1:
            p = valid_paths[0]
            self.selected_file.set(p)
            self.drop_title.configure(
                text=f"📄  Đã chọn: {os.path.basename(p)}\n(Kéo thả file khác hoặc bấm để đổi file)",
                fg="#0F5132",
            )
        else:
            names = ", ".join(os.path.basename(p) for p in valid_paths)
            self.selected_file.set(f"[{len(valid_paths)} file] {names}")
            self.drop_title.configure(
                text=f"📚  Đã chọn {len(valid_paths)} file dữ liệu (sẽ gộp toàn bộ các tab)\n(Kéo thả hoặc bấm để chọn lại)",
                fg="#0F5132",
            )
        self.drop_frame.configure(highlightbackground="#198754")
        self.run_btn.configure(state="normal", bg="#198754", activebackground="#146C43")

        parsed = core.parse_date_range_from_filenames(valid_paths)
        if parsed:
            self.year_var.set(parsed["year"])
            self.month_var.set(parsed["month"])
            self._refresh_week_slots(keep_idx=parsed["slot_idx"])
            self.days_var.set(parsed["days_count"])
            self.status_var.set(f"Đã nhận diện: {parsed['slot']['label']} ({parsed['days_count']} ngày). Sẵn sàng tổng hợp.")
        else:
            self.status_var.set("Không tự nhận diện được ngày từ tên file — vui lòng chọn Năm / Tháng / Tuần ở khung bên dưới.")
            messagebox.showinfo(
                "Chọn tuần báo cáo",
                "Không tự nhận diện được ngày/tháng từ tên file.\nVui lòng chọn Năm, Tháng và Tuần trong tháng trước khi bấm Tổng hợp.",
            )

    def on_drop(self, event):
        self.drop_frame.configure(bg="#FFFFFF")
        self.drop_title.configure(bg="#FFFFFF")
        paths = parse_dropped_files(event.data, root=self.root)
        if paths:
            self.set_files(paths)

    def browse_weekly_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Chọn file báo cáo tuần",
            initialdir=ROOT_DIR,
            filetypes=[("Excel & ZIP files", "*.xls *.xlsx *.zip"), ("All files", "*.*")],
        )
        if path:
            self.set_files([path])

    def browse_weekly_files(self):
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="Chọn nhiều file báo cáo (gộp toàn bộ các tab)",
            initialdir=ROOT_DIR,
            filetypes=[("Excel & ZIP files", "*.xls *.xlsx *.zip"), ("All files", "*.*")],
        )
        if paths:
            self.set_files(list(paths))

    def browse_base_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Chọn file báo cáo tháng gốc",
            initialdir=os.path.dirname(self.base_file.get()) if self.base_file.get() else ROOT_DIR,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.base_file.set(os.path.abspath(path))

    def start_processing(self):
        report_paths = list(self.selected_files)
        if not report_paths:
            single = self.selected_file.get()
            if single and os.path.isfile(single):
                report_paths = [single]
        if not report_paths or not all(os.path.isfile(p) for p in report_paths):
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất 1 file dữ liệu tuần.")
            return
        base_path = self.base_file.get()
        if not base_path or not os.path.isfile(base_path):
            messagebox.showwarning("Thiếu file gốc", "Vui lòng chọn file báo cáo tháng gốc (.xlsx).")
            return

        y = int(self.year_var.get())
        m = int(self.month_var.get())
        sidx = self._selected_slot_idx()
        days = int(self.days_var.get())

        self.run_btn.configure(state="disabled", text="⏳  ĐANG XỬ LÝ VÀ TẠO BÁO CÁO...")
        self.browse_btn.configure(state="disabled")
        self.browse_multi_btn.configure(state="disabled")
        self.browse_base_btn.configure(state="disabled")
        self.status_var.set("Đang đọc dữ liệu, chuyển mã Unicode và tạo báo cáo tháng...")

        weekly_arg = report_paths[0] if len(report_paths) == 1 else report_paths
        threading.Thread(
            target=self._worker,
            args=(weekly_arg, base_path, y, m, sidx, days),
            daemon=True,
        ).start()

    def _worker(self, report_path, base_path, y, m, sidx, days):
        try:
            res = core.generate_monthly_report(
                weekly_path=report_path,
                base_path=base_path,
                year=y,
                month=m,
                slot_idx=sidx,
                operating_days_override=days,
            )
            self.root.after(0, lambda: self._on_success(res))
        except Exception as exc:
            self.root.after(0, lambda err=str(exc): self._on_error(err))

    def _on_success(self, res: dict):
        out_file = res["out_file"]
        self.status_var.set(f"Hoàn tất! Đang mở: {out_file}")
        try:
            os.startfile(os.path.abspath(out_file))
        except Exception as exc:
            messagebox.showwarning("Không tự mở được file", f"Đã xuất báo cáo tại:\n{out_file}\n\nChi tiết: {exc}")
        self.root.after(300, self.root.destroy)

    def _on_error(self, err_msg: str):
        self.run_btn.configure(state="normal", text="▶   TỔNG HỢP VÀ MỞ BÁO CÁO THÁNG")
        self.browse_btn.configure(state="normal")
        self.browse_multi_btn.configure(state="normal")
        self.browse_base_btn.configure(state="normal")
        self.status_var.set("Có lỗi xảy ra khi xử lý.")
        messagebox.showerror("Lỗi khi xử lý báo cáo", err_msg)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--run-cli":
        sys.argv = [sys.argv[0]] + sys.argv[2:]
        core.main()
        return
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    MonthlyReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
