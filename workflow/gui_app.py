# -*- coding: utf-8 -*-
"""
Giao diện kéo-thả / chọn file dữ liệu tuần để chạy kiểm tra nhập hàng nhượng quyền.
- Cho phép kéo thả file (.xls / .xlsx) hoặc bấm chọn file.
- Khi bấm "XỬ LÝ":
  + Nếu tồn tại .\workflow\kiem_tra_nhap_hang.py và Python ngoài hệ thống -> chạy lệnh
    python .\workflow\kiem_tra_nhap_hang.py "<file đã chọn>"
    (hoặc chạy trực tiếp module đã đóng gói trong .exe để đảm bảo hoạt động trên mọi máy).
- Sau khi xuất file output -> tự động mở file Excel output và đóng ứng dụng.
"""
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

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
import kiem_tra_nhap_hang as core  # noqa: E402


def parse_dropped_file(raw_data: str) -> str:
    """Trích xuất đường dẫn đầu tiên từ chuỗi sự kiện kéo thả của TkDND."""
    raw = (raw_data or "").strip()
    if not raw:
        return ""
    if raw.startswith("{"):
        m = re.match(r"^\{([^}]+)\}", raw)
        if m:
            return m.group(1)
    return raw.split()[0]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Kiểm tra nhập hàng Nhượng Quyền - 1 Phút 30 Giây")
        self.root.geometry("580x360")
        self.root.resizable(False, False)
        self.root.configure(bg="#F5F7FA")

        self.selected_file = tk.StringVar(value="")
        self.per_nq_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Sẵn sàng. Hãy kéo thả hoặc chọn file báo cáo tuần (.xls / .xlsx).")

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # Tiêu đề
        header = tk.Frame(self.root, bg="#1F4E78", height=54)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="KIỂM TRA NHẬP HÀNG NHƯỢNG QUYỀN THEO TUẦN",
            bg="#1F4E78",
            fg="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
        ).pack(expand=True)

        body = tk.Frame(self.root, bg="#F5F7FA", padx=22, pady=16)
        body.pack(fill="both", expand=True)

        # Khung kéo thả (Drop Zone)
        self.drop_frame = tk.Frame(
            body,
            bg="#FFFFFF",
            highlightbackground="#4A90E2",
            highlightcolor="#1F4E78",
            highlightthickness=2,
            bd=0,
            cursor="hand2",
            height=130,
        )
        self.drop_frame.pack(fill="x", pady=(0, 12))
        self.drop_frame.pack_propagate(False)

        self.drop_title = tk.Label(
            self.drop_frame,
            text="📥  Kéo và thả file báo cáo tuần (.xls / .xlsx) vào đây\nhoặc bấm vào khung này để chọn file",
            bg="#FFFFFF",
            fg="#1F4E78",
            font=("Segoe UI", 10, "bold"),
            justify="center",
        )
        self.drop_title.pack(expand=True)

        for widget in (self.drop_frame, self.drop_title):
            widget.bind("<Button-1>", lambda _e: self.browse_file())

        if HAS_DND:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self.on_drop)
            self.drop_frame.dnd_bind("<<DragEnter>>", lambda _e: self.drop_frame.configure(bg="#EBF3FC") or self.drop_title.configure(bg="#EBF3FC"))
            self.drop_frame.dnd_bind("<<DragLeave>>", lambda _e: self.drop_frame.configure(bg="#FFFFFF") or self.drop_title.configure(bg="#FFFFFF"))

        # Dòng hiển thị đường dẫn file + nút Chọn file
        path_row = tk.Frame(body, bg="#F5F7FA")
        path_row.pack(fill="x", pady=(0, 10))

        self.path_entry = tk.Entry(
            path_row,
            textvariable=self.selected_file,
            font=("Segoe UI", 9),
            state="readonly",
            readonlybackground="#FFFFFF",
            fg="#222222",
            relief="solid",
            bd=1,
        )
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))

        self.browse_btn = tk.Button(
            path_row,
            text="Chọn file...",
            command=self.browse_file,
            bg="#E2E8F0",
            fg="#1E293B",
            activebackground="#CBD5E1",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
        )
        self.browse_btn.pack(side="right")

        # Tùy chọn xuất file lẻ từng NQ + Nút Xử lý
        action_row = tk.Frame(body, bg="#F5F7FA")
        action_row.pack(fill="x", pady=(4, 10))

        tk.Checkbutton(
            action_row,
            text="Xuất thêm file riêng cho từng Nhượng Quyền (--per-nq)",
            variable=self.per_nq_var,
            bg="#F5F7FA",
            activebackground="#F5F7FA",
            font=("Segoe UI", 9),
        ).pack(side="left")

        self.run_btn = tk.Button(
            body,
            text="▶   XỬ LÝ VÀ MỞ BÁO CÁO",
            command=self.start_processing,
            bg="#1F4E78",
            fg="#FFFFFF",
            activebackground="#163A59",
            activeforeground="#FFFFFF",
            disabledforeground="#CBD5E1",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            pady=9,
            cursor="hand2",
            state="disabled",
        )
        self.run_btn.pack(fill="x", pady=(2, 8))

        # Thanh trạng thái
        tk.Label(
            body,
            textvariable=self.status_var,
            bg="#F5F7FA",
            fg="#475569",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x")

    def set_file(self, path: str):
        path = os.path.abspath(path.strip().strip('"'))
        if not os.path.isfile(path):
            messagebox.showwarning("File không hợp lệ", f"Không tìm thấy file:\n{path}")
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xls", ".xlsx"):
            messagebox.showwarning("Sai định dạng", "Vui lòng chọn file Excel (.xls hoặc .xlsx).")
            return
        self.selected_file.set(path)
        self.drop_title.configure(
            text=f"📄  Đã chọn: {os.path.basename(path)}\n(Kéo thả file khác hoặc bấm để đổi file)",
            fg="#0F5132",
        )
        self.drop_frame.configure(highlightbackground="#198754")
        self.run_btn.configure(state="normal", bg="#198754", activebackground="#146C43")
        self.status_var.set(f'Sẵn sàng chạy: .\\workflow\\kiem_tra_nhap_hang.py "{os.path.basename(path)}"')

    def on_drop(self, event):
        self.drop_frame.configure(bg="#FFFFFF")
        self.drop_title.configure(bg="#FFFFFF")
        path = parse_dropped_file(event.data)
        if path:
            self.set_file(path)

    def browse_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Chọn file báo cáo tuần",
            initialdir=ROOT_DIR,
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.set_file(path)

    def start_processing(self):
        report_path = self.selected_file.get()
        if not report_path or not os.path.isfile(report_path):
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn file dữ liệu tuần trước khi xử lý.")
            return

        self.run_btn.configure(state="disabled", text="⏳  ĐANG XỬ LÝ DỮ LIỆU...")
        self.browse_btn.configure(state="disabled")
        self.status_var.set(f'Đang thực thi .\\workflow\\kiem_tra_nhap_hang.py "{os.path.basename(report_path)}" ...')

        threading.Thread(target=self._worker, args=(report_path, self.per_nq_var.get()), daemon=True).start()

    def _worker(self, report_path: str, per_nq: bool):
        try:
            script_path = os.path.join(ROOT_DIR, "workflow", "kiem_tra_nhap_hang.py")
            master_path = os.path.join(ROOT_DIR, "CHECK_2026.xlsx")
            if not os.path.exists(master_path):
                master_path = core.DEFAULT_MASTER

            out_file = None
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            # Gọi lệnh .\workflow\kiem_tra_nhap_hang.py "<tên của file đã chọn>"
            if os.path.isfile(script_path):
                py_exe = "python" if getattr(sys, "frozen", False) else sys.executable
                cmd = [py_exe, os.path.join(".", "workflow", "kiem_tra_nhap_hang.py"), report_path]
                if per_nq:
                    cmd.append("--per-nq")
                try:
                    proc = subprocess.run(
                        cmd,
                        cwd=ROOT_DIR,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        creationflags=creationflags,
                    )
                    if proc.returncode == 0:
                        for line in reversed(proc.stdout.splitlines()):
                            if line.startswith("Đã xuất:"):
                                out_file = line.split("Đã xuất:", 1)[1].strip()
                                break
                    elif not getattr(sys, "frozen", False):
                        raise RuntimeError(proc.stderr or proc.stdout or "Lỗi khi chạy script")
                except FileNotFoundError:
                    pass

            # Nếu chạy trên máy không cài Python (hoặc chưa lấy được đường dẫn), chạy trực tiếp bằng bộ xử lý đã đóng gói trong .exe
            if not out_file or not os.path.isfile(out_file):
                out_file = core.process_report(report_path, master_path=master_path, per_nq=per_nq)

            self.root.after(0, lambda: self._on_success(out_file))
        except Exception as exc:
            self.root.after(0, lambda: self._on_error(str(exc)))

    def _on_success(self, out_file: str):
        self.status_var.set(f"Hoàn tất! Đang mở: {out_file}")
        try:
            os.startfile(os.path.abspath(out_file))
        except Exception as exc:
            messagebox.showwarning("Không tự mở được file", f"Đã xuất báo cáo tại:\n{out_file}\n\nChi tiết: {exc}")
        self.root.after(300, self.root.destroy)

    def _on_error(self, err_msg: str):
        self.run_btn.configure(state="normal", text="▶   XỬ LÝ VÀ MỞ BÁO CÁO")
        self.browse_btn.configure(state="normal")
        self.status_var.set("Có lỗi xảy ra khi xử lý.")
        messagebox.showerror("Lỗi khi xử lý báo cáo", err_msg)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--run-cli":
        sys.argv = [sys.argv[0]] + sys.argv[2:]
        core.main()
        return
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
