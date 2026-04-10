#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_check_kltn.py — Giao diện đồ họa kiểm tra định dạng KLTN
=============================================================
Chạy:   python gui_check_kltn.py
"""

import os, sys, re, threading
from pathlib import Path
from datetime import datetime

# ── Tự cài thư viện ──────────────────────────────────────────────
def _ensure(pkg, imp=None):
    try: __import__(imp or pkg)
    except ImportError:
        os.system(f"{sys.executable} -m pip install {pkg} -q")

_ensure("python-docx", "docx")
_ensure("openpyxl")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import engine từ check_format_kltn.py
sys.path.insert(0, str(Path(__file__).parent))
from check_format_kltn import check_file, scan_directory, export_excel, CheckResult


# ════════════════════════════════════════════════════════════════
#  MÀU SẮC (Sáng sủa & Đẹp mắt)
# ════════════════════════════════════════════════════════════════
C = {
    "bg":        "#F5F7FA",   # Nền Panel xám nhạt
    "panel":     "#FFFFFF",   # Màu trắng sáng trung tâm
    "card":      "#FFFFFF",   # Nền các khối
    "border":    "#E4E7EB",   # Đường viền
    "accent":    "#2563EB",   # Xanh lục (Tailwind blue-600)
    "acc2":      "#7C3AED",   # Tím (Tailwind violet-600)
    "green":     "#10B981",   # Xanh lá OK
    "yellow":    "#F59E0B",   # Vàng cảnh báo
    "red":       "#EF4444",   # Đỏ báo lỗi
    "text":      "#1F2937",   # Xám than sậm (Text chính)
    "text2":     "#4B5563",   # Text thứ cấp
    "text3":     "#9CA3AF",   # Mờ nhạt
    "row_even":  "#FFFFFF",
    "row_odd":   "#F9FAFB",
    "row_sel":   "#E0E7FF",
}

SEV_COLOR = {"ERROR": C["red"], "WARNING": C["yellow"], "INFO": C["green"]}
SEV_BG    = {"ERROR": "#FEE2E2", "WARNING": "#FEF3C7", "INFO": "#D1FAE5"}


# ════════════════════════════════════════════════════════════════
#  WIDGET TIỆN ÍCH
# ════════════════════════════════════════════════════════════════
class RoundButton(tk.Button):
    """Nút có màu nền tùy chỉnh, tương thích macOS."""
    def __init__(self, parent, text="", command=None, width=160, height=36,
                 bg=C["accent"], fg=C["text"], font=("Segoe UI", 10, "bold"),
                 radius=10, hover_bg=None, **kw):
        self._bg = bg
        self._hover = hover_bg or self._darken(bg)
        kw.pop("radius", None)
        super().__init__(
            parent, text=text, command=command,
            bg=bg, fg=fg, font=font,
            activebackground=self._hover, activeforeground=fg,
            relief="flat", bd=0, padx=12, pady=6,
            cursor="hand2", **kw
        )
        self.bind("<Enter>", lambda e: self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._bg))

    def _darken(self, hex_color):
        r, g, b = int(hex_color[1:3],16), int(hex_color[3:5],16), int(hex_color[5:7],16)
        f = 0.80
        return f"#{int(r*f):02x}{int(g*f):02x}{int(b*f):02x}"


class ScoreRing(tk.Canvas):
    """Vòng điểm số dạng arc."""
    def __init__(self, parent, size=110, **kw):
        super().__init__(parent, width=size, height=size,
                         bg=parent["bg"], highlightthickness=0, **kw)
        self._sz = size
        self._score = 0

    def set_score(self, score):
        self._score = score
        self._draw()

    def _draw(self):
        sz = self._sz
        self.delete("all")
        pad = 10
        # Track
        self.create_arc(pad, pad, sz-pad, sz-pad, start=90, extent=360,
                        outline=C["border"], width=8, style="arc")
        # Arc
        if self._score > 0:
            color = C["green"] if self._score >= 70 else (C["yellow"] if self._score >= 50 else C["red"])
            ext = int(self._score / 100 * 360)
            self.create_arc(pad, pad, sz-pad, sz-pad, start=90, extent=-ext,
                            outline=color, width=8, style="arc")
        # Số
        color = C["green"] if self._score >= 70 else (C["yellow"] if self._score >= 50 else C["red"])
        self.create_text(sz//2, sz//2 - 8, text=f"{self._score}", fill=color,
                         font=("Segoe UI", 18, "bold"))
        self.create_text(sz//2, sz//2 + 14, text="/ 100", fill=C["text2"],
                         font=("Segoe UI", 9))


# ════════════════════════════════════════════════════════════════
#  CỬA SỔ CHÍNH
# ════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CheckForm KLTN — v1.0")
        self.configure(bg=C["bg"])
        self.geometry("1080x720")
        self.minsize(980, 640)
        self.resizable(True, True)

        # Dữ liệu
        self._results: list[CheckResult] = []
        self._selected_idx = -1
        self._running = False
        self._stop_event = threading.Event()

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── BUILD UI ────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        hdr = tk.Frame(self, bg=C["panel"], height=64)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="⚙", bg=C["panel"], fg=C["accent"],
                 font=("Segoe UI", 20)).pack(side="left", padx=(18,6), pady=12)
        tk.Label(hdr, text="Kiểm tra Định dạng KLTN",
                 bg=C["panel"], fg=C["text"], font=("Segoe UI", 15, "bold")).pack(side="left", pady=12)
        tk.Label(hdr, text="Trường ĐH Thủy Lợi · Khoa Kinh tế và Quản lý - Khoa Kế toán và Kinh doanh",
                 bg=C["panel"], fg=C["text3"], font=("Segoe UI", 10)).pack(side="left", padx=14, pady=18)

        # Version badge
        badge = tk.Label(hdr, text=" v1.0 ", bg=C["acc2"], fg="#FFFFFF",
                         font=("Segoe UI", 8, "bold"))
        badge.pack(side="right", padx=18, pady=22)

        # ── Toolbar ──
        tb = tk.Frame(self, bg=C["bg"], pady=10)
        tb.pack(fill="x", padx=18)

        self._pick_btn = RoundButton(tb, "📄  Chọn File .docx", command=self._pick_files,
                    width=180, height=36, bg=C["accent"])
        self._pick_btn.pack(side="left", padx=(0,10))
        self._folder_btn = RoundButton(tb, "📂  Chọn Thư mục", command=self._pick_folder,
                    width=180, height=36, bg=C["acc2"])
        self._folder_btn.pack(side="left", padx=(0,10))

        self._run_btn = RoundButton(tb, "▶  Bắt đầu kiểm tra", command=self._run_check,
                                    width=190, height=36, bg="#2E6D45", hover_bg="#3A8F5C")
        self._run_btn.pack(side="left", padx=(0,10))

        self._export_btn = RoundButton(tb, "💾  Xuất Excel", command=self._export,
                    width=150, height=36, bg="#D97706", hover_bg="#F59E0B")
        self._export_btn.pack(side="left", padx=(0,10))

        # Label đường dẫn
        self._path_var = tk.StringVar(value="Chưa chọn file / thư mục")
        tk.Label(tb, textvariable=self._path_var, bg=C["bg"], fg=C["text3"],
                 font=("Segoe UI", 9), anchor="w").pack(side="left", padx=10)

        # Clear button & Config Button
        RoundButton(tb, "👥 Cấu hình GVHD", command=self._show_config_dialog,
                    width=170, height=36, bg="#4A235A", hover_bg="#6C3483").pack(side="right", padx=(10,0))

        RoundButton(tb, "✕  Xóa", command=self._clear,
                    width=90, height=36, bg=C["border"], hover_bg="#505580",
                    fg=C["text2"]).pack(side="right")

        # ── Body (paned) ──
        paned = tk.PanedWindow(self, orient="horizontal", bg=C["bg"],
                               sashwidth=6, sashrelief="flat", sashpad=2)
        paned.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        # Left — danh sách file
        left = tk.Frame(paned, bg=C["bg"])
        paned.add(left, width=480, minsize=300)
        self._build_file_list(left)

        # Right — chi tiết lỗi
        right = tk.Frame(paned, bg=C["bg"])
        paned.add(right, minsize=340)
        self._build_detail_panel(right)

        # ── Status bar (2 dòng) ──
        sb = tk.Frame(self, bg=C["panel"])
        sb.pack(fill="x", side="bottom")

        # Dòng 1: file hiện tại
        row1 = tk.Frame(sb, bg=C["panel"])
        row1.pack(fill="x", padx=10, pady=(4, 0))

        self._file_icon_var = tk.StringVar(value="")
        tk.Label(row1, textvariable=self._file_icon_var, bg=C["panel"],
                 fg=C["accent"], font=("Segoe UI", 9, "bold"), width=3).pack(side="left")

        self._file_name_var = tk.StringVar(value="Sẵn sàng.")
        tk.Label(row1, textvariable=self._file_name_var, bg=C["panel"], fg=C["text"],
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left", fill="x", expand=True)

        self._counter_var = tk.StringVar(value="")
        tk.Label(row1, textvariable=self._counter_var, bg=C["panel"], fg=C["text3"],
                 font=("Segoe UI", 9)).pack(side="right", padx=6)

        # Dòng 2: status text + progressbar
        row2 = tk.Frame(sb, bg=C["panel"])
        row2.pack(fill="x", padx=10, pady=(2, 4))

        self._status_var = tk.StringVar(value="")
        tk.Label(row2, textvariable=self._status_var, bg=C["panel"], fg=C["text3"],
                 font=("Segoe UI", 8), anchor="w").pack(side="left", fill="x", expand=True)

        self._progress = ttk.Progressbar(row2, mode="determinate", length=260)
        self._progress.pack(side="right", padx=(8, 0))

        # Spinner state
        self._spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self._spinner_idx = 0
        self._spinner_job = None

    # ── ĐỔI CẤU HÌNH GVHD ───────────────────────────────────────
    def _show_config_dialog(self):
        import json
        cfg_path = "config_kltn.json"
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except:
            cfg = {"advisors": []}

        dlg = tk.Toplevel(self)
        dlg.title("Cấu hình Danh sách Cán bộ Hướng dẫn")
        dlg.geometry("500x600")
        dlg.configure(bg=C["bg"])
        dlg.transient(self)
        dlg.grab_set()

        lbl = tk.Label(dlg, text="Danh sách Cán bộ Hướng dẫn (Mỗi GV 1 dòng):", bg=C["bg"], fg=C["text"], font=("Segoe UI", 11, "bold"))
        lbl.pack(pady=(20, 10), padx=20, anchor="w")
        
        lbl_hint = tk.Label(dlg, text="Bạn có thể copy và paste cả 1 danh sách từ Excel/Word vào đây.", bg=C["bg"], fg=C["text3"], font=("Segoe UI", 9))
        lbl_hint.pack(padx=20, anchor="w", pady=(0, 10))

        text_area = tk.Text(dlg, font=("Segoe UI", 11), bg=C["panel"], fg=C["text"], relief="flat", insertbackground=C["text"])
        text_area.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Đổ dữ liệu
        text_area.insert("1.0", "\n".join(cfg.get("advisors", [])))

        def _save():
            content = text_area.get("1.0", "end").strip()
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            cfg["advisors"] = lines
            try:
                with open(cfg_path, 'w', encoding='utf-8') as f:
                    json.dump(cfg, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Thành công", f"Đã lưu danh sách {len(lines)} giảng viên!")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu file config: {str(e)}")

        btn_frame = tk.Frame(dlg, bg=C["bg"])
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        RoundButton(btn_frame, "💾 Lưu cấu hình", command=_save, width=150, height=36, bg=C["accent"]).pack(side="right")
        RoundButton(btn_frame, "Hủy", command=dlg.destroy, width=100, height=36, bg=C["border"], hover_bg="#505580", fg=C["text2"]).pack(side="right", padx=(0, 10))


    # ── FILE LIST ───────────────────────────────────────────────
    def _build_file_list(self, parent):
        # Header
        hdr = tk.Frame(parent, bg=C["card"], pady=8, padx=12)
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="Danh sách file kiểm tra", bg=C["card"],
                 fg=C["text"], font=("Segoe UI", 11, "bold")).pack(side="left")
        self._count_lbl = tk.Label(hdr, text="", bg=C["card"], fg=C["text3"],
                                   font=("Segoe UI", 9))
        self._count_lbl.pack(side="right")

        # Summary bar
        self._sum_frame = tk.Frame(parent, bg=C["bg"])
        self._sum_frame.pack(fill="x", pady=(0, 4))

        # Treeview
        cols = ("filename", "student", "mssv", "advisor", "errors", "warns", "score", "letter", "eval")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings",
                                  selectmode="browse")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                         background=C["row_odd"], foreground=C["text"],
                         fieldbackground=C["row_odd"], borderwidth=0,
                         font=("Segoe UI", 9), rowheight=28)
        style.configure("Treeview.Heading",
                         background=C["card"], foreground=C["text2"],
                         font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0)
        style.map("Treeview",
                  background=[("selected", C["row_sel"])],
                  foreground=[("selected", C["text"])])

        col_conf = {
            "filename": ("Tên file",    180, "w"),
            "student":  ("Sinh viên",   120, "w"),
            "mssv":     ("MSSV",        90, "center"),
            "advisor":  ("GVHD",        100, "w"),
            "errors":   ("Lỗi",         45, "center"),
            "warns":    ("Cảnh báo",    55, "center"),
            "score":    ("Điểm",        45, "center"),
            "letter":   ("Chữ",         45, "center"),
            "eval":     ("Đánh giá",    80, "center"),
        }
        for col, (heading, width, anchor) in col_conf.items():
            self._tree.heading(col, text=heading,
                               command=lambda c=col: self._sort_tree(c, False))
            self._tree.column(col, width=width, anchor=anchor, minwidth=30)

        self._tree.tag_configure("even",   background=C["row_even"])
        self._tree.tag_configure("odd",    background=C["row_odd"])
        self._tree.tag_configure("good",   foreground=C["green"])
        self._tree.tag_configure("warn",   foreground=C["yellow"])
        self._tree.tag_configure("bad",    foreground=C["red"])

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self._tree.pack(fill="both", expand=True)

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

    # ── DETAIL PANEL ────────────────────────────────────────────
    def _build_detail_panel(self, parent):
        # Card thông tin SV
        info_card = tk.Frame(parent, bg=C["card"])
        info_card.pack(fill="x", pady=(0, 6))

        # Score ring
        self._ring = ScoreRing(info_card, size=110)
        self._ring.grid(row=0, column=0, rowspan=5, padx=14, pady=10, sticky="nsew")

        # Info fields
        fields = [
            ("👤 Sinh viên", "_lbl_sv"),
            ("🎓 GVHD",      "_lbl_gv"),
            ("🔢 MSSV",      "_lbl_ms"),
            ("📝 Đề tài",    "_lbl_dt"),
        ]
        for fi, (label, attr) in enumerate(fields):
            tk.Label(info_card, text=label, bg=C["card"], fg=C["text3"],
                     font=("Segoe UI", 9)).grid(row=fi, column=1, sticky="w",
                                                padx=(10, 4), pady=2)
            lbl = tk.Label(info_card, text="—", bg=C["card"], fg=C["text"],
                           font=("Segoe UI", 10), anchor="w", wraplength=280)
            lbl.grid(row=fi, column=2, sticky="ew", padx=(0, 10), pady=2)
            setattr(self, attr, lbl)

        info_card.columnconfigure(2, weight=1)

        # Eval row
        self._eval_lbl = tk.Label(info_card, text="", bg=C["card"],
                                  font=("Segoe UI", 11, "bold"))
        self._eval_lbl.grid(row=4, column=1, columnspan=2, sticky="w",
                            padx=(10, 10), pady=(0, 8))

        # Tab: Lỗi / Cảnh báo / Thông tin
        nb_frame = tk.Frame(parent, bg=C["bg"])
        nb_frame.pack(fill="both", expand=True)

        self._nb = ttk.Notebook(nb_frame)
        style = ttk.Style()
        style.configure("TNotebook", background=C["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=C["card"], foreground=C["text2"],
                        padding=[12, 6], font=("Segoe UI", 9, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", C["accent"])],
                  foreground=[("selected", "#FFFFFF")])
        self._nb.pack(fill="both", expand=True)

        self._issue_tabs = {}
        for tab_key, tab_name in [("ERROR", "❌  Lỗi"), ("WARNING", "⚠  Cảnh báo"), ("INFO", "ℹ  Thông tin")]:
            frame = tk.Frame(self._nb, bg=C["bg"])
            self._nb.add(frame, text=tab_name)
            self._issue_tabs[tab_key] = self._build_issue_list(frame)

    def _build_issue_list(self, parent):
        """Tạo scrollable list cho một tab."""
        canvas = tk.Canvas(parent, bg=C["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)

        inner = tk.Frame(canvas, bg=C["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win_id, width=canvas.winfo_width())

        inner.bind("<Configure>", _resize)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        canvas.bind("<Button-4>",  lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>",  lambda e: canvas.yview_scroll(1, "units"))

        return inner   # Trả về frame chứa các issue card

    # ── ACTIONS ─────────────────────────────────────────────────
    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn file KLTN (.docx)",
            filetypes=[("Word Document", "*.docx"), ("Tất cả", "*.*")]
        )
        if files:
            self._pending = list(files)
            short = Path(files[0]).name
            suffix = f" + {len(files)-1} file khác" if len(files) > 1 else ""
            self._path_var.set(f"📄 {short}{suffix}")
            self._set_status(f"Đã chọn {len(files)} file. Nhấn 'Bắt đầu kiểm tra'.")

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa KLTN")
        if folder:
            self._pending = [folder]
            docx_count = len(list(Path(folder).rglob("*.docx")))
            self._path_var.set(f"📂 {Path(folder).name}  ({docx_count} file .docx)")
            self._set_status(f"Thư mục: {folder}  —  {docx_count} file .docx. Nhấn 'Bắt đầu kiểm tra'.")

    def _run_check(self):
        if self._running:
            self._stop_event.set()
            self._set_status("⏳ Đang dừng lại...")
            return

        if not hasattr(self, '_pending') or not self._pending:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn file hoặc thư mục trước!")
            return
            
        self._running = True
        self._stop_event.clear()
        self._run_btn.config(text="⏹  Hủy / Dừng", bg=C["red"], activebackground=C["red"])
        self._pick_btn.config(state="disabled")
        self._folder_btn.config(state="disabled")
        self._export_btn.config(state="disabled")
        
        self._set_status("⏳ Đang kiểm tra...", progress=0)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        results = []
        targets = self._pending

        # Xác định danh sách file
        all_files = []
        for t in targets:
            p = Path(t)
            if p.is_dir():
                all_files += sorted([f for f in p.rglob("*.docx") if not f.name.startswith("~$")])
            elif p.suffix.lower() == ".docx":
                all_files.append(p)

        total = len(all_files)
        if total == 0:
            self.after(0, self._finish, [])
            return

        for i, fp in enumerate(all_files, 1):
            if self._stop_event.is_set():
                break
            
            # Cập nhật UI trước khi xử lý
            pct = int((i - 1) / total * 100)
            self.after(0, self._update_progress, i, total, fp.name, pct, "running")

            try:
                result = check_file(str(fp))
            except Exception as e:
                from check_format_kltn import Issue
                result = CheckResult(str(fp), 0)
                result.issues.append(Issue("System", "ERROR", f"Lỗi đọc file: {str(e)}", ""))
            
            results.append(result)

            # Hiển thị kết quả ngay sau khi xong 1 file
            pct_done = int(i / total * 100)
            self.after(0, self._stream_result, result, i, total, pct_done)

        self.after(0, self._finish, results)

    def _update_progress(self, current, total, filename, pct, state):
        """Cập nhật status bar: spinner + tên file + counter + progressbar."""
        self._file_name_var.set(f"{filename[:70]}")
        self._counter_var.set(f"{current}/{total}")
        self._status_var.set(f"Đang đọc và kiểm tra...")
        self._progress["value"] = pct
        if state == "running" and not self._spinner_job:
            self._start_spinner()

    def _start_spinner(self):
        """Bắt đầu animation spinner."""
        def _tick():
            if not self._running:
                self._file_icon_var.set("")
                self._spinner_job = None
                return
            self._file_icon_var.set(self._spinner_frames[self._spinner_idx % len(self._spinner_frames)])
            self._spinner_idx += 1
            self._spinner_job = self.after(80, _tick)
        self._spinner_job = self.after(0, _tick)

    def _stream_result(self, result, current, total, pct):
        """Hiển thị kết quả ngay sau khi xong từng file (real-time)."""
        score = result.score
        if score >= 90:   icon = "✅"
        elif score >= 70: icon = "✔"
        elif score >= 50: icon = "⚠️"
        else:             icon = "❌"

        self._file_icon_var.set(icon)
        self._status_var.set(
            f"{icon} {result.error_count} lỗi  ·  {result.warn_count} cảnh báo  ·  {score}/100 điểm"
        )
        self._progress["value"] = pct

        # Thêm vào bảng ngay (không đợi hết)
        self._add_tree_row(current - 1, result)
        self._count_lbl.config(text=f"{current}/{total}")

    def _add_tree_row(self, idx, res):
        """Thêm hoặc cập nhật 1 hàng trong Treeview."""
        errors = res.error_count
        warns  = res.warn_count
        score  = res.score

        if score >= 90:   ev, tag = "✅ Đạt tốt", "good"
        elif score >= 70: ev, tag = "✔ Đạt",    "good"
        elif score >= 50: ev, tag = "⚠ Cần sửa", "warn"
        else:             ev, tag = "❌ Không đạt", "bad"

        row_tag = "even" if idx % 2 == 0 else "odd"
        iid = str(idx)

        values = (
            Path(res.filepath).name[:30],
            res.student_name[:20] or "—",
            res.student_id or "—",
            res.advisor[:18] or "—",
            errors or "",
            warns  or "",
            score,
            res.letter_grade,
            ev,
        )

        if self._tree.exists(iid):
            self._tree.item(iid, values=values, tags=(row_tag, tag))
        else:
            self._tree.insert("", "end", iid=iid, values=values, tags=(row_tag, tag))
            # Tự scroll xuống dòng mới nhất
            self._tree.see(iid)

    def _finish(self, results):
        self._results = results
        self._running = False
        
        # Phục hồi UI
        self._run_btn.config(text="▶  Bắt đầu kiểm tra", bg="#2E6D45", activebackground="#3A8F5C")
        self._pick_btn.config(state="normal")
        self._folder_btn.config(state="normal")
        self._export_btn.config(state="normal")

        # Dừng spinner
        if self._spinner_job:
            self.after_cancel(self._spinner_job)
            self._spinner_job = None

        # Cập nhật summary cards
        self._update_summary()

        self._progress["value"] = 100
        self._count_lbl.config(text=f"{len(results)} file")

    def _refresh_tree(self):
        """Xóa và vẽ lại toàn bộ bảng (dùng sau khi sort hoặc load lại)."""
        self._tree.delete(*self._tree.get_children())
        self._update_summary()
        for i, res in enumerate(self._results):
            self._add_tree_row(i, res)
        self._count_lbl.config(text=f"{len(self._results)} file")

    def _update_summary(self):
        for w in self._sum_frame.winfo_children():
            w.destroy()
        if not self._results:
            return

        total  = len(self._results)
        passed = sum(1 for r in self._results if r.score >= 70)
        errors = sum(r.error_count for r in self._results)

        for val, label, color in [
            (total,          "Tổng",         C["text2"]),
            (passed,         "Đạt",          C["green"]),
            (total - passed, "Không đạt",    C["red"]),
            (errors,         "Tổng lỗi",     C["yellow"]),
        ]:
            card = tk.Frame(self._sum_frame, bg=C["card"], padx=14, pady=6)
            card.pack(side="left", padx=(0, 4), pady=2)
            tk.Label(card, text=str(val), bg=C["card"], fg=color,
                     font=("Segoe UI", 16, "bold")).pack()
            tk.Label(card, text=label, bg=C["card"], fg=C["text3"],
                     font=("Segoe UI", 8)).pack()

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx >= len(self._results):
            return
        self._selected_idx = idx
        res = self._results[idx]
        self._show_detail(res)

    def _show_detail(self, res: CheckResult):
        # Score ring
        self._ring.set_score(res.score)

        # Info labels
        self._lbl_sv.config(text=res.student_name or "—")
        self._lbl_gv.config(text=res.advisor or "—")
        self._lbl_ms.config(text=res.student_id or "—")
        self._lbl_dt.config(text=(res.title[:80] + "…" if len(res.title) > 80 else res.title) or "—")

        score = res.score
        if score >= 90:   et, ec = "✅ Đạt tốt",    C["green"]
        elif score >= 70: et, ec = "✔ Đạt",         C["green"]
        elif score >= 50: et, ec = "⚠ Cần sửa",    C["yellow"]
        else:             et, ec = "❌ Không đạt",  C["red"]
        self._eval_lbl.config(text=et, fg=ec)

        # Lọc issues theo severity
        for sev_key, inner in self._issue_tabs.items():
            for w in inner.winfo_children():
                w.destroy()
            filtered = [i for i in res.issues if i.severity == sev_key]
            if not filtered:
                tk.Label(inner, text="Không có mục nào.", bg=C["bg"],
                         fg=C["text3"], font=("Segoe UI", 9), pady=12).pack()
            for iss in filtered:
                self._add_issue_card(inner, iss)

    def _add_issue_card(self, parent, iss):
        color = SEV_COLOR.get(iss.severity, C["text"])
        bg    = SEV_BG.get(iss.severity, C["card"])

        card = tk.Frame(parent, bg=bg, pady=8, padx=10)
        card.pack(fill="x", padx=4, pady=(0, 4))

        # Header
        top = tk.Frame(card, bg=bg)
        top.pack(fill="x")
        tk.Label(top, text=iss.category, bg=bg, fg=color,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        if iss.location:
            tk.Label(top, text=f"  [{iss.location[:40]}]", bg=bg, fg=C["text3"],
                     font=("Segoe UI", 8, "italic")).pack(side="left")

        # Message
        tk.Label(card, text=iss.message, bg=bg, fg=C["text"],
                 font=("Segoe UI", 9), anchor="w", justify="left",
                 wraplength=340).pack(fill="x", pady=(3, 0))

        # Suggestion
        if iss.suggestion:
            hint = tk.Frame(card, bg=C["border"], pady=1, padx=8)
            hint.pack(fill="x", pady=(5, 0))
            tk.Label(hint, text="→ " + iss.suggestion, bg=C["border"], fg=C["text2"],
                     font=("Segoe UI", 8), anchor="w", justify="left",
                     wraplength=320).pack(fill="x")

    # ── SORT ────────────────────────────────────────────────────
    def _sort_tree(self, col, reverse):
        col_idx = {
            "filename":0, "student":1, "mssv":2, "advisor":3,
            "errors":4, "warns":5, "score":6, "letter":7, "eval":8
        }[col]
        data = [(self._tree.set(k, col), k) for k in self._tree.get_children("")]
        try:
            data.sort(key=lambda x: int(x[0]) if x[0].lstrip('-').isdigit() else x[0],
                      reverse=reverse)
        except Exception:
            data.sort(reverse=reverse)
        for idx, (_, k) in enumerate(data):
            self._tree.move(k, "", idx)
        self._tree.heading(col, command=lambda: self._sort_tree(col, not reverse))

    # ── EXPORT ──────────────────────────────────────────────────
    def _export(self):
        if not self._results:
            messagebox.showwarning("Chưa có dữ liệu", "Hãy kiểm tra ít nhất một file trước!")
            return
        out = filedialog.asksaveasfilename(
            title="Lưu kết quả Excel",
            defaultextension=".xlsx",
            initialfile=f"KIEM_TRA_KLTN_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out:
            return
        try:
            export_excel(self._results, out)
            self._set_status(f"✅ Đã xuất: {Path(out).name}")
            if messagebox.askyesno("Hoàn tất", f"Đã lưu file:\n{out}\n\nMở file ngay?"):
                if sys.platform == "darwin":
                    os.system(f'open "{out}"')
                elif sys.platform == "win32":
                    os.startfile(out)
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    # ── CLEAR ───────────────────────────────────────────────────
    def _clear(self):
        self._results = []
        self._pending = []
        self._selected_idx = -1
        self._tree.delete(*self._tree.get_children())
        for sev_key, inner in self._issue_tabs.items():
            for w in inner.winfo_children(): w.destroy()
        self._ring.set_score(0)
        for attr in ("_lbl_sv", "_lbl_gv", "_lbl_ms", "_lbl_dt"):
            getattr(self, attr).config(text="—")
        self._eval_lbl.config(text="")
        self._path_var.set("Chưa chọn file / thư mục")
        for w in self._sum_frame.winfo_children(): w.destroy()
        self._count_lbl.config(text="")
        self._file_icon_var.set("")
        self._file_name_var.set("Sẵn sàng.")
        self._counter_var.set("")
        self._status_var.set("")
        self._progress["value"] = 0
        if self._spinner_job:
            self.after_cancel(self._spinner_job)
            self._spinner_job = None

    # ── STATUS (giữ lại cho _run_check) ───────────────────────
    def _set_status(self, msg, progress=None):
        self._file_name_var.set(msg)
        if progress is not None:
            self._progress["value"] = progress


# ════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
