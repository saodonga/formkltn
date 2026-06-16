#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_check_kltn.py — Giao diện Native đồ họa kiểm tra định dạng KLTN
Phiên bản 3.0: CustomTkinter 3D Neon Dark Mode
"""

import os, sys, threading
from pathlib import Path
from datetime import datetime

# ── Tự cài thư viện ──────────────────────────────────────────────
def _ensure(pkg, imp=None):
    try: __import__(imp or pkg)
    except ImportError:
        os.system(f"{sys.executable} -m pip install {pkg} -q")

_ensure("python-docx", "docx")
_ensure("openpyxl")
_ensure("customtkinter")
_ensure("darkdetect")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk

# Import engine từ check_format_kltn.py
sys.path.insert(0, str(Path(__file__).parent))
from check_format_kltn import check_file, scan_directory, export_excel, CheckResult


# ════════════════════════════════════════════════════════════════
#  MÀU SẮC (3D Neon Dark Mode)
# ════════════════════════════════════════════════════════════════
C = {
    "bg":        "#0a0f18",   # Nền app tối sâu
    "panel":     "#111827",   # Nền Panel 
    "card":      "#1f2937",   # Nền khối (sáng hơn panel)
    "border":    "#374151",   # Đường viền
    "accent":    "#3b82f6",   # Xanh Neon
    "accent_h":  "#60a5fa",   # Xanh hover
    "violet":    "#8b5cf6",   # Tím
    "green":     "#10b981",   # Xanh lá OK
    "yellow":    "#f59e0b",   # Vàng cảnh báo
    "red":       "#ef4444",   # Đỏ báo lỗi
    "text":      "#f3f4f6",   # Text chính
    "text2":     "#9ca3af",   # Text phụ
    "text_dark": "#1f2937",   # Text tối
}

SEV_COLOR = {"ERROR": C["red"], "WARNING": C["yellow"], "INFO": C["green"]}


# ════════════════════════════════════════════════════════════════
#  WIDGET TIỆN ÍCH
# ════════════════════════════════════════════════════════════════
class ScoreRing(tk.Canvas):
    """Vòng điểm số dạng arc."""
    def __init__(self, parent, size=120, **kw):
        super().__init__(parent, width=size, height=size,
                         bg=C["panel"], highlightthickness=0, **kw)
        self._sz = size
        self._score = 0

    def set_score(self, score):
        self._score = score
        self._draw()

    def _draw(self):
        sz = self._sz
        self.delete("all")
        pad = 12
        # Track 3D
        self.create_arc(pad, pad, sz-pad, sz-pad, start=90, extent=360,
                        outline=C["border"], width=10, style="arc")
        # Arc
        if self._score > 0:
            color = C["green"] if self._score >= 70 else (C["yellow"] if self._score >= 50 else C["red"])
            ext = int(self._score / 100 * 360)
            self.create_arc(pad, pad, sz-pad, sz-pad, start=90, extent=-ext,
                            outline=color, width=10, style="arc", capstyle=tk.ROUND)
            
            # Glow effect giả lập 3D
            self.create_arc(pad-1, pad-1, sz-pad+1, sz-pad+1, start=90, extent=-ext,
                            outline=color, width=2, style="arc", stipple="gray50")

        # Số
        color = C["green"] if self._score >= 70 else (C["yellow"] if self._score >= 50 else C["red"])
        self.create_text(sz//2, sz//2 - 10, text=f"{self._score}", fill=color,
                         font=("Segoe UI", 24, "bold"))
        self.create_text(sz//2, sz//2 + 18, text="/ 100", fill=C["text2"],
                         font=("Segoe UI", 11, "bold"))


# ════════════════════════════════════════════════════════════════
#  CỬA SỔ CHÍNH
# ════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Cấu hình CustomTkinter Theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        
        self.title("CheckForm KLTN — v3.0 (3D Native)")
        self.geometry("1180x800")
        self.minsize(1024, 700)
        self.configure(fg_color=C["bg"])

        # Dữ liệu
        self._results: list[CheckResult] = []
        self._selected_idx = -1
        self._running = False
        self._stop_event = threading.Event()
        
        self._style_treeview()
        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _style_treeview(self):
        """Custom style cho ttk.Treeview để hòa hợp với Dark Mode CTk"""
        style = ttk.Style()
        style.theme_use("default")
        
        # Style bảng
        style.configure("Treeview", 
                        background=C["card"],
                        foreground=C["text"],
                        rowheight=35,
                        fieldbackground=C["card"],
                        borderwidth=0,
                        font=("Segoe UI", 10))
        style.map('Treeview', background=[('selected', C["accent"])])
        
        # Style Header
        style.configure("Treeview.Heading",
                        background=C["panel"],
                        foreground=C["text2"],
                        relief="flat",
                        borderwidth=0,
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview.Heading", background=[('active', C["border"])])
        
        # Xóa viền ngoài
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

    # ── BUILD UI ────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        hdr = ctk.CTkFrame(self, height=72, fg_color=C["panel"], corner_radius=0)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        ctk.CTkLabel(hdr, text="⚡", font=("Segoe UI", 28), text_color=C["violet"]).pack(side="left", padx=(24,8), pady=12)
        
        title_box = ctk.CTkFrame(hdr, fg_color="transparent")
        title_box.pack(side="left", pady=12)
        ctk.CTkLabel(title_box, text="Kiểm tra Định dạng KLTN", font=("Segoe UI", 18, "bold"), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(title_box, text="Trường ĐH Thủy Lợi · Khoa Kinh tế và Quản lý", font=("Segoe UI", 11), text_color=C["text2"]).pack(anchor="w", pady=(0,0))

        # Version badge
        badge = ctk.CTkLabel(hdr, text=" v3.0 3D Edition ",
                             font=("Segoe UI", 11, "bold"), 
                             fg_color=C["violet"], 
                             text_color="white",
                             corner_radius=12)
        badge.pack(side="right", padx=24)

        # ── Toolbar ──
        tb = ctk.CTkFrame(self, fg_color="transparent")
        tb.pack(fill="x", padx=24, pady=16)

        self._pick_btn = ctk.CTkButton(tb, text="📄 Chọn File .docx", font=("Segoe UI", 13, "bold"), 
                                       fg_color=C["card"], text_color=C["text"], hover_color=C["border"],
                                       border_width=1, border_color=C["border"], command=self._pick_files)
        self._pick_btn.pack(side="left", padx=(0,12))
        
        self._folder_btn = ctk.CTkButton(tb, text="📂 Chọn Thư mục", font=("Segoe UI", 13, "bold"),
                                         fg_color=C["card"], text_color=C["text"], hover_color=C["border"],
                                         border_width=1, border_color=C["border"], command=self._pick_folder)
        self._folder_btn.pack(side="left", padx=(0,12))

        self._run_btn = ctk.CTkButton(tb, text="▶ Bắt đầu kiểm tra", font=("Segoe UI", 14, "bold"),
                                      fg_color=C["accent"], hover_color=C["accent_h"],
                                      command=self._run_check)
        self._run_btn.pack(side="left", padx=(0,12))

        self._rerun_btn = ctk.CTkButton(tb, text="🔄 Chạy lại", font=("Segoe UI", 13, "bold"),
                                        fg_color=C["card"], text_color=C["text"], hover_color=C["border"], width=100,
                                        command=self._run_again)
        self._rerun_btn.pack(side="left", padx=(0,12))

        self._export_btn = ctk.CTkButton(tb, text="💾 Xuất Excel", font=("Segoe UI", 13, "bold"),
                                         fg_color=C["green"], hover_color="#059669", width=100,
                                         command=self._export)
        self._export_btn.pack(side="left", padx=(0,12))

        self._clear_btn = ctk.CTkButton(tb, text="🗑 Xóa", font=("Segoe UI", 13, "bold"),
                                        fg_color=C["card"], text_color=C["red"], hover_color=C["border"], width=80,
                                        command=self._clear)
        self._clear_btn.pack(side="left", padx=(0,12))

        # Config Button
        ctk.CTkButton(tb, text="👥 Cấu hình GVHD", font=("Segoe UI", 13, "bold"),
                      fg_color="transparent", text_color=C["accent"], hover_color=C["card"], width=120,
                      command=self._show_config_dialog).pack(side="right")

        # Label đường dẫn
        self._path_var = ctk.StringVar(value="Chưa chọn file / thư mục")
        path_lbl = ctk.CTkLabel(tb, textvariable=self._path_var,
                                font=("Segoe UI", 12), text_color=C["text2"])
        path_lbl.pack(side="right", padx=16)

        # ── Body (Chia cột) ──
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        # Left — danh sách file
        left = ctk.CTkFrame(body, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 16))
        self._build_file_list(left)

        # Right — chi tiết lỗi
        right = ctk.CTkFrame(body, fg_color="transparent", width=450)
        right.pack(side="right", fill="both")
        right.pack_propagate(False)
        self._build_detail_panel(right)

        # ── Status bar ──
        sb = ctk.CTkFrame(self, height=40, fg_color=C["panel"], corner_radius=0)
        sb.pack(fill="x", side="bottom")

        self._file_name_var = ctk.StringVar(value="Sẵn sàng.")
        self._counter_var = ctk.StringVar(value="")
        
        ctk.CTkLabel(sb, textvariable=self._file_name_var, font=("Segoe UI", 12, "bold"), text_color=C["text"]).pack(side="left", padx=24)
        ctk.CTkLabel(sb, text="Kinh tế số TLU - ver 3.0", font=("Segoe UI", 11, "italic"), text_color=C["text2"]).pack(side="right", padx=16)
        ctk.CTkLabel(sb, textvariable=self._counter_var, font=("Segoe UI", 12, "bold"), text_color=C["accent"]).pack(side="right", padx=(16, 0))
        
        self._progress = ctk.CTkProgressBar(sb, width=300, height=8, progress_color=C["violet"], fg_color=C["border"])
        self._progress.set(0)
        self._progress.pack(side="right", padx=16)

    # ── ĐỔI CẤU HÌNH GVHD ───────────────────────────────────────
    def _show_config_dialog(self):
        import json
        cfg_path = "config_kltn.json"
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except:
            cfg = {"advisors": []}

        dlg = ctk.CTkToplevel(self)
        dlg.title("Cấu hình Cán bộ Hướng dẫn")
        dlg.geometry("500x600")
        dlg.transient(self)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Danh sách Cán bộ Hướng dẫn (Mỗi GV 1 dòng):", font=("Segoe UI", 14, "bold")).pack(pady=(20, 5), padx=20, anchor="w")
        ctk.CTkLabel(dlg, text="Bạn có thể copy và paste từ Excel/Word vào đây.", font=("Segoe UI", 12), text_color=C["text2"]).pack(padx=20, anchor="w", pady=(0, 15))

        text_area = ctk.CTkTextbox(dlg, font=("Segoe UI", 13), fg_color=C["card"], border_color=C["border"], border_width=1)
        text_area.pack(fill="both", expand=True, padx=20, pady=(0, 20))
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

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        ctk.CTkButton(btn_frame, text="💾 Lưu cấu hình", font=("Segoe UI", 13, "bold"), command=_save).pack(side="right")
        ctk.CTkButton(btn_frame, text="Hủy", font=("Segoe UI", 13), fg_color=C["card"], hover_color=C["border"], command=dlg.destroy).pack(side="right", padx=(0, 10))

    # ── FILE LIST ───────────────────────────────────────────────
    def _build_file_list(self, parent):
        # Summary bar
        self._sum_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._sum_frame.pack(fill="x", pady=(0, 16))

        # Wrapper cho Treeview để tạo viền 3D
        tree_wrap = ctk.CTkFrame(parent, fg_color=C["panel"], border_width=1, border_color=C["border"], corner_radius=8)
        tree_wrap.pack(fill="both", expand=True)

        cols = ("filename", "student", "mssv", "advisor", "errors", "warns", "score", "letter", "eval")
        self._tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", selectmode="browse")

        col_conf = {
            "filename": ("Tên file",    200, "w"),
            "student":  ("Sinh viên",   140, "w"),
            "mssv":     ("MSSV",        100, "center"),
            "advisor":  ("GVHD",        140, "w"),
            "errors":   ("Lỗi",         60, "center"),
            "warns":    ("Cảnh báo",    70, "center"),
            "score":    ("Điểm",        60, "center"),
            "letter":   ("Chữ",         50, "center"),
            "eval":     ("Đánh giá",    100, "center"),
        }
        for col, (heading, width, anchor) in col_conf.items():
            self._tree.heading(col, text=heading, command=lambda c=col: self._sort_tree(c, False))
            self._tree.column(col, width=width, anchor=anchor, minwidth=50)

        # Tags màu
        self._tree.tag_configure("good",   foreground=C["green"])
        self._tree.tag_configure("warn",   foreground=C["yellow"])
        self._tree.tag_configure("bad",    foreground=C["red"])

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y", pady=2)
        hsb.pack(side="bottom", fill="x", padx=2)
        self._tree.pack(fill="both", expand=True, padx=2, pady=2)

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

    # ── DETAIL PANEL ────────────────────────────────────────────
    def _build_detail_panel(self, parent):
        # Card thông tin SV (3D nổi)
        info_card = ctk.CTkFrame(parent, fg_color=C["panel"], border_width=1, border_color=C["border"], corner_radius=12)
        info_card.pack(fill="x", pady=(0, 16))

        # Bố cục 2 cột cho info_card
        info_grid = ctk.CTkFrame(info_card, fg_color="transparent")
        info_grid.pack(fill="both", expand=True, padx=16, pady=16)

        # Score ring
        ring_frame = ctk.CTkFrame(info_grid, fg_color="transparent", width=120)
        ring_frame.pack(side="left")
        self._ring = ScoreRing(ring_frame, size=120)
        self._ring.pack()

        # Info fields
        fields_frame = ctk.CTkFrame(info_grid, fg_color="transparent")
        fields_frame.pack(side="left", fill="both", expand=True, padx=(20, 0))

        fields = [
            ("👤 Sinh viên", "_lbl_sv"),
            ("🎓 GVHD",      "_lbl_gv"),
            ("🔢 MSSV",      "_lbl_ms"),
            ("📝 Đề tài",    "_lbl_dt"),
        ]
        
        for i, (label, attr) in enumerate(fields):
            row = ctk.CTkFrame(fields_frame, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=label, font=("Segoe UI", 12, "bold"), text_color=C["text2"], width=80, anchor="w").pack(side="left")
            lbl = ctk.CTkLabel(row, text="—", font=("Segoe UI", 13), text_color=C["text"], anchor="w", justify="left", wraplength=220)
            lbl.pack(side="left", fill="x", expand=True)
            setattr(self, attr, lbl)

        # Eval row
        self._eval_lbl = ctk.CTkLabel(fields_frame, text="", font=("Segoe UI", 16, "bold"))
        self._eval_lbl.pack(anchor="w", pady=(8, 0))

        # Tabs bằng SegmentedButton
        self.tab_var = ctk.StringVar(value="ERROR")
        tab_ctrl = ctk.CTkSegmentedButton(parent, values=["ERROR", "WARNING", "INFO"], 
                                          variable=self.tab_var, command=self._switch_tab,
                                          font=("Segoe UI", 12, "bold"), 
                                          selected_color=C["accent"], selected_hover_color=C["accent_h"])
        tab_ctrl.pack(fill="x", pady=(0, 12))

        # Vùng chứa issue
        self._issue_container = ctk.CTkScrollableFrame(parent, fg_color=C["panel"], border_width=1, border_color=C["border"], corner_radius=12)
        self._issue_container.pack(fill="both", expand=True)
        
        # Lưu trữ tạm danh sách issue để render theo tab
        self._current_issues = []

    def _switch_tab(self, value):
        self._render_issues()

    def _render_issues(self):
        # Xóa cũ
        for w in self._issue_container.winfo_children():
            w.destroy()
            
        sev_key = self.tab_var.get()
        filtered = [i for i in self._current_issues if i.severity == sev_key]
        
        if not filtered:
            ctk.CTkLabel(self._issue_container, text="Không có mục nào.", font=("Segoe UI", 13), text_color=C["text2"]).pack(pady=40)
            return
            
        for iss in filtered:
            self._add_issue_card(self._issue_container, iss)

    def _add_issue_card(self, parent, iss):
        color = SEV_COLOR.get(iss.severity, C["text"])
        
        # Thẻ nổi 3D
        card = ctk.CTkFrame(parent, fg_color=C["card"], border_color=color, border_width=1, corner_radius=8)
        card.pack(fill="x", padx=4, pady=(0, 12))

        # Header
        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=12, pady=(10, 4))
        ctk.CTkLabel(top, text=iss.category, font=("Segoe UI", 13, "bold"), text_color=color).pack(side="left")
        if iss.location:
            ctk.CTkLabel(top, text=f"  [{iss.location[:40]}]", font=("Segoe UI", 11, "italic"), text_color=C["text2"]).pack(side="left")

        # Message
        ctk.CTkLabel(card, text=iss.message, font=("Segoe UI", 13), text_color=C["text"], 
                     anchor="w", justify="left", wraplength=380).pack(fill="x", padx=12, pady=(0, 6))

        # Suggestion
        if iss.suggestion:
            hint = ctk.CTkFrame(card, fg_color=C["bg"], corner_radius=6)
            hint.pack(fill="x", padx=12, pady=(0, 10))
            ctk.CTkLabel(hint, text="→ " + iss.suggestion, font=("Segoe UI", 12), text_color=C["accent_h"], 
                         anchor="w", justify="left", wraplength=360).pack(fill="x", padx=10, pady=8)

    # ── ACTIONS ─────────────────────────────────────────────────
    def _pick_files(self):
        files = filedialog.askopenfilenames(title="Chọn file KLTN (.docx)", filetypes=[("Word Document", "*.docx"), ("Tất cả", "*.*")])
        if files:
            self._pending = list(files)
            short = Path(files[0]).name
            suffix = f" + {len(files)-1} file khác" if len(files) > 1 else ""
            self._path_var.set(f"📄 {short}{suffix}")
            self._file_name_var.set(f"Đã chọn {len(files)} file. Nhấn 'Bắt đầu kiểm tra'.")

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa KLTN")
        if folder:
            self._pending = [folder]
            docx_count = len(list(Path(folder).rglob("*.docx")))
            self._path_var.set(f"📂 {Path(folder).name}  ({docx_count} file .docx)")
            self._file_name_var.set(f"Thư mục: {folder} — {docx_count} file .docx. Nhấn 'Bắt đầu kiểm tra'.")

    def _run_check(self):
        if self._running:
            self._stop_event.set()
            self._file_name_var.set("⏳ Đang dừng lại...")
            return

        if not hasattr(self, '_pending') or not self._pending:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn file hoặc thư mục trước!")
            return
            
        self._running = True
        self._stop_event.clear()
        self._run_btn.configure(text="⏹ Hủy / Dừng", fg_color=C["red"], hover_color="#dc2626")
        self._pick_btn.configure(state="disabled")
        self._folder_btn.configure(state="disabled")
        self._export_btn.configure(state="disabled")
        
        self._file_name_var.set("⏳ Đang chuẩn bị kiểm tra...")
        self._progress.set(0)
        threading.Thread(target=self._worker, daemon=True).start()

    def _run_again(self):
        if not hasattr(self, '_pending') or not self._pending:
            messagebox.showwarning("Chưa có file", "Chưa có danh sách file nào để chạy lại!")
            return
        self._clear_results()
        self._run_check()

    def _worker(self):
        results = []
        targets = self._pending
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
            if self._stop_event.is_set(): break
            
            pct = (i - 1) / total
            self.after(0, self._update_progress, i, total, fp.name, pct)

            try:
                result = check_file(str(fp))
            except Exception as e:
                from check_format_kltn import Issue
                result = CheckResult(str(fp), 0)
                result.issues.append(Issue("System", "ERROR", f"Lỗi đọc file: {str(e)}", ""))
            
            results.append(result)
            pct_done = i / total
            self.after(0, self._stream_result, result, i, total, pct_done)

        self.after(0, self._finish, results)

    def _update_progress(self, current, total, filename, pct):
        self._file_name_var.set(f"Đang kiểm tra: {filename[:70]}")
        self._counter_var.set(f"{current}/{total}")
        self._progress.set(pct)

    def _stream_result(self, result, current, total, pct):
        self._progress.set(pct)
        self._add_tree_row(current - 1, result)

    def _add_tree_row(self, idx, res):
        score = res.score
        if score >= 90:   ev, tag = "✅ Đạt tốt", "good"
        elif score >= 70: ev, tag = "✔ Đạt",    "good"
        elif score >= 50: ev, tag = "⚠ Cần sửa", "warn"
        else:             ev, tag = "❌ Không đạt", "bad"

        iid = str(idx)
        values = (Path(res.filepath).name[:30], res.student_name[:20] or "—", res.student_id or "—", res.advisor[:18] or "—", res.error_count or "", res.warn_count or "", score, res.letter_grade, ev)

        if self._tree.exists(iid):
            self._tree.item(iid, values=values, tags=(tag,))
        else:
            self._tree.insert("", "end", iid=iid, values=values, tags=(tag,))
            self._tree.see(iid)

    def _finish(self, results):
        self._results = results
        self._running = False
        
        self._run_btn.configure(text="▶ Bắt đầu kiểm tra", fg_color=C["accent"], hover_color=C["accent_h"])
        self._pick_btn.configure(state="normal")
        self._folder_btn.configure(state="normal")
        self._export_btn.configure(state="normal")

        self._update_summary()
        self._progress.set(1.0)
        self._file_name_var.set(f"Hoàn thành kiểm tra {len(results)} file.")

    def _update_summary(self):
        for w in self._sum_frame.winfo_children(): w.destroy()
        if not self._results: return

        total  = len(self._results)
        passed = sum(1 for r in self._results if r.score >= 70)
        errors = sum(r.error_count for r in self._results)

        for val, label, color in [(total, "Tổng", C["accent_h"]), (passed, "Đạt", C["green"]), (total - passed, "Không đạt", C["red"]), (errors, "Tổng lỗi", C["yellow"])]:
            card = ctk.CTkFrame(self._sum_frame, fg_color=C["panel"], corner_radius=8, border_width=1, border_color=C["border"])
            card.pack(side="left", padx=(0, 12), pady=2, fill="x", expand=True)
            ctk.CTkLabel(card, text=str(val), text_color=color, font=("Segoe UI", 24, "bold")).pack(pady=(12, 0))
            ctk.CTkLabel(card, text=label, text_color=C["text2"], font=("Segoe UI", 12)).pack(pady=(0, 12))

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel: return
        idx = int(sel[0])
        if idx >= len(self._results): return
        self._selected_idx = idx
        self._show_detail(self._results[idx])

    def _show_detail(self, res: CheckResult):
        self._ring.set_score(res.score)

        self._lbl_sv.configure(text=res.student_name or "—")
        self._lbl_gv.configure(text=res.advisor or "—")
        self._lbl_ms.configure(text=res.student_id or "—")
        self._lbl_dt.configure(text=(res.title[:80] + "…" if len(res.title) > 80 else res.title) or "—")

        score = res.score
        if score >= 90:   et, ec = "✅ Đạt tốt", C["green"]
        elif score >= 70: et, ec = "✔ Đạt", C["green"]
        elif score >= 50: et, ec = "⚠ Cần sửa", C["yellow"]
        else:             et, ec = "❌ Không đạt", C["red"]
        self._eval_lbl.configure(text=et, text_color=ec)

        self._current_issues = res.issues
        self._render_issues()

    def _sort_tree(self, col, reverse):
        col_idx = {"filename":0, "student":1, "mssv":2, "advisor":3, "errors":4, "warns":5, "score":6, "letter":7, "eval":8}[col]
        data = [(self._tree.set(k, col), k) for k in self._tree.get_children("")]
        try: data.sort(key=lambda x: int(x[0]) if x[0].lstrip('-').isdigit() else x[0], reverse=reverse)
        except Exception: data.sort(reverse=reverse)
        for idx, (_, k) in enumerate(data): self._tree.move(k, "", idx)
        self._tree.heading(col, command=lambda: self._sort_tree(col, not reverse))

    def _export(self):
        if not self._results:
            messagebox.showwarning("Chưa có dữ liệu", "Hãy kiểm tra ít nhất một file trước!")
            return
        out = filedialog.asksaveasfilename(title="Lưu kết quả Excel", defaultextension=".xlsx", initialfile=f"KIEM_TRA_KLTN_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
        if not out: return
        try:
            export_excel(self._results, out)
            self._file_name_var.set(f"✅ Đã xuất: {Path(out).name}")
            if messagebox.askyesno("Hoàn tất", f"Đã lưu file:\n{out}\n\nMở file ngay?"):
                os.system(f'open "{out}"') if sys.platform == "darwin" else os.startfile(out)
        except Exception as e: messagebox.showerror("Lỗi", str(e))

    def _clear_results(self):
        self._results = []
        self._selected_idx = -1
        self._tree.delete(*self._tree.get_children())
        self._ring.set_score(0)
        for attr in ("_lbl_sv", "_lbl_gv", "_lbl_ms", "_lbl_dt"): getattr(self, attr).configure(text="—")
        self._eval_lbl.configure(text="")
        for w in self._sum_frame.winfo_children(): w.destroy()
        for w in self._issue_container.winfo_children(): w.destroy()
        self._progress.set(0)
        self._file_name_var.set("Sẵn sàng.")
        self._counter_var.set("")

    def _clear(self):
        self._clear_results()
        self._pending = []
        self._path_var.set("Chưa chọn file / thư mục")

if __name__ == "__main__":
    app = App()
    app.mainloop()
