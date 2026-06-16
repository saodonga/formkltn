import re
import os

def refactor():
    with open("gui_check_kltn.py", "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Add sv_ttk
    content = content.replace('_ensure("openpyxl")', '_ensure("openpyxl")\n_ensure("sv_ttk")')
    
    # 2. Init theme & v2.0
    content = content.replace(
        'super().__init__()',
        'super().__init__()\n        import sv_ttk\n        sv_ttk.set_theme("light")'
    )
    content = content.replace('CheckForm KLTN — v1.0', 'CheckForm KLTN — v2.0')
    content = content.replace(' v1.0 ', ' v2.0 ')
    
    # 3. Remove RoundButton
    content = re.sub(r'class RoundButton\(tk\.Button\):.*?(?=class ScoreRing\(tk\.Canvas\):)', '', content, flags=re.DOTALL)
    
    # 4. Buttons
    btn_replacements = [
        (
            'self._pick_btn = RoundButton(tb, "📄  Chọn File .docx", command=self._pick_files,\n                    width=180, height=36, bg=C["accent"])',
            'self._pick_btn = ttk.Button(tb, text="📄  Chọn File .docx", command=self._pick_files)'
        ),
        (
            'self._folder_btn = RoundButton(tb, "📂  Chọn Thư mục", command=self._pick_folder,\n                    width=180, height=36, bg=C["acc2"])',
            'self._folder_btn = ttk.Button(tb, text="📂  Chọn Thư mục", command=self._pick_folder)'
        ),
        (
            'self._run_btn = RoundButton(tb, "▶  Bắt đầu kiểm tra", command=self._run_check,\n                                    width=190, height=36, bg="#2E6D45", hover_bg="#3A8F5C")',
            'self._run_btn = ttk.Button(tb, text="▶  Bắt đầu kiểm tra", command=self._run_check, style="Accent.TButton")'
        ),
        (
            'self._export_btn = RoundButton(tb, "💾  Xuất Excel", command=self._export,\n                    width=150, height=36, bg="#D97706", hover_bg="#F59E0B")',
            'self._export_btn = ttk.Button(tb, text="💾  Xuất Excel", command=self._export)'
        ),
        (
            'RoundButton(tb, "👥 Cấu hình GVHD", command=self._show_config_dialog,\n                    width=170, height=36, bg="#4A235A", hover_bg="#6C3483").pack(side="right", padx=(10,0))',
            'ttk.Button(tb, text="👥 Cấu hình GVHD", command=self._show_config_dialog).pack(side="right", padx=(10,0))'
        ),
        (
            'RoundButton(tb, "✕  Xóa", command=self._clear,\n                    width=90, height=36, bg=C["border"], hover_bg="#505580",\n                    fg=C["text2"]).pack(side="right")',
            'ttk.Button(tb, text="✕  Xóa", command=self._clear).pack(side="right")'
        ),
        (
            'RoundButton(btn_frame, "💾 Lưu cấu hình", command=_save, width=150, height=36, bg=C["accent"]).pack(side="right")',
            'ttk.Button(btn_frame, text="💾 Lưu cấu hình", command=_save, style="Accent.TButton").pack(side="right")'
        ),
        (
            'RoundButton(btn_frame, "Hủy", command=dlg.destroy, width=100, height=36, bg=C["border"], hover_bg="#505580", fg=C["text2"]).pack(side="right", padx=(0, 10))',
            'ttk.Button(btn_frame, text="Hủy", command=dlg.destroy).pack(side="right", padx=(0, 10))'
        )
    ]
    for old, new in btn_replacements:
        content = content.replace(old, new)
        
    # 5. Fix tk.Frame -> ttk.Frame, tk.Label -> ttk.Label and strip colors
    # Find all tk.Frame(...) and replace with ttk.Frame(..., clean_args)
    def clean_widget(match):
        widget = match.group(1) # e.g. tk.Frame
        args = match.group(2)
        if widget in ('tk.Frame', 'tk.Label', 'tk.PanedWindow'):
            new_widget = widget.replace('tk.', 'ttk.')
            # remove bg, fg, activebackground
            args = re.sub(r',\s*bg=[^,)]+', '', args)
            args = re.sub(r',\s*fg=[^,)]+', '', args)
            args = re.sub(r',\s*activebackground=[^,)]+', '', args)
            args = re.sub(r',\s*sashwidth=[^,)]+', '', args)
            args = re.sub(r',\s*sashrelief=[^,)]+', '', args)
            args = re.sub(r',\s*sashpad=[^,)]+', '', args)
            return f"{new_widget}({args})"
        return match.group(0)

    content = re.sub(r'(tk\.Frame|tk\.Label|tk\.PanedWindow)\((.*?)\)', clean_widget, content, flags=re.DOTALL)
    
    # 6. Some specific fixes: self.configure(bg=...) in App.__init__
    content = content.replace('self.configure(bg=C["bg"])', '# self.configure(bg=C["bg"])')
    
    # 7. Treeview style updates for sv_ttk
    tree_style_old = '''        style = ttk.Style()
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
                  foreground=[("selected", C["text"])])'''
    tree_style_new = '''        # sv_ttk handles Treeview styles automatically'''
    content = content.replace(tree_style_old, tree_style_new)

    # 8. Fix Score ring (it needs 'bg' for canvas)
    content = content.replace('super().__init__(parent, width=size, height=size,\n                         bg=parent["bg"], highlightthickness=0, **kw)',
                              'super().__init__(parent, width=size, height=size,\n                         highlightthickness=0, **kw)')
    
    # Fix btn_frame bg in _show_config_dialog
    content = content.replace('btn_frame = tk.Frame(dlg, bg=C["bg"])', 'btn_frame = ttk.Frame(dlg)')
    content = content.replace('btn_frame = ttk.Frame(dlg)', 'btn_frame = ttk.Frame(dlg)', 1) # just in case
    
    with open("gui_check_kltn.py", "w", encoding="utf-8") as f:
        f.write(content)
        
if __name__ == "__main__":
    refactor()
    print("Refactoring complete.")
