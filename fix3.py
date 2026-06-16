with open("gui_check_kltn.py", "r", encoding="utf-8") as f:
    content = f.read()

# Fix Line 229
content = content.replace('lbl = ttk.Label(dlg, text="Danh sách Cán bộ Hướng dẫn (Mỗi GV 1 dòng):", bg=C["bg"], fg=C["text"], font=("Segoe UI", 11, "bold"))',
                          'lbl = ttk.Label(dlg, text="Danh sách Cán bộ Hướng dẫn (Mỗi GV 1 dòng):", font=("Segoe UI", 11, "bold"))')

# Fix run_btn configs
content = content.replace('self._run_btn.config(text="⏹  Hủy / Dừng", bg=C["red"], activebackground=C["red"])',
                          'self._run_btn.config(text="⏹  Hủy / Dừng")')
content = content.replace('self._run_btn.config(text="▶  Bắt đầu kiểm tra", bg="#2E6D45", activebackground="#3A8F5C")',
                          'self._run_btn.config(text="▶  Bắt đầu kiểm tra")')

# Fix Card label
content = content.replace('ttk.Label(card, text=str(val), bg=C["card"], fg=color,',
                          'ttk.Label(card, text=str(val), foreground=color,')

# Fix other tk.Label left
content = content.replace('ttk.Label(card, text=label, bg=C["card"], fg=C["text3"],',
                          'ttk.Label(card, text=label, foreground=C["text3"],')

# Check if there are any ttk.Label with bg= left
import re
content = re.sub(r'ttk\.Label\(([^)]*),\s*bg=[^,)]+', r'ttk.Label(\1', content)
content = re.sub(r'ttk\.Label\(([^)]*),\s*fg=', r'ttk.Label(\1, foreground=', content)

with open("gui_check_kltn.py", "w", encoding="utf-8") as f:
    f.write(content)
