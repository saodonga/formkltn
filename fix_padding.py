import re

with open("gui_check_kltn.py", "r", encoding="utf-8") as f:
    content = f.read()

# Replace padding kwargs in ttk.Frame
content = re.sub(r'(ttk\.Frame\([^)]*?),\s*pady=[0-9]+', r'\1', content)
content = re.sub(r'(ttk\.Frame\([^)]*?),\s*padx=[0-9]+', r'\1', content)

# Check ttk.Label just in case
content = re.sub(r'(ttk\.Label\([^)]*?),\s*pady=[0-9]+', r'\1', content)
content = re.sub(r'(ttk\.Label\([^)]*?),\s*padx=[0-9]+', r'\1', content)

# Write back
with open("gui_check_kltn.py", "w", encoding="utf-8") as f:
    f.write(content)
