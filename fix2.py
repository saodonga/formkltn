import re

with open("gui_check_kltn.py", "r", encoding="utf-8") as f:
    content = f.read()

def clean_ttk(match):
    # match.group(0) is the full match like ttk.Label(..., bg=...)
    widget = match.group(1)
    args = match.group(2)
    # Remove bg and fg
    args = re.sub(r',\s*bg=[a-zA-Z0-9_\[\]"\'#]+', '', args)
    args = re.sub(r',\s*fg=[a-zA-Z0-9_\[\]"\'#]+', '', args)
    return f"{widget}({args})"

# We only target ttk.Label, ttk.Frame, ttk.PanedWindow
content = re.sub(r'(ttk\.[a-zA-Z]+)\((.*?)\)', clean_ttk, content, flags=re.DOTALL)

with open("gui_check_kltn.py", "w", encoding="utf-8") as f:
    f.write(content)
