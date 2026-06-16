import re
import os

def fix():
    with open("gui_check_kltn.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Remove bg=..., fg=..., activebackground=... explicitly
    content = re.sub(r',\s*bg=[a-zA-Z0-9_\[\]"\'#]+', '', content)
    content = re.sub(r',\s*fg=[a-zA-Z0-9_\[\]"\'#]+', '', content)
    
    # tk.Canvas is the only one that needs bg back, but it's ScoreRing
    # Wait, if I remove all `bg=`, it will break `tk.Canvas` and `tk.Text` if they need bg
    pass

if __name__ == "__main__":
    fix()
