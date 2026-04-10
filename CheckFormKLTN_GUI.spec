# -*- mode: python ; coding: utf-8 -*-
"""
CheckFormKLTN_GUI.spec
PyInstaller spec cho bản GUI (Tkinter desktop)
"""
import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['gui_check_kltn.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Đóng gói file config và engine cùng exe
        ('config_kltn.json', '.'),
        ('check_format_kltn.py', '.'),
    ],
    hiddenimports=[
        'check_format_kltn',
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.enum.text',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL', 'scipy', 'cv2'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CheckFormKLTN',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # Không hiện cửa sổ CMD (windowed mode)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',    # Bỏ comment nếu có file icon
)
