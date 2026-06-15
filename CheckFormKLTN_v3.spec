# -*- mode: python ; coding: utf-8 -*-
"""
CheckFormKLTN_v3.spec
PyInstaller spec cho bản v3 (PyWebview 3D Desktop)
"""
import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['webview_app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('config_kltn.json', '.'),
        ('check_format_kltn.py', '.'),
        ('web_app.py', '.'),
        ('web_static/*', 'web_static'),
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
        'flask',
        'webview',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL', 'scipy', 'cv2', 'tkinter'],
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
    name='CheckFormKLTN_v3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
