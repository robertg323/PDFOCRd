# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas = []
binaries = []
hiddenimports = []

# If you place an extracted Poppler tree under ./poppler, include it in the bundle.
# This will copy the folder into the frozen app under a top-level `poppler/` path
# so runtime detection in `pdf_ocr_check.py` can find `poppler/Library/bin/pdftoppm.exe`.
spec_dir = os.getcwd()  # Current working directory (where pyinstaller is run from)
poppler_src = os.path.join(spec_dir, 'poppler')

if os.path.isdir(poppler_src):
    file_count = 0
    for root, dirs, files in os.walk(poppler_src):
        for fname in files:
            src_path = os.path.join(root, fname)
            rel_dir = os.path.relpath(root, poppler_src)
            if rel_dir == '.':
                dest_dir = 'poppler'
            else:
                dest_dir = os.path.join('poppler', rel_dir)
            datas.append((src_path, dest_dir))
            file_count += 1

# customtkinter (themes, images, etc.)
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Pillow (PIL) — needed for ImageTk preview rendering
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pdf2image
tmp_ret = collect_all('pdf2image')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['pdf_ocr_check.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + [
        'PIL._tkinter_finder',
        'fitz',
        'pytesseract',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PDFOCRd',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
