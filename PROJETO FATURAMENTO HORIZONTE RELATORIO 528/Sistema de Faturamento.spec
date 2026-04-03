# -*- mode: python ; coding: utf-8 -*-
import os

pyd_tk = r"C:\Users\ismae\AppData\Local\Programs\Python\Python311\DLLs\_tkinter.pyd"
dll_tcl = r"C:\Users\ismae\AppData\Local\Programs\Python\Python311\DLLs\tcl86t.dll"
dll_tk = r"C:\Users\ismae\AppData\Local\Programs\Python\Python311\DLLs\tk86t.dll"
dir_tcl = r"C:\Users\ismae\AppData\Local\Programs\Python\Python311\tcl\tcl8.6"
dir_tk = r"C:\Users\ismae\AppData\Local\Programs\Python\Python311\tcl\tk8.6"

binaries = []
for src, dst in [
    (pyd_tk, '.'),
    (dll_tcl, '.'),
    (dll_tk, '.'),
]:
    if os.path.exists(src):
        binaries.append((src, dst))

datas = [('logo.png', '.'), ('logo.ico', '.')]
if os.path.exists(dir_tcl):
    datas.append((dir_tcl, 'tcl/tcl8.6'))
if os.path.exists(dir_tk):
    datas.append((dir_tk, 'tcl/tk8.6'))

a = Analysis(
    ['sistema_faturamento.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'openpyxl', 'pandas', 'fitz', 'customtkinter', 'xlrd', 'matplotlib',
        'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox', '_tkinter'
    ],
    hookspath=['pyinstaller_hooks'],
    hooksconfig={},
    runtime_hooks=['pyi_rth_tk_fix.py'],
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
    name='Sistema de Faturamento',
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
    icon=['logo.ico'],
)



