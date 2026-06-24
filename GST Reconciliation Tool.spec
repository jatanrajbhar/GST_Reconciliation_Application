# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\tech\\Documents\\backup\\product2\\gst_reconciliation_app.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\tech\\Documents\\backup\\product2\\logo small.png', '.'), ('C:\\Users\\tech\\Documents\\backup\\product2\\Template.xlsx', '.'), ('C:\\Users\\tech\\Documents\\backup\\product2\\Template all.xlsx', '.'), ('C:\\Users\\tech\\Documents\\backup\\product2\\Privacy Policy.docx', '.'), ('C:\\Users\\tech\\Documents\\backup\\product2\\Terms and Conditions.docx', '.')],
    hiddenimports=['customtkinter', 'PIL._tkinter_finder', 'PIL.Image', 'PIL.ImageTk', 'openpyxl.styles', 'openpyxl.cell'],
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
    [],
    exclude_binaries=True,
    name='GST Reconciliation Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\tech\\Documents\\backup\\product2\\app_icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GST Reconciliation Tool',
)
