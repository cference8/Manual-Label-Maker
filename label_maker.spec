# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['label_maker.py'],
    pathex=[],
    binaries=[],
    datas=[('Label_Template_BLANK.docx', '.'), ('GENERATED_Label_Template.docx', '.'), ('scribe-logo-final.png', '.'), ('scribe-icon-2.ico', '.')],
    hiddenimports=[],
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
    name='label_maker',
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
    icon=['scribe-icon-2.ico'],
)
