# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pdf_rename_expert.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "app_ui",
        "naming",
        "excel_db",
        "file_builder",
        "report",
        "pdf_preview",
        "suffix_history",
        "theme",
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
    name='PDF_Rename_Expert',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX часто конфліктує з антивірусами та старими ОС — для Win7 безпечніше вимкнути.
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
