# -*- mode: python ; coding: utf-8 -*-
#
# Windows 7: збирайте на Python 3.8.x (див. build_win7.bat та requirements-win7.txt).
# Офіційні збірки Python 3.9+ не підтримують Win7 як цільову ОС для frozen exe.

a = Analysis(
    ['pdf_rename_expert.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('rules.json', '.'),
        ('INSTRUKTSIYA_KORYSTUVACHA_APP.md', '.'),
    ],
    hiddenimports=[
        "app_ui",
        "naming",
        "excel_db",
        "meters_model",
        "meters_excel",
        "meters_form",
        "win32com.client",
        "pythoncom",
        "file_builder",
        "report",
        "pdf_preview",
        "suffix_history",
        "theme",
        "user_settings",
        "user_manual",
        "version",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter.test",
        "test",
        "unittest",
        "pydoc",
        "distutils.tests",
        # великі пакети з глобального Python, які програмі не потрібні
        "scipy",
        "matplotlib",
        "numba",
        "sklearn",
        "IPython",
        "notebook",
        "pytest",
        "sqlalchemy",
    ],
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
