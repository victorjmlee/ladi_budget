# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec – LADI 예산서 추출 도구."""

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("index.html", "."),
        ("parser.py", "."),
        ("pdf_extract.py", "."),
        ("excel_build.py", "."),
    ],
    hiddenimports=[
        "flask",
        "openpyxl",
        "fitz",
        "fitz.fitz",
        "pymupdf",
        "parser",
        "pdf_extract",
        "excel_build",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "scipy", "PIL"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="LADI예산서",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,           # 콘솔 표시 (에러 확인용, 안정화 후 False로 변경)
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=None,              # 아이콘 추가 시: icon="ladi.ico"
)
