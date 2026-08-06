# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Doc Metadata Remover.

Build a single-file standalone executable with:

    pyinstaller build_exe.spec

The resulting binary is written to the ``dist/`` folder:
    * Windows : dist/DocMetadataRemover.exe
    * macOS   : dist/DocMetadataRemover
    * Linux   : dist/DocMetadataRemover
"""

block_cipher = None

# Collect the optional tkinterdnd2 data files (its Tk/tcl extensions) if the
# package is installed, so drag-and-drop keeps working inside the frozen app.
datas = []
hiddenimports = []
try:
    from PyInstaller.utils.hooks import collect_data_files
    datas += collect_data_files("tkinterdnd2")
    hiddenimports += ["tkinterdnd2"]
except Exception:
    pass

# These optional packages are imported lazily (inside functions), so PyInstaller's
# static analysis can miss them. List them explicitly so PDF (.pdf) and legacy
# Office (.doc/.ppt/.xls) support is bundled into the frozen app when installed.
for _opt in ("olefile", "pikepdf"):
    try:
        __import__(_opt)
        hiddenimports.append(_opt)
    except Exception:
        pass


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='DocMetadataRemover',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # windowed GUI app (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
