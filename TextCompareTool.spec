# -*- mode: python ; coding: utf-8 -*-

import os
import sys


conda_bin_dir = os.path.join(sys.base_prefix, "Library", "bin")
required_binaries = [
    (os.path.join(conda_bin_dir, "libbz2.dll"), "."),
    (os.path.join(conda_bin_dir, "libcrypto-3-x64.dll"), "."),
    (os.path.join(conda_bin_dir, "liblzma.dll"), "."),
    (os.path.join(conda_bin_dir, "libmpdec-4.dll"), "."),
    (os.path.join(conda_bin_dir, "tcl86t.dll"), "."),
    (os.path.join(conda_bin_dir, "tk86t.dll"), "."),
]


a = Analysis(
    ['text_compare_tool.py'],
    pathex=[],
    binaries=required_binaries,
    datas=[],
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
    name='TextCompareTool',
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
