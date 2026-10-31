# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成单目录（onedir）的 Windows 图形程序。

用法（通过 uv 运行，依赖已锁定）：
    uv run pyinstaller fee_app.spec --noconfirm --clean

产物：dist/FeeMan/FeeMan.exe（连同依赖目录，供 Inno Setup 打包安装程序）
"""

import os

block_cipher = None

# 项目根目录（spec 文件所在目录）
ROOT = os.path.abspath(os.getcwd())

# 剔除 tcl/tk 里用不到的大文件（时区库约 700 个文件、若干示例图片），显著减小体积
_UNUSED_TCL = ("tzdata", "encoding", "msgs", "images", "demos", "http1.0")


def _keep(entry):
    dest = entry[0].replace("\\", "/").lower()
    return not any(part in dest for part in _UNUSED_TCL)


a = Analysis(
    ["fee_app.py"],
    pathex=[ROOT],
    binaries=[],
    datas=[],
    # tkinter 一般能自动收集；openpyxl 的 et_xmlfile 需要显式声明隐式导入
    hiddenimports=[
        "openpyxl",
        "openpyxl.workbook",
        "openpyxl.styles",
        "openpyxl.utils",
        "et_xmlfile",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 明确排除用不到的重量级模块，压缩体积
        "numpy",
        "pandas",
        "matplotlib",
        "scipy",
        "PIL",
        "pytest",
        "unittest",
        "pydoc",
        "doctest",
        "email",
        "html",
        "http",
        "xmlrpc",
        "pdb",
        "tests",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 过滤掉 tcl/tk 中无用的数据文件
a.datas = [d for d in a.datas if _keep(d)]
a.zipfiles = [d for d in a.zipfiles if _keep(d)]
a.binaries = [d for d in a.binaries if _keep(d)]

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FeeMan",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # 图形程序，不弹控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/app.ico",
    version=None,           # 有版本信息资源时改成 "version_info.txt"
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="FeeMan",
)
