# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path

# 添加当前目录到路径
current_dir = os.path.dirname(os.path.abspath(SPEC))
sys.path.insert(0, current_dir)

block_cipher = None

# 明确指定需要的隐藏导入
hiddenimports = [
    'requests',
    'pandas',
    'numpy',
    'openpyxl',
    'PyQt6.QtWidgets',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.sip',
    'beautifulsoup4',
    'bs4',
    'bs4.builder',
    'bs4.builder._lxml',
    'bs4.builder._htmlparser',
    'premailer',
    'premailer.premailer',
    'lxml',
    'lxml.etree',
    'lxml.html',
    'lxml._elementpath',
    'cssutils',
    'cssselect',
    'cssselect.parser',
    'cssselect.xpath',
    'xml.etree.ElementTree',
    'html.parser',
    'urllib3',
    'certifi',
    'charset_normalizer',
    'idna',
]

# 数据文件
datas = []

# 二进制文件
binaries = []

a = Analysis(
    ['wechat_draft_creator.py'],
    pathex=[current_dir],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'tkinter', 'test', 'tests'],
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
    name='微信存稿工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设为False使用窗口模式
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
) 