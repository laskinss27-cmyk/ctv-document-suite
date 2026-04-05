# -*- mode: python ; coding: utf-8 -*-
# Файл сборки PyInstaller для CTV Document Suite
# Запуск: pyinstaller CTV_Suite.spec

import os
import sys
from pathlib import Path

block_cipher = None

# Динамический путь к шрифтам Windows
_fonts_dir = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')

a = Analysis(
    ['app.py'],
    pathex=[str(Path('app.py').parent)],
    binaries=[],
    datas=[
        # Шрифты Arial — стандартные шрифты Windows с поддержкой кириллицы
        (os.path.join(_fonts_dir, 'arial.ttf'),   'fonts'),
        (os.path.join(_fonts_dir, 'arialbd.ttf'), 'fonts'),
        (os.path.join(_fonts_dir, 'ariali.ttf'),  'fonts'),
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'reportlab.graphics.barcode',
        'reportlab.graphics.charts',
        'openpyxl.cell._writer',
        'pandas._libs.tslibs.timedeltas',
    ],
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
    name='CTV_Document_Suite',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,             # отключено — UPX вызывает ложные срабатывания антивируса
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # без консоли (GUI-приложение)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',       # раскомментируйте если есть иконка
)
