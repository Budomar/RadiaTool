# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
import os

block_cipher = None

# Список файлов данных
data_files = [
    ('Матрица.xlsx', '.'), 
    ('Прайс-лист.xlsx', '.'), 
    ('Расчет мощностей METEOR.xlsx', '.'), 
    ('Формуляр для регистрации проектов.xlsm', '.'), 
    ('favicon.ico', '.'),
    ('Lagar.png', '.'),
    ('inst.pdf', '.'),
    # Добавляем изображения для подсказок
    ('1.png', '.'),
    ('2.png', '.'),
    ('3.png', '.')
]

# Добавляем дополнительные файлы из папки, если они существуют
additional_files = [
    'icon.ico',
    'Lagar.png'
]

for file in additional_files:
    if os.path.exists(file):
        data_files.append((file, '.'))

a = Analysis(
    ['start_v8.8.py'],  
    pathex=[os.getcwd()],  
    binaries=[],
    datas=data_files,
    hiddenimports=[
        'pandas', 
        'openpyxl', 
        'tkinter', 
        'pyperclip',
        'numpy',  # pandas зависит от numpy
        'xlrd',   # для чтения старых .xls файлов
        'xlwt',   # для записи .xls файлов
        'pyexcel',
        'pyexcel_xls',
        'pyexcel_xlsx'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy'],  # Исключаем ненужные большие библиотеки
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
    name='RadiaTool_1.9',  
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Использовать UPX для сжатия (должен быть установлен)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Не показывать консольное окно
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join('favicon.ico'),  # Путь к иконке
)