# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# 修正虚拟环境中 pycel 的路径
if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
    # 我们在虚拟环境中
    if sys.platform == 'win32':
        # Windows 路径
        site_packages = os.path.join(sys.prefix, 'Lib', 'site-packages')
    else:
        # Unix 路径
        site_packages = os.path.join(sys.prefix, 'lib', 'python' + sys.version[:3], 'site-packages')
    pycel_path = os.path.join(site_packages, 'pycel')
else:
    # 我们不在虚拟环境中，使用系统路径
    pycel_path = os.path.dirname(sys.modules['pycel'].__file__)

# 验证路径是否存在
if not os.path.exists(pycel_path):
    raise FileNotFoundError(f"pycel path not found: {pycel_path}")

a = Analysis(['app.py'],
             pathex=['.'],
             binaries=[],
             datas=[
                 ('templates', 'templates'),
                 ('static', 'static'),
                 (pycel_path, 'pycel')  # 包含整个 pycel 目录
             ] + collect_data_files('pycel'),  # 收集 pycel 的所有数据文件
             hiddenimports=[
                 'model',
                 'openpyxl',
                 'pycel',
                 'ast',
                 'datetime',
                 'openpyxl.utils',
                 'openpyxl.styles',
                 'openpyxl.formula.translate',
                 'openpyxl.worksheet.formula',
                 'json',
                 'copy',
                 're',
             ],
             hookspath=['.'],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)


# 确保包含 openpyxl 的数据文件
openpyxl_data = Tree('venv/Lib/site-packages/openpyxl', prefix='openpyxl')
a.datas += openpyxl_data

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='excel_processing_app',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
