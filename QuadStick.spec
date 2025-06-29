# -*- mode: python -*-
import os
import shutil

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Copy locales
locales_src = "./locales/"
locales_dest = "./dist/locales/"

if not os.path.exists(locales_dest):
    os.makedirs(locales_dest)

for locale in os.listdir(os.path.join(locales_src)):
    if locale.endswith('.ini'):
        shutil.copyfile(locales_src + locale, locales_dest + locale)

# Copy themes
themes_src = "./themes/"
themes_dest = "./dist/themes/"

if not os.path.exists(themes_dest):
    os.makedirs(themes_dest)

for theme in os.listdir(os.path.join(themes_src)):
    if theme.endswith('.json'):
        shutil.copyfile(themes_src + theme, themes_dest + theme)


a = Analysis(['QuadStick.py'],
             pathex=['D:\\Workspace\\QuadStick Configurator'],
             binaries=None,
             datas=collect_data_files("customtkinter"),
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None,
             excludes=None,
             win_no_prefer_redirects=None,
             win_private_assemblies=None,
             cipher=block_cipher)

a.datas += [('ViGEmClient.dll', 'ViGEmClient.dll', 'DATA')]
a.datas += [('quadstickx.ico', 'quadstickx.ico', 'DATA')]

pyz = PYZ(a.pure,
          a.zipped_data,
          cipher=block_cipher)

# Create exe file
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='QuadStick Configurator',
          debug=False,
          strip=False,
          upx=True,
          console=False,
          icon='quadstickx.ico')
