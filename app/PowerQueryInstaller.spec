# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


a = Analysis(
    ['src/pq_installer/main.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('src/pq_installer/ui/assets/checkmark-dark.svg', 'pq_installer/ui/assets'),
        ('src/pq_installer/ui/assets/checkmark-light.svg', 'pq_installer/ui/assets'),
    ],
    hiddenimports=['pythoncom', 'pywintypes', 'win32com.client'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
cli_a = Analysis(
    ['src/pq_installer/cli.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=['pythoncom', 'pywintypes', 'win32com.client'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PySide6', 'shiboken6'],
    noarchive=False,
    optimize=0,
)
MERGE((a, 'main', 'PowerQueryManager'), (cli_a, 'cli', 'PowerQueryManagerCLI'))

# The app does not install Qt translators and uses English source strings. It
# also does not use Qt Network, QML/Quick, PDF, virtual keyboard, or OpenGL APIs;
# these hook-collected modules, plugins, catalogs, and renderers are unused.
_unused_qt_basenames = {
    'qt6opengl.dll', 'qt6pdf.dll', 'qt6qml.dll', 'qt6qmlmeta.dll',
    'qt6qmlmodels.dll', 'qt6qmlworkerscript.dll', 'qt6quick.dll',
    'qt6virtualkeyboard.dll', 'opengl32sw.dll', 'qpdf.dll',
    'qtvirtualkeyboardplugin.dll', 'qt6network.dll', 'qtnetwork.pyd',
    'qnetworklistmanager.dll', 'qopensslbackend.dll', 'qschannelbackend.dll',
    'qcertonlybackend.dll', 'libcrypto-3-x64.dll', 'libssl-3-x64.dll',
}
a.binaries = [entry for entry in a.binaries
              if Path(str(entry[0]).replace('\\', '/')).name.casefold() not in _unused_qt_basenames]
a.datas = [entry for entry in a.datas
           if Path(str(entry[0]).replace('\\', '/')).name.casefold() not in _unused_qt_basenames]
a.datas = [entry for entry in a.datas
           if not str(entry[0]).replace('\\', '/').casefold().startswith('pyside6/translations/')]

pyz = PYZ(a.pure)
cli_pyz = PYZ(cli_a.pure)

exe = EXE(
    pyz,
    a.dependencies,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PowerQueryManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
cli_exe = EXE(
    cli_pyz,
    cli_a.dependencies,
    cli_a.scripts,
    [],
    exclude_binaries=True,
    name='PowerQueryManagerCLI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    cli_exe,
    a.binaries + cli_a.binaries,
    a.datas + cli_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PowerQueryManager',
)
