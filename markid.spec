# PyInstaller spec - build with:  pyinstaller --noconfirm markid.spec
import os, sys
hidden = ['openpyxl', 'webview']
if sys.platform == 'win32': hidden += ['webview.platforms.edgechromium', 'webview.platforms.winforms']
elif sys.platform == 'darwin': hidden += ['webview.platforms.cocoa']
a = Analysis(['desktop.py'], datas=[('schema.sql', '.')], hiddenimports=hidden)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='MarkidPayroll', console=bool(os.environ.get('MARKID_CONSOLE')))
coll = COLLECT(exe, a.binaries, a.datas, name='MarkidPayroll')
if sys.platform == 'darwin':
    app = BUNDLE(coll, name='Markid Payroll.app', bundle_identifier='com.markid.payroll', info_plist={'NSHighResolutionCapable': True})
