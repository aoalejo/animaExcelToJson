import py2exe

import sys
sys.setrecursionlimit(5000)

py2exe.freeze(windows=[
    "excelToJsonUI.py",
],
    options = {'py2exe': {'bundle_files': 1, 'compressed': True}},
)