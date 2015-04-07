"""
@python version:
    Python 3.4

@summary:
    Script to build an exe file

@note:
	Based on http://cx-freeze.readthedocs.org/en/latest/overview.html
    Usage - run in cmd 'python setup.py build'

@author:
    Venancio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (01/09/2014):
        - Creation of the script

    1.1 (09/12/2014):
        - Modification of the python version from 3.3 to 3.4.
        - Diferent tlc dlls

@since:
    01/09/2014
"""

import sys
from cx_Freeze import setup, Executable

# Run in cmd: python setup.py build

base = None
if sys.platform == "win32":
    base = "Win32GUI"

# python 3.3
#buildOptions = dict(include_files = ['fotos/', 'icons/', 'docs/', 'tcl85.dll', 'tk85.dll', 'erros.log'])

# python 3.4
buildOptions = dict(include_files = ['fotos/', 'icons/', 'docs/', 'tcl86t.dll', 'tk86t.dll', 'erros.log'])

setup(  name = "XLS to KMZ Transformer",
        version = "1.6",
        description = "XLS to KMZ GUI application!",
        author = "Nuno Venâncio",
        options = dict(build_exe = buildOptions),
        executables = [Executable("xls2kmz.py", base=base, icon="kmz5.ico")])
