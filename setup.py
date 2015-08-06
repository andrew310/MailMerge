__author__ = 'Andrew.Brown'

from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

setup(
    windows = [{'script': "mailmonstrrr2.py"}],
    zipfile = None,
    options = {
        'py2exe': {
                'bundle_files': 1,
                'dll_excludes': ["MSVCP90.dll"],
                }
        }
)
