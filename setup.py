#!/usr/bin/python
# -*- coding: utf-8 -*-
#setup.py
#Builds webprintfolder using py2exe
#To build, run the command:
# python setup.py py2exe

from distutils.core import setup
import py2exe
import sys
from glob import glob
import os.path

mfcdir = '..\\'
vc90_crt_dir = '..\\'

sys.path.append(vc90_crt_dir)

mfcfiles = [os.path.join(mfcdir, i) for i in ["mfc90.dll",
                "mfc90u.dll", "mfcm90.dll", "mfcm90u.dll",
                "Microsoft.VC90.MFC.manifest"]]

crt_files = [os.path.join(vc90_crt_dir, i) for i in ['Microsoft.VC90.manifest',
'msvcm90.dll', 'msvcp90.dll', 'msvcr90.dll']]

options = {
            'py2exe':{'dist_dir': '..\\dist'},
            'build': {'build_base': '..\\build'}
            }

data_files = [("Microsoft.VC90.CRT", crt_files), ("Microsoft.VC90.MFC", mfcfiles),
                ('', ['..\\cacert.pem'])]

setup(
    windows=[
        {   'script': 'webprintfolder.pyw',
            'icon_resources': [(1, '..\\printer-icon.ico')]
        }
    ],
    data_files=data_files,
    options=options
)