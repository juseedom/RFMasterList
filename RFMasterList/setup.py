# setup.py
from distutils.core import setup
import py2exe
import sys, os
sys.path.append(os.path.abspath(os.curdir))
import coloreditorfactory, RFKml, RFDB, Logging

if len(sys.argv) == 1:
	sys.argv.append("py2exe")


setup( options = {"py2exe": {
							"dll_excludes": ["MSVCP90.dll", "HID.DLL", "w9xpopen.exe"],
							"compressed": 1, 
							"optimize": 2, 
							"bundle_files": 3,
							"packages":["xlrd","RFKml", "RFDB", "coloreditorfactory", "Logging"]
							}},
       zipfile = None,

       ## data_files = ['apple.jpg', 'cheese.jpg'],

       #Your py-file can use windows or console
       windows = [{"script": 'RFUI.py'}])
