from distutils.core import setup
import py2exe
import argparse
import procZip
import procXlsx
import secrets
import tempfile
import os
import zipfile
import openpyxl
import datetime
import os
import pyminizip
import lxml.etree as etree

setup(
    console=['main.py'],
options={
    'py2exe':
        {
            'includes': ['lxml._elementpath'],
        }
})