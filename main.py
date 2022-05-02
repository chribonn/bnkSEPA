# Payments Business Electronic Banking Services
# SEPA Credit Transfers file layout
# Pain.001.001.03
# Alan Bonnici - chribonn@gmail.com
# Last update: 202205
# version - 1.35.00
# Project repository: https://www.github.com/chribonn/bovSEPA

import argparse
import procZip
import procXlsx
import tempfile
import secrets

# Get the argument of the number to process
parser = argparse.ArgumentParser(description='bank SEPA file processor',
                                 epilog='If upload file doesn\'t have a valid email and password it will be deleted.')
parser.add_argument('--zipname', type=str, help='Enter the ZIP file you wish to process', required=True)
parser.add_argument("--zippath", type=str, help='Directory where the file is located', default=tempfile.gettempdir())
parser.add_argument("--zippass", type=str, help='The password of the zip file', default=secrets.tmp_zippass())
parser.add_argument("--bankSCTE", type=str, help='The password to archive the bank SCTE file', default=secrets.bnk_scte())
args = parser.parse_args()

print('Processing :', args.zippath, "\\", args.zipname)

# The following details will be extracted from the user's database record
# the name of the file to look for in the zip archive
xlname = 'bank SEPA.xlsm'
# this is the archive the zip file will be protected with. Eventually it will be read from a database.
zippass = args.zippass

bankPass = sh['A2'].value
if bankPass is None or bankPass.strip() == '':
    raise Exception('Invalid SCTE archive password')

# Extract the Zip
with tempfile.TemporaryDirectory() as tmpdirname:
    xlsx_filepath = procZip.extractXL(args.zippath, args.zipname, zippass, xlname, tmpdirname)
    procXlsx.procXL(args.zippath, xlsx_filepath, tmpdirname, args.bankSCTE)
    # clean up
    del xlsx_filepath

# clear out all variables before termination
del xlname, zippass