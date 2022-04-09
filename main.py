# Payments Business Electronic Banking Services
# SEPA Credit Transfers file layout
# Pain.001.001.03
# Alan Bonnici - chribonn@gmail.com
# Last update: 202204
# version - 1.30.00

import argparse
import procZip
import procXlsx
import secrets
import tempfile

# Get the argument of the number to process
parser = argparse.ArgumentParser(description='BOV SEPA file processor',
                                 epilog='If upload file doesn\'t have a valid email and password it will be deleted.')
parser.add_argument('--zipname', type=str, help='Enter the ZIP file you wish to process', required=True)
parser.add_argument("--zippath", type=str, help='Directory where the file is located', default=tempfile.gettempdir())
args = parser.parse_args()

print('Processing :', args.zippath, "\\", args.zipname)

# The following details will be extracted from the user's database record
# the name of the file to look for in the zip archive
xlname = 'BOV SEPA.xlsm'
# this is the archive the zip file will be protected with. Eventually it will be read from a database.
zippass = secrets.tmp_zippass()

# Extract the Zip
with tempfile.TemporaryDirectory() as tmpdirname:
    xlsx_filepath = procZip.extractXL(args.zippath, args.zipname, zippass, xlname, tmpdirname)
    procXlsx.procXL(args.zippath, xlsx_filepath, tmpdirname)
    # clean up
    del xlsx_filepath

# clear out all variables before termination
del xlname, zippass


