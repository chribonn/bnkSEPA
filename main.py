# Payments Business Electronic Banking Services
# SEPA Credit Transfers file layout
# Pain.001.001.03
# Alan Bonnici - chribonn@gmail.com
# Last update: 202205
# version - 1.35.00
# Project repository: https://www.github.com/chribonn/bovSEPA

import argparse
import procXlsx
import tempfile
import os
import zipfile
import secrets


def extractXL(zip_path, zip_file, zip_pass, xl_file, tmpdirname):
    procFile = os.path.join(zip_path, zip_file)
    if zipfile.is_zipfile(procFile):
        try:
            with zipfile.ZipFile(procFile, 'r') as myzip:
                myzip.setpassword(bytes(zip_pass, 'utf-8'))
                if xl_file in myzip.namelist():
                    extractFile = myzip.extract(xl_file, tmpdirname)
                    return extractFile
                else:
                    raise Exception('Unable to find zipped xls file')
        except zipfile.BadZipFile: # if the zip file has any errors then it prints the error message which you wrote under the 'except' block
            raise Exception('File has errors')
    else:
        raise Exception('Unable to process file')


# Get the argument of the number to process
parser = argparse.ArgumentParser(description='bank SEPA file processor',
                                 epilog='If upload file doesn\'t have a valid email and password it will be deleted.')
parser.add_argument('--zipname', type=str, help='Enter the ZIP file you wish to process', required=True)
parser.add_argument("--zippath", type=str, help='Directory where the file is located', default=tempfile.gettempdir())
parser.add_argument("--zippass", type=str, help='The password of the zip file', default=secrets.tmp_zippass())
parser.add_argument("--bankSCTE", type=str, help='The password to archive the bank SCTE file', default=secrets.bnk_scte())
parser.add_argument("--xlfile", type=str, help='The name of the xlsm file in the archive', default='BnkSEPA.xlsm')
args = parser.parse_args()

print('Processing :', args.zippath, "\\", args.zipname)

# Extract the Zip
with tempfile.TemporaryDirectory() as tmpdirname:
    xlsx_filepath = extractXL(args.zippath, args.zipname, args.zippass, args.xlfile, tmpdirname)
    procXlsx.procXL(args.zippath, xlsx_filepath, tmpdirname, args.bankSCTE)

    # clean up
    del xlsx_filepath, tmpdirname
