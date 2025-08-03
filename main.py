# Payments Business Electronic Banking Services
# SEPA Credit Transfers file layout
# Pain.001.001.09 (February 2025)
# Alan Bonnici - chribonn@gmail.com
# U: https://www.AlanBonnici.com
# Yt: https://www.youtube.com/@chribonn
# Last update: 202503
# version - 2.00.00
# Project repository: https://www.github.com/chribonn/bnkSEPA

import argparse
import procXlsx
import tempfile
import os
import zipfile
import secrets
import sys


def extract_xl(zip_path, zip_file, zip_pass, xl_file, tmpdirname):
    proc_file = os.path.join(zip_path, zip_file)
    if zipfile.is_zipfile(proc_file):
        try:
            with zipfile.ZipFile(proc_file, 'r') as myzip:
                myzip.setpassword(bytes(zip_pass, 'utf-8'))
                if xl_file in myzip.namelist():
                    extract_file = myzip.extract(xl_file, tmpdirname)
                    return extract_file
                else:
                    critical_err = 'Unable to find zipped xls file'
                    print('\n\n' + critical_err+ '\n\n')
                    raise Exception(critical_err)
        # if the zip file has any errors then it prints the error message which you wrote under the 'except' block
        except zipfile.BadZipFile:
            critical_err = 'File has errors'
            print('\n\n' + critical_err+ '\n\n')
            raise Exception(critical_err)
    else:
        critical_err = 'Unable to process file'
        print('\n\n' + critical_err+ '\n\n')
        raise Exception(critical_err)


if __name__ == '__main__':
    # Get the argument of the number to process
    parser = argparse.ArgumentParser(
        description='bank SEPA file processor',
        epilog='If upload file doesn\'t have a valid email and password it will be deleted.')
    parser.add_argument(
        '--zipname',
        type=str,
        help='Enter the ZIP file you wish to process',
        default=secrets.zip_file()
    )
    parser.add_argument(
        "--zippath",
        type=str,
        help='Directory where the file is located',
        default=secrets.zip_path()
    )
    parser.add_argument(
        "--zippass",
        type=str,
        help='The password of the zip file',
        default=secrets.tmp_zippass()
    )
     
    parser.add_argument(
        "--xlfile",
        type=str,
        help='The name of the xlsm file in the archive',
        default=secrets.xl_file()
    )
    args = parser.parse_args()

    if args.zipname is None or len(args.zipname) < 1:
        critical_err = 'Zip filename is mandatory'
        print('\n\n' + critical_err+ '\n\n')
        raise Exception(critical_err)

    print('Processing : ', args.zippath, "\\", args.zipname, sep='')

    # Capture processing errors to a file
    error_log_file = os.path.join(args.zippath, "error_log.txt")
    # Redirect standard error to a file
    sys.stderr = open(error_log_file, "w")  # Open in append mode
    
    # Extract the Zip
    with tempfile.TemporaryDirectory() as tmpdirname:
        try:
            xlsx_filepath = extract_xl(args.zippath, args.zipname, args.zippass, args.xlfile, tmpdirname)
            procXlsx.procXL(args.zippath, xlsx_filepath, sys.stderr)

            # clean up
            del xlsx_filepath
        except Exception as e:
            print(f"An error occurred: {e}", file=sys.stderr)

    sys.stderr.close()

    # Check if the error log file has been written to
    if os.path.exists(error_log_file) and os.path.getsize(error_log_file) > 0:
        print(f"\n\nErrors were logged to: {error_log_file}\n\n")
    else:
        print('\n\nProcess completed successfully\n\n')
        os.remove(error_log_file) 

    del error_log_file

    input('Press Enter to terminate.')
