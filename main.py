# Payments Business Electronic Banking Services
# SEPA Credit Transfers file layout
# Pain.001.001.09 (February 2025 (email))
# Alan Bonnici - chribonn@gmail.com
# Last update: 202503
# version - 2.10.00
# Project repository: https://www.github.com/chribonn/bnkSEPA

# From: Joseph Dxxxxx <xxxxxxxx@bov.com>
# Date: Tue, Feb 25, 2025 at 12:01 PM
# Subject: Action Needed for SEPA Payment File Formats
# Email instruction: 
#        Updates you need to make to your ‘SEPA Payment File Formats’
#
#        The main change from the previous version is that this version supports more structured fields in the postal address. 
#        The structured address details of payee must meet the following requirements:
#
#        a. Data element “Address Line” < Adrline> cannot be used.
#        b. The Data elements “Town Name” < TwnNm> and “Country” <Ctry> must be used.
#        c. Other Data elements like “Street Name” <StrtNm>,“Building Number” <BldgNb>,“Building Name” < BldgNm , “Post Code” <PstCd> etc may be used 
#           depending on the components of the Address.
#
#        Also, kindly note change in header: Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.001.001.09">
#

import argparse
import procXlsx
import tempfile
import os
import zipfile
import secrets


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
                    input('Press Enter to terminate.')
                    raise Exception(critical_err)
        # if the zip file has any errors then it prints the error message which you wrote under the 'except' block
        except zipfile.BadZipFile:
            critical_err = 'File has errors'
            print('\n\n' + critical_err+ '\n\n')
            input('Press Enter to terminate.')
            raise Exception(critical_err)
    else:
        critical_err = 'Unable to process file'
        print('\n\n' + critical_err+ '\n\n')
        input('Press Enter to terminate.')
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
        default=tempfile.gettempdir()
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
        input('Press Enter to terminate.')
        raise Exception(critical_err)

    print('Processing : ', args.zippath, "\\", args.zipname, sep='')

    # Extract the Zip
    with tempfile.TemporaryDirectory() as tmpdirname:
        xlsx_filepath = extract_xl(args.zippath, args.zipname, args.zippass, args.xlfile, tmpdirname)
        procXlsx.procXL(args.zippath, xlsx_filepath)

        # clean up
        del xlsx_filepath, tmpdirname

        print('\n\nProcess completed successfully\n\n')
        input('Press Enter to terminate.')
