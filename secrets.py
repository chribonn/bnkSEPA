# This file should not be loaded to GitHub WITH LIVE DATA
# =======================================================

# This contains the name of the ZIP File that stores within it the Excel XLSM file
def zip_file():
    return 'BnkSEPA.zip'

# This contains the password that has been used to encrypt the ZIP file that contains the Excel XLSM file
def tmp_zippass():
    return '87654321'

# This contains the name of the Excel XLSM file that contain the transactions that need to be processed
def xl_file():
    return 'BnkSEPA.xlsm'

# This contains the password that will be used to password protect the file for transmission to the bank. This is shared with the bank.
def bnk_scte():
    return '12345678'
