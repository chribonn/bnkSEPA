import os, zipfile


def extract_xl(zip_path, zip_file, zip_pass, xl_file, tmp_dir_name):
    procFile = os.path.join(zip_path, zip_file)
    if zipfile.is_zipfile(procFile):
        try:
            with zipfile.ZipFile(procFile, 'r') as myzip:
                myzip.setpassword(bytes(zip_pass, 'utf-8'))
                if xl_file in myzip.namelist():
                    extractFile = myzip.extract(xl_file, tmp_dir_name)
                    return extractFile
                else:
                    critical_err = 'Unable to find zipped xlsx file'
                    print('\n\n' + critical_err + '\n\n')
                    input('Press Enter to terminate.')
                    raise Exception(critical_err)
        except zipfile.BadZipFile: # if the zip file has any errors then it prints the error message which you wrote under the 'except' block
            critical_err = 'File has errors'
            print('\n\n' + critical_err + '\n\n')
            input('Press Enter to terminate.')
            raise Exception(critical_err)
    else:
        critical_err = 'Unable to process file: Check whether the zip file is correct.'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)
 