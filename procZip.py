import os, zipfile

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
                    raise Exception('Unable to find zipped xlsl file')
        except zipfile.BadZipFile: # if the zip file has any errors then it prints the error message which you wrote under the 'except' block
            raise Exception('File has errors')
    else:
        raise Exception('Unable to process file')
