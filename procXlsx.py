import openpyxl
import datetime
import os
import lxml.etree as etree
import warnings


def procXL(zip_path, xlsx_file, err_dev):
    warnings.simplefilter(action='ignore', category=UserWarning)
    workbook = openpyxl.load_workbook(filename=xlsx_file, data_only=True)
    warnings.resetwarnings()
    # check that the required sheets are in the Excel File
    setXLFiles = set(workbook.sheetnames)
    if not {'Header Record', 'Payment Information Record', 'Credit Instruction Record', 'Control',
            'Control Data (Hidden)'}.issubset(setXLFiles):

        critical_err = 'The XL File is not structured properly'
        print(critical_err)
        raise Exception(critical_err)

    # Build the XML document
    nsmap = {
        'xsi': "http://www.w3.org/2001/XMLSchema-instance",
        None: "urn:iso:std:iso:20022:tech:xsd:pain.001.001.09"
    }
    root = etree.Element('Document', nsmap=nsmap)

    # Fill in the Computed MsgId and PmtInfld to (necessary if the user leaves these blank)
    sh = workbook['Control Data (Hidden)']
    computedMsgId = sh['B18'].value
    computedPmtInfld = sh['B20'].value

    CstmrCdtTrfInitn = etree.SubElement(root, 'CstmrCdtTrfInitn')

    # Header Record
    CstmrCdtTrfInitn = bldHeader(CstmrCdtTrfInitn, computedMsgId, workbook)

    # This tag covers both PIR and CIR sections
    PmtInf = etree.SubElement(CstmrCdtTrfInitn, "PmtInf")

    # Payment Information Record
    PmtInf = bldPIR(PmtInf, computedPmtInfld, workbook)

    # Credit Instruction Record
    PmtInf = bldCIR(PmtInf, workbook)

    datastr = etree.tostring(root, xml_declaration=True, encoding='utf-8', pretty_print=True)

    # Get the name of the XML file that will store the transactions
    sh = workbook['Control']
    xmlFile = sh['B2'].value
    if xmlFile is None or xmlFile.strip() == '':
        critical_err = 'Invalid SCT file name'
        print(critical_err)
        input('Press Enter to terminate.')
        raise Exception(critical_err)

    srcFile = xmlFile.strip() + ".SCT"
    # fileSCT = os.path.join(tempdir, srcFile) - write file to final directory rather than tempdir
    fileSCT = os.path.join(zip_path, srcFile)
    try:
        with open(fileSCT, 'wb') as file:
            file.write(datastr)
    except:
        critical_err = 'Unable to create SCT file'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)


def bldCIR(PmtInf, workbook):
    sh = workbook['Credit Instruction Record']

    row = 5
    lstEndToEndId = []
    # list stores the sEndToEndId values. If there are duplicate entries raises an exception  - ACB 202309
    while row < 200:
        sInstrId = sh['A' + str(row)].value
        if sInstrId is None:
            break

        result = bldCIRrow(sh, PmtInf, workbook, row)
        PmtInf = result[0]
        sEndToEndId = result[1]

        if sEndToEndId in lstEndToEndId:
            critical_err = 'EndToEndId {0} has been already used in this batch'.format(sEndToEndId)
            print('\n\n' + critical_err+ '\n\n')
            input('Press Enter to terminate.')
            raise Exception(critical_err)
        else:
            lstEndToEndId.append(sEndToEndId)
            
        row += 2

    return PmtInf


def bldCIRrow(sh, PmtInf, workbook, row):
    # Process the particular row

    # Read the Fields from this worksheet
    sInstrId = sh['A' + str(row)].value.strip()
    sEndToEndId = sh['B' + str(row)].value.strip()
    sCcy = sh['C' + str(row)].value.strip()
    sInstdAmt = '{0:.2f}'.format(sh['D' + str(row)].value)
    sBICFI = sh['E' + str(row)].value.strip()
    sNm = sh['F' + str(row)].value.strip()
    
    #################################################################
    # Changes as per version v 8.7 of the SCT document - ACB 202503
    sStrtNm = sh['G5'].value
    if sStrtNm is None:
        sStrtNm = ''
    else:
        sStrtNm = sStrtNm.strip()
    sBldgNb = sh['H5'].value
    if sBldgNb is None:
        sBldgNb = ''
    else:
        sBldgNb = sBldgNb.strip()
    sBldgNm = sh['I5'].value
    if sBldgNm is None:
        sBldgNm = ''
    else:
        sBldgNm = sBldgNm.strip()
    sPstCd = sh['J5'].value
    if sPstCd is None:
        sPstCd = ''
    else:
        sPstCd = sPstCd.strip()
    sTwnNm = sh['K5'].value
    if sTwnNm is None:
        sTwnNm = ''
    else:
        sTwnNm = sTwnNm.strip()
    sCtry = sh['L5'].value
    if sCtry is None:
        sCtry = ''
    else:
        sCtry = sCtry.strip()

    """
    # Removed to reflect changes in address logic as per version v 8.7 of the SCT document - ACB 202503
    if sStrtNm == '':
        sStrtNm = sBldgNb
    """
    #################################################################

    sIBAN = sh['M' + str(row)].value.strip()
    sCd = sh['N' + str(row)].value.strip()
    sUstrd = sh['O' + str(row)].value.strip()

    CdtTrfTxInf = etree.SubElement(PmtInf, "CdtTrfTxInf")
    PmtId = etree.SubElement(CdtTrfTxInf, "PmtId")
    InstrId = etree.SubElement(PmtId, "InstrId")
    InstrId.text = sInstrId
    EndToEndId = etree.SubElement(PmtId, "EndToEndId")
    EndToEndId.text = sEndToEndId
    Amt = etree.SubElement(CdtTrfTxInf, "Amt")
    InstdAmt = etree.SubElement(Amt, "InstdAmt")
    InstdAmt.set('Ccy', sCcy)
    InstdAmt.text = sInstdAmt
    CdtrAgt = etree.SubElement(CdtTrfTxInf, "CdtrAgt")
    FinInstnId = etree.SubElement(CdtrAgt, "FinInstnId")
    BICFI = etree.SubElement(FinInstnId, "BICFI")
    BICFI.text = sBICFI
    Cdtr = etree.SubElement(CdtTrfTxInf, "Cdtr")
    Nm = etree.SubElement(Cdtr, "Nm")
    Nm.text = sNm
    
    """
    # Removed to reflect changes in address logic as per version v 8.7 of the SCT document - ACB 202503
    
    # Only fill in the subnodes if the address lines are not blank
    if sStrtNm != "":
        PstlAdr = etree.SubElement(Cdtr, "PstlAdr")
        AdrLine1 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine1.text = sStrtNm
        AdrLine2 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine2.text = sBldgNb
    """
    PstlAdr = etree.SubElement(Cdtr, "PstlAdr")
    if sStrtNm != "":
        StrtNm = etree.SubElement(PstlAdr, "StrtNm")
        StrtNm.text = sStrtNm
    if sBldgNb != "":
        BldgNb = etree.SubElement(PstlAdr, "BldgNb")
        BldgNb.text = sBldgNb
    if sBldgNm != "":
        BldgNm = etree.SubElement(PstlAdr, "BldgNm")
        BldgNm.text = sBldgNm
    if sPstCd != "":
        PstCd = etree.SubElement(PstlAdr, "PstCd")
        PstCd.text = sPstCd
    TwnNm = etree.SubElement(PstlAdr, "TwnNm")
    TwnNm.text = sTwnNm
    Ctry = etree.SubElement(PstlAdr, "Ctry")
    Ctry.text = sCtry
    
    CdtrAcct = etree.SubElement(CdtTrfTxInf, "CdtrAcct")
    Id = etree.SubElement(CdtrAcct, "Id")
    IBAN = etree.SubElement(Id, "IBAN")
    IBAN.text = sIBAN
    Purp = etree.SubElement(CdtTrfTxInf, "Purp")
    Cd = etree.SubElement(Purp, "Cd")
    Cd.text = sCd
    RmtInf = etree.SubElement(CdtTrfTxInf, "RmtInf")
    Ustrd = etree.SubElement(RmtInf, "Ustrd")
    Ustrd.text = sUstrd

    return PmtInf, sEndToEndId


def bldPIR(PmtInf, computedPmtInfld, workbook):
    sh = workbook['Payment Information Record']

    # Read the Fields from this worksheet
    sPmtInfId = sh['A5'].value
    if sPmtInfId is None:
        sPmtInfId = computedPmtInfld
    sPmtInfId = sPmtInfId.strip()
    # Check for a space condition
    if sPmtInfId == '':
        sPmtInfId = computedPmtInfld.strip()

    sPmtMtd = sh['B5'].value.strip()
    sBtchBookg = sh['C5'].value.strip()
    sNbOfTxs = str(int(sh['D5'].value))
    try:
        sNbOfTxs = str(int(sNbOfTxs))
    except:
        critical_err = 'Payment Information Record: Format error : Cell D5'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)
    
    try:
        sCtrlSum = '{0:.2f}'.format(sh['E5'].value)
    except:
        critical_err = 'Payment Information Record: Format error : Cell E5'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)

    sCd = sh['F5'].value.strip()
    
    try:
        sReqdExctnDt = sh['G5'].value
        sReqdExctnDt = datetime.datetime.strftime(sReqdExctnDt, '%Y-%m-%d')
    except:
        critical_err = 'Payment Information Record: Format error : Cell G5'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)
        
    sNm = sh['H5'].value.strip()
    """
    # Removed the address lines from the PIR section - ACB 202503
    # as per version v 8.7 of the SCT document
    sAdrLine1 = sh['I5'].value
    if sAdrLine1 is None:
        sAdrLine1 = ''
    else:
        sAdrLine1 = sAdrLine1.strip()
    sAdrLine2 = sh['J5'].value
    if sAdrLine2 is None:
        sAdrLine2 = ''
    else:
        sAdrLine2 = sAdrLine2.strip()
    if sAdrLine1 == '':
        sAdrLine1 = sAdrLine2
    `"""
    sIBAN = sh['K5'].value.strip()
    # Removed from the PIR section as per version v 8.7 of the SCT document - ACB 202503
    # sCcy = sh['L5'].value.strip()
    sBIC = sh['M5'].value.strip()

    PmtInfId = etree.SubElement(PmtInf, "PmtInfId")
    PmtInfId.text = sPmtInfId
    PmtMtd = etree.SubElement(PmtInf, "PmtMtd")
    PmtMtd.text = sPmtMtd
    BtchBookg = etree.SubElement(PmtInf, "BtchBookg")
    BtchBookg.text = sBtchBookg
    NbOfTxs = etree.SubElement(PmtInf, "NbOfTxs")
    NbOfTxs.text = sNbOfTxs
    CtrlSum = etree.SubElement(PmtInf, "CtrlSum")
    CtrlSum.text = sCtrlSum
    PmtTpInf = etree.SubElement(PmtInf, "PmtTpInf")
    SvcLvl = etree.SubElement(PmtTpInf, "SvcLvl")
    Cd = etree.SubElement(SvcLvl, "Cd")
    Cd.text = sCd
    ReqdExctnDt = etree.SubElement(PmtInf, "ReqdExctnDt")
    # Added <Dt> subtag as per version v 8.7 of the SCT document - ACB 202503
    Dt = etree.SubElement(ReqdExctnDt, "Dt")
    Dt.text = sReqdExctnDt
    Dbtr = etree.SubElement(PmtInf, "Dbtr")
    Nm = etree.SubElement(Dbtr, "Nm")
    Nm.text = sNm
    """
    # Removed from the PIR section as per version v 8.7 of the SCT document - ACB 202503
    
    # Only fill in the subnodes if the address lines are not blank
    if sAdrLine1 != "":
        PstlAdr = etree.SubElement(Dbtr, "PstlAdr")
        AdrLine1 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine1.text = sAdrLine1
        # Only fill if Address line 2 is not null
        if sAdrLine2 != "":
            AdrLine2 = etree.SubElement(PstlAdr, "AdrLine")
            AdrLine2.text = sAdrLine2
    """
    DbtrAcct = etree.SubElement(PmtInf, "DbtrAcct")
    Id = etree.SubElement(DbtrAcct, "Id")
    IBAN = etree.SubElement(Id, "IBAN")
    IBAN.text = sIBAN
    """
    # Removed from the PIR section as per version v 8.7 of the SCT document - ACB 202503
    Ccy = etree.SubElement(DbtrAcct, "Ccy")
    Ccy.text = sCcy
    """
    DbtrAgt = etree.SubElement(PmtInf, "DbtrAgt")
    FinInstnId = etree.SubElement(DbtrAgt, "FinInstnId")
    BICFI = etree.SubElement(FinInstnId, "BICFI")
    BICFI.text = sBIC

    return PmtInf


def bldHeader(CstmrCdtTrfInitn, computedMsgId, workbook):
    sh = workbook['Header Record']

    # Read the Fields from this worksheet
    sMsgId = sh['A5'].value
    if sMsgId is None:
        sMsgId = computedMsgId
    sMsgId = sMsgId.strip()
    # check for a space condition
    if sMsgId == '':
        sMsgId = computedMsgId.strip()

    sCreDtTm = str(sh['B5'].value)
    # cater for different formats with microseconds and without
    try:
        sCreDtTm = datetime.datetime.strptime(sCreDtTm, "%Y-%m-%d %H:%M:%S.%f").replace(microsecond=0).isoformat()
    except:
        try:
            sCreDtTm = datetime.datetime.strptime(sCreDtTm, "%Y-%m-%d %H:%M:%S").isoformat()
        except:
            critical_err = 'Header Record: Format error : Cell B5'
            print('\n\n' + critical_err + '\n\n')
            input('Press Enter to terminate.')
            raise Exception(critical_err)

    sNbOfTxs = str(int(sh['C5'].value))
    try:
        sNbOfTxs = str(int(sNbOfTxs))
    except:
        critical_err = 'Header Record: Format error : Cell C5'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)
        
    try:
        sCtrlSum = '{0:.2f}'.format(sh['D5'].value)
    except:
        critical_err = 'Header Record: Format error : Cell D5'
        print('\n\n' + critical_err + '\n\n')
        input('Press Enter to terminate.')
        raise Exception(critical_err)
    
    sNm = sh['E5'].value.strip()
    sId = sh['F5'].value.strip()

    GrpHdr = etree.SubElement(CstmrCdtTrfInitn, "GrpHdr")
    MsgId = etree.SubElement(GrpHdr, "MsgId")
    MsgId.text = sMsgId
    CreDtTm = etree.SubElement(GrpHdr, "CreDtTm")
    CreDtTm.text = sCreDtTm
    NbOfTxs = etree.SubElement(GrpHdr, "NbOfTxs")
    NbOfTxs.text = sNbOfTxs
    CtrlSum = etree.SubElement(GrpHdr, "CtrlSum")
    CtrlSum.text = sCtrlSum
    InitgPty = etree.SubElement(GrpHdr, "InitgPty")
    Nm = etree.SubElement(InitgPty, "Nm")
    Nm.text = sNm
    Id1 = etree.SubElement(InitgPty, "Id")
    OrgId = etree.SubElement(Id1, "OrgId")
    Othr = etree.SubElement(OrgId, "Othr")
    Id2 = etree.SubElement(Othr, "Id")
    Id2.text = sId

    return CstmrCdtTrfInitn
