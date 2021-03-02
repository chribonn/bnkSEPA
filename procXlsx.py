import openpyxl
import datetime
import os
import pyminizip
import lxml.etree as etree


def procXL(zip_path, xlsx_file, tempdir):
    workbook = openpyxl.load_workbook(filename=xlsx_file, data_only=True)
    # check that the required sheets are in the Excel File
    setXLFiles = set(workbook.sheetnames)
    if not {'Header Record', 'Payment Information Record', 'Credit Instruction Record', 'Control',
            'Control Data (Hidden)'}.issubset(setXLFiles):
        raise Exception('The XL File is not structured properly')

    # Build the XML document
    nsmap = {
        'xsi': "http://www.w3.org/2001/XMLSchema-instance",
        None: "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
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
    bovPass = sh['A2'].value
    if bovPass is None or bovPass.strip() == '':
        raise Exception('Invalid SCTE archive password')
    xmlFile = sh['B2'].value
    if xmlFile is None or xmlFile.strip() == '':
        raise Exception('Invalid SCT file name')

    srcFile = xmlFile.strip() + ".SCT"
    fileSCT = os.path.join(tempdir, srcFile)
    try:
        with open(fileSCT, 'wb') as file:
            file.write(datastr)
    except:
        raise Exception('Unable to create SCT file')

    # package everything in the zip file
    # in web interface replace C:\Temp with tempdir as the file will be emailed
    zipSCTE = os.path.join(zip_path, xmlFile.strip() + ".SCTE")
    pyminizip.compress(fileSCT, None, zipSCTE, bovPass, 0)


def bldCIR(PmtInf, workbook):
    sh = workbook['Credit Instruction Record']

    row = 5
    while row < 100:
        sInstrId = sh['A' + str(row)].value
        if sInstrId is None:
            break

        PmtInf = bldCIRrow(sh, PmtInf, workbook, row)
        row += 2

    return PmtInf


def bldCIRrow(sh, PmtInf, workbook, row):
    # Process the particular row

    # Read the Fields from this worksheet
    sInstrId = sh['A' + str(row)].value.strip()
    sEndToEndId = sh['B' + str(row)].value.strip()
    sCcy = sh['C' + str(row)].value.strip()
    sInstdAmt = '{0:.2f}'.format(sh['D' + str(row)].value)
    sBIC = sh['E' + str(row)].value.strip()
    sNm = sh['F' + str(row)].value.strip()
    sAdrLine1 = sh['G5'].value
    if sAdrLine1 is None:
        sAdrLine1 = ''
    else:
        sAdrLine1 = sAdrLine1.strip()
    sAdrLine2 = sh['H5'].value
    if sAdrLine2 is None:
        sAdrLine2 = ''
    else:
        sAdrLine2 = sAdrLine2.strip()
    if sAdrLine1 == '':
        sAdrLine1 = sAdrLine2
    sIBAN = sh['I' + str(row)].value.strip()
    sCd = sh['J' + str(row)].value.strip()
    sUstrd = sh['K' + str(row)].value.strip()

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
    BIC = etree.SubElement(FinInstnId, "BIC")
    BIC.text = sBIC
    Cdtr = etree.SubElement(CdtTrfTxInf, "Cdtr")
    Nm = etree.SubElement(Cdtr, "Nm")
    Nm.text = sNm
    # Only fill in the subnodes if the address lines are not blank
    if sAdrLine1 != "":
        PstlAdr = etree.SubElement(Cdtr, "PstlAdr")
        AdrLine1 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine1.text = sAdrLine1
        AdrLine2 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine2.text = sAdrLine2
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

    return PmtInf


def bldPIR(PmtInf, computedPmtInfld, workbook):
    sh = workbook['Payment Information Record']

    # Read the Fields from this worksheet
    sPmtInfId = sh['A5'].value
    if sPmtInfId is None:
        sPmtInfId = computedPmtInfld
    sPmtInfId = sPmtInfId.strip()
    # Cechk for a space condition
    if sPmtInfId == '':
        sPmtInfId = computedPmtInfld.strip()

    sPmtMtd = sh['B5'].value.strip()
    sBtchBookg = sh['C5'].value.strip()
    sNbOfTxs = str(int(sh['D5'].value))
    sCtrlSum = '{0:.2f}'.format(sh['E5'].value)
    sCd = sh['F5'].value.strip()
    sReqdExctnDt = sh['G5'].value
    sReqdExctnDt = datetime.datetime.strftime(sReqdExctnDt, '%Y-%m-%d')
    sNm = sh['H5'].value.strip()
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
    sIBAN = sh['K5'].value.strip()
    sCcy = sh['L5'].value.strip()
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
    ReqdExctnDt.text = sReqdExctnDt
    Dbtr = etree.SubElement(PmtInf, "Dbtr")
    Nm = etree.SubElement(Dbtr, "Nm")
    Nm.text = sNm
    # Only fill in the subnodes if the address lines are not blank
    if sAdrLine1 != "":
        PstlAdr = etree.SubElement(Dbtr, "PstlAdr")
        AdrLine1 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine1.text = sAdrLine1
        AdrLine2 = etree.SubElement(PstlAdr, "AdrLine")
        AdrLine2.text = sAdrLine2
    DbtrAcct = etree.SubElement(PmtInf, "DbtrAcct")
    Id = etree.SubElement(DbtrAcct, "Id")
    IBAN = etree.SubElement(Id, "IBAN")
    IBAN.text = sIBAN
    Ccy = etree.SubElement(DbtrAcct, "Ccy")
    Ccy.text = sCcy
    DbtrAgt = etree.SubElement(PmtInf, "DbtrAgt")
    FinInstnId = etree.SubElement(DbtrAgt, "FinInstnId")
    BIC = etree.SubElement(FinInstnId, "BIC")
    BIC.text = sBIC

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
        sCreDtTm = datetime.datetime.strptime(sCreDtTm, "%Y-%m-%d %H:%M:%S").isoformat()

    sNbOfTxs = str(int(sh['C5'].value))
    sCtrlSum = '{0:.2f}'.format(sh['D5'].value)
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
