import pandas as pd
import numpy as np

import zipfile
from xml.etree import ElementTree as ET
import os
from oletools.olevba import VBA_Parser
import re



ROOT            = r"Z:/38 - Private Debt/OPERATIONS/50_Projects/99_ESG/_EET/_EET_InternalProduction/2023 EET/_Production/ESG EET  I  3SC Pride Fund  I  2023 12.xlsm"
XML_NAME        = 'xl/workbook.xml'
VBA_NAME        = 'xl/vbaProject.bin'
VBA_TEMP        = 'vbaProject.bin'


EXTRACT         = []
VBA_STARTERS    = ["Sub", "Private Sub", "Function", "Private Function"]
VBA_ENDERS      = ["End Sub", "End Function"]




def CHECK_FOR_STARTERS(VBA_CODE, STARTER_LIST):
    for STARTER in STARTER_LIST:
        if STARTER in VBA_CODE:     return True
    return False


def SPLIT_VBA_CODE(VBA_CODE, VBA_ENDERS, VBA_FILENAME):
    PATTERN     = r'(?P<block>(' + '|'.join(re.escape(starter) for starter in VBA_STARTERS) + r').*?(' + '|'.join(re.escape(ender) for ender in VBA_ENDERS) + r'))'
    MATCHES     = re.finditer(PATTERN, VBA_CODE, re.DOTALL | re.MULTILINE)
    return [MATCH.group('block').strip() for MATCH in MATCHES]



def VISIBLE_SHEET_NAME(ZIP_FILE):
    with ZIP_FILE.open('xl/workbook.xml') as WORKBOOK_XML:
        TREE            = ET.parse(WORKBOOK_XML)
        TREE_ROOT       = TREE.getroot()
        NS              = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        SHEET_NAMES     = []

        for SHEET in TREE_ROOT.findall('.//n:sheets/n:sheet', NS):
            SHEET_NAMES.append(SHEET.attrib['name'])
        return SHEET_NAMES


def CODE_NAMES(EXCEL_ZIP):
    CODE_NAMES, CLEAN_NAME, MOD_NAMES, MOD_REF = [], [], [], []

    if VBA_NAME in EXCEL_ZIP.namelist():
        with open(VBA_TEMP, 'wb') as VBA_FILE:
            VBA_FILE.write(EXCEL_ZIP.read(VBA_NAME))
        
        VBA_PARSER = VBA_Parser(VBA_TEMP)
        if VBA_PARSER.detect_vba_macros():

            for (FILENAME, STREAM_PATH, VBA_FILENAME, VBA_CODE) in VBA_PARSER.extract_macros():
                if "ThisWorkbook" not in VBA_FILENAME and VBA_FILENAME.endswith('.cls'): 
                    CODE_NAMES.append(VBA_FILENAME)
                    CLEAN_CODE_NAME = VBA_FILENAME.split('.cls')[0]
                    CLEAN_NAME.append(CLEAN_CODE_NAME)
                
                if "ThisWorkbook" not in VBA_FILENAME and VBA_FILENAME.endswith('.bas'): 
                    MOD_NAMES.append(VBA_FILENAME)
                    MOD_REF_1 = VBA_FILENAME.split('.bas')[0]
                    MOD_REF.append(MOD_REF_1)


    return CODE_NAMES, CLEAN_NAME, MOD_NAMES, MOD_REF








def VBA_SOURCING(ROOT, VBA_NAME, VBA_TEMP, EXTRACT, VBA_STARTERS, VBA_ENDERS):

    with zipfile.ZipFile(ROOT, 'r') as EXCEL_ZIP:
        if VBA_NAME in EXCEL_ZIP.namelist():
            with open(VBA_TEMP, 'wb') as VBA_FILE:
                VBA_FILE.write(EXCEL_ZIP.read(VBA_NAME))
            

            VBA_PARSER = VBA_Parser(VBA_TEMP)
            

            if VBA_PARSER.detect_vba_macros():

                for (FILENAME, STREAM_PATH, VBA_FILENAME, VBA_CODE) in VBA_PARSER.extract_macros():
                    if VBA_CODE.strip():
                    
                        if any(STARTER in VBA_CODE for STARTER in VBA_STARTERS):
                            print(f"Processing VBA Code in Module: {VBA_FILENAME}")
                            SPLIT_LIST = SPLIT_VBA_CODE(VBA_CODE, VBA_ENDERS, VBA_FILENAME)

                            for ITEM in SPLIT_LIST:
                                EXTRACT.append({    'VBA_FILENAME'  : VBA_FILENAME,
                                                    'VBA_CODE'      : ITEM})

                    else: print(f"No VBA code found in Module: {VBA_FILENAME}")
            else: print("No VBA macros found.")
        else: print(f"No VBA project found in {ROOT}.")

    VBA_DF = pd.DataFrame(EXTRACT)

    return VBA_DF



def SHEET_NAME_MAP(ROOT):


    with zipfile.ZipFile(ROOT, 'r') as ZIP_FILE:    
        VISIBLE_SHEET_NAMES                                 = VISIBLE_SHEET_NAME(ZIP_FILE)
        CODE_SHEET_NAMES, CLEAN_NAMES, MOD_NAMES, MOD_REF   = CODE_NAMES(ZIP_FILE)

        SHEETS_DF           = pd.DataFrame({'MODULE_TYPE'           : 'Class Module',
                                            'SHEET_NAME'            : VISIBLE_SHEET_NAMES,
                                            'VBA_CLEAN_SHEET_NAME'  : CLEAN_NAMES,
                                            'VBA_SHEET_NAME'        : CODE_SHEET_NAMES})

        MOD_SHEETS_DF       = pd.DataFrame({'MODULE_TYPE'           : 'Standard Module',
                                            'SHEET_NAME'            : 'NA',
                                            'VBA_CLEAN_SHEET_NAME'  : MOD_REF,
                                            'VBA_SHEET_NAME'        : MOD_NAMES})


        SHEETS_DF = (pd.concat([SHEETS_DF, MOD_SHEETS_DF], ignore_index=True)).reset_index().iloc[:,1:]

    return SHEETS_DF






SHEETS_DF   = SHEET_NAME_MAP(ROOT)
VBA_DF      = VBA_SOURCING(ROOT, VBA_NAME, VBA_TEMP, EXTRACT, VBA_STARTERS, VBA_ENDERS)



print('')
print(SHEETS_DF) 
print('')
print(VBA_DF)
print('')


