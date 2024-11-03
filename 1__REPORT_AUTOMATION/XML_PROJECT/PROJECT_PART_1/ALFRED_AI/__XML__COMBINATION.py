import re
import json
import ast
import zipfile
import string
import openpyxl
import pandas as pd
import networkx as nx
import xml.etree.ElementTree as ET
from collections import defaultdict
from openpyxl import load_workbook
from oletools.olevba import VBA_Parser













def PROCESS_TABLES(PARSED_DATA, LABELS, DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS, P3_RESULTS):

    def CELL_DEPENDENCIES(GRAPH, CELL_REFERENCE):
        if CELL_REFERENCE in GRAPH:
            DESCENDENTS                 = list(nx.descendants(GRAPH, CELL_REFERENCE))
            ANCESTORS                   = list(nx.ancestors(GRAPH, CELL_REFERENCE))
        else:DESCENDENTS, ANCESTORS     = [], []
        return DESCENDENTS, ANCESTORS
    

    def CONVERT_HEADERS_TO_DICT(DETECTED_TABLES):
        for SHEET_NAME, TABLES in DETECTED_TABLES.items():
            for TABLE in TABLES:
                HEADERS = TABLE.get('HEADER', [])
                START_CELL = TABLE.get('START_CELL', '')
                if HEADERS and START_CELL:
                    START_COLUMN_LETTER = START_CELL[0]
                    TABLE['HEADER'] = {chr(ord(START_COLUMN_LETTER) + I): HEADER for I, HEADER in enumerate(HEADERS)}
        return DETECTED_TABLES

    DETECTED_TABLES = CONVERT_HEADERS_TO_DICT(DETECTED_TABLES)

    def GET_RANGE_COORDINATES(RANGE_STR):
        MATCH = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', RANGE_STR)
        if MATCH:
            COL_START, ROW_START, COL_END, ROW_END = MATCH.groups()
            return (COL_START, int(ROW_START)), (COL_END, int(ROW_END))
        return None

    def COL_TO_NUM(COL):
        NUM = 0
        for C in COL:
            NUM = NUM * 26 + (ord(C) - ord('A') + 1)
        return NUM

    def CHECK_CLUSTER_RANGE(OUTPUT_DICT):
        CLUSTERS = OUTPUT_DICT.get('CLUSTERS', [])
        PIVOT_LOCATION = OUTPUT_DICT.get('PIVOT_PARAMS', {}).get('LOCATION')
        PIVOT_RANGE = GET_RANGE_COORDINATES(PIVOT_LOCATION) if PIVOT_LOCATION else None
        TABLE_LOCATION = OUTPUT_DICT.get('TABLE_PARAMS', {}).get('LOCATION')
        TABLE_RANGE = GET_RANGE_COORDINATES(TABLE_LOCATION) if TABLE_LOCATION else None

        for CLUSTER in CLUSTERS:
            START_CELL = CLUSTER.get('START_CELL')
            ROWS = CLUSTER.get('ROWS', 0)
            COLUMNS = CLUSTER.get('COLUMNS', 0)
            CLUSTER_START_COL = START_CELL[0]
            CLUSTER_START_ROW = int(START_CELL[1:])
            CLUSTER_END_COL = chr(ord(CLUSTER_START_COL) + COLUMNS - 1)
            CLUSTER_END_ROW = CLUSTER_START_ROW + ROWS - 1
            CLUSTER_RANGE = ((CLUSTER_START_COL, CLUSTER_START_ROW), (CLUSTER_END_COL, CLUSTER_END_ROW))

            if PIVOT_RANGE:
                CLUSTER['PIVOT_RANGE'] = CLUSTER_RANGE == PIVOT_RANGE
            if TABLE_RANGE:
                CLUSTER['TABLE_RANGE'] = CLUSTER_RANGE == TABLE_RANGE

        return OUTPUT_DICT

    def GET_COLUMN_LETTER(START_CELL, COLUMN_INDEX):
        START_COL_LETTER = ''.join([char for char in START_CELL if char.isalpha()])  # Extract letters from START_CELL
        START_COL_NUM = COL_TO_NUM(START_COL_LETTER)
        RETURN_COL_NUM = START_COL_NUM + COLUMN_INDEX
        
        # Convert number back to letter(s) for columns greater than "Z"
        RETURN_COL_LETTER = ''
        while RETURN_COL_NUM > 0:
            RETURN_COL_NUM, remainder = divmod(RETURN_COL_NUM - 1, 26)
            RETURN_COL_LETTER = chr(65 + remainder) + RETURN_COL_LETTER  # 65 is ASCII for 'A'
        
        return RETURN_COL_LETTER

    def CHECK_AND_POPULATE_HEADERS(APPEND_DICT):
        if isinstance(APPEND_DICT['HEADER'], list):
            APPEND_DICT['HEADER'] = {}
        HEADER = APPEND_DICT.get('HEADER', {})
        COLUMNS = APPEND_DICT.get('COLUMNS', 0)
        if not HEADER or len(HEADER) != COLUMNS:
            START_CELL = APPEND_DICT.get('START_CELL')
            NEW_HEADER = HEADER.copy()
            for COL_IDX in range(COLUMNS):
                COL_LETTER = GET_COLUMN_LETTER(START_CELL, COL_IDX)
                if COL_LETTER not in NEW_HEADER:
                    NEW_HEADER[COL_LETTER] = None
            APPEND_DICT['HEADER'] = NEW_HEADER
        return APPEND_DICT

    def UPDATE_COLUMN_DICT(COLUMN_DICT, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, SHEET_NAME):
        COLUMN_LETTER = COLUMN_DICT["COLUMN_LETTER"]
        if SHEET_NAME in CLUSTERED_NESTED_OUTPUT:
            for NESTED_KEY, CELL_LIST in CLUSTERED_NESTED_OUTPUT[SHEET_NAME].items():
                for CELL_REF in CELL_LIST:
                    if CELL_REF[0] == COLUMN_LETTER:
                        COLUMN_DICT["NESTED_STRING"] = NESTED_KEY
                        COLUMN_DICT["NESTED_CLUSTER"] = CELL_LIST
                        break
        if SHEET_NAME in NESTED_OUTPUT:
            for CELL_REF, NESTED_INFO in NESTED_OUTPUT[SHEET_NAME].items():
                if CELL_REF[0] == COLUMN_LETTER:
                    COLUMN_DICT["NESTED_FUNCTION"] = NESTED_INFO["NESTED_DICT"]
                    COLUMN_DICT["NESTED_STRING"] = NESTED_INFO["NESTED_STRING"]
                    break
        return COLUMN_DICT

    def POPULATE_CELLS(COLUMN_DICT, APPEND_DICT, UPDATED_WORKSHEETS, SHEET_NAME, GRAPH):
        COLUMN_LETTER = COLUMN_DICT['COLUMN_LETTER']
        
        for CELL_REF in APPEND_DICT['REF_CELLS']:
            CELL_COLUMN_LETTER = CELL_REF[0]
            if CELL_COLUMN_LETTER == COLUMN_LETTER:
                if CELL_REF in UPDATED_WORKSHEETS[SHEET_NAME]:
                    CELL_DATA = UPDATED_WORKSHEETS[SHEET_NAME][CELL_REF]
                    FULL_CELL_REF = f"{SHEET_NAME}!{CELL_REF}"
                    DESCENDENTS, ANCESTORS = CELL_DEPENDENCIES(GRAPH, FULL_CELL_REF)
                    
                    # Populate dependency information
                    COLUMN_DICT['CELLS'][CELL_REF] = {
                        'TYPE': CELL_DATA.get('TYPE', None),
                        'STYLE_INDEX': CELL_DATA.get('STYLE_INDEX', None),
                        'FORMULA': CELL_DATA.get('FORMULA', None),
                        'VALUE': CELL_DATA.get('VALUE', None),
                        'DEPENDENCY': {
                            'ANCESTORS_LENGTH': len(ANCESTORS),
                            'DESCENDENTS_LENGTH': len(DESCENDENTS),
                            'ANCESTORS': ANCESTORS,
                            'DESCENDENTS': DESCENDENTS
                        }
                    }


    WORKSHEETS = {}
    for SHEET_NAME, SHEET_DATA in PARSED_DATA['WORKSHEETS'].items():
        SHEET_INFO = {
            "SHEET_TYPE": LABELS.get(SHEET_NAME, "UNKNOWN"),
            "TABLE_PRESENT": False,
            "TABLE_PARAMS": {},
            "PIVOT_PRESENT": False,
            "PIVOT_PARAMS": {},
            "NAMED_RANGE_PRESENT":False,
            "NAMED_RANGE_PARAMS":[],
            "CLUSTERS": []
        }

        if PARSED_DATA['TABLES']:
            for TABLE_KEY, TABLE_DATA in PARSED_DATA['TABLES'].items():
                if SHEET_NAME == TABLE_DATA['LOCATION']:
                    SHEET_INFO["TABLE_PRESENT"] = True
                    SHEET_INFO["TABLE_PARAMS"] = {
                        "NAME": TABLE_DATA.get("NAME", None),
                        "START_CELL": TABLE_DATA.get("SHEET_LOCATION", None),
                        "LOCATION": TABLE_DATA.get("REF", None),
                        "DIMENSIONS": {
                            "ROWS": TABLE_DATA['DIMENSIONS'].get("ROWS", None),
                            "COLUMNS": TABLE_DATA['DIMENSIONS'].get("ROCOLSWS", None)
                        },
                        "COLUMNS": TABLE_DATA.get("COLUMNS", [])
                    }

        if PARSED_DATA['PIVOTS']:
            for PIVOT_KEY, PIVOT_DATA in PARSED_DATA['PIVOTS'].items():
                if SHEET_NAME == PIVOT_DATA['SHEET']:
                    SHEET_INFO["PIVOT_PRESENT"] = True
                    SHEET_INFO["PIVOT_PARAMS"] = {
                        "NAME": PIVOT_DATA.get("NAME", None),
                        "LOCATION": PIVOT_DATA.get("LOCATION", None),
                        "CONTENTS": PIVOT_DATA.get("TABLE", None),
                        "SOURCE_TABLE": PIVOT_DATA.get("SOURCE_TABLE", None),
                        "SOURCE_INFO": {
                            "NAME": PIVOT_DATA['SOURCE_INFO'].get("NAME", None),
                            "SHEET": PIVOT_DATA['SOURCE_INFO'].get("SHEET", None),
                            "LOCATION": PIVOT_DATA['SOURCE_INFO'].get("REF", None)
                        }
                    }


        if PARSED_DATA['NAMED_RANGE']:
            for N_R_DATA in PARSED_DATA['NAMED_RANGE']:
                if SHEET_NAME == N_R_DATA['SHEET_NAME']:
                    SHEET_INFO["NAMED_RANGE_PRESENT"] = True
                    SHEET_INFO["NAMED_RANGE_PARAMS"].append(N_R_DATA)


        if SHEET_NAME in DETECTED_TABLES:
            for A in range(len(DETECTED_TABLES[SHEET_NAME])):
                T1 = DETECTED_TABLES[SHEET_NAME][A]
                APPEND_DICT = {
                    "TABLE_TYPE": "CLUSTER",
                    "TABLE_RANGE": False,
                    "PIVOT_RANGE": False,
                    "START_CELL": T1.get("START_CELL", None),
                    "HEADER": T1.get("HEADER", None),
                    "ROWS": T1.get("ROWS", None),
                    "COLUMNS": T1.get("COLUMNS", None),
                    "REF_CELLS": T1.get("CELLS", None),
                    "COLUMN": [],
                }

                APPEND_DICT = CHECK_AND_POPULATE_HEADERS(APPEND_DICT)
                for LETTER, HEADER in APPEND_DICT['HEADER'].items():
                    COLUMN_DICT = {
                        "COLUMN_LETTER": LETTER,
                        "COLUMN_HEADER": HEADER,
                        "FORMULA": False,
                        "NESTED_STRING": 'NA',
                        "NESTED_FUNCTION": {},
                        "NESTED_CLUSTER": [],
                        "CELLS": {}
                    }

                    COLUMN_DICT = UPDATE_COLUMN_DICT(COLUMN_DICT, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, SHEET_NAME)
                    POPULATE_CELLS(COLUMN_DICT, APPEND_DICT, UPDATED_WORKSHEETS, SHEET_NAME, P3_RESULTS['DEPENDENCY_GRAPH'])
                    APPEND_DICT['COLUMN'].append(COLUMN_DICT)

                SHEET_INFO["CLUSTERS"].append(APPEND_DICT)
                SHEET_INFO = CHECK_CLUSTER_RANGE(SHEET_INFO)

                for J in range(len(SHEET_INFO["CLUSTERS"])):
                    if SHEET_INFO["CLUSTERS"][J]['TABLE_RANGE']:
                        SHEET_INFO["CLUSTERS"][J]['TABLE_TYPE'] = "TABLE"
                    elif SHEET_INFO["CLUSTERS"][J]['PIVOT_RANGE'] == True: SHEET_INFO["CLUSTERS"][J]['TABLE_TYPE'] = "PIVOT"



        WORKSHEETS[SHEET_NAME] = SHEET_INFO

    return {"WORKSHEETS": WORKSHEETS}







def OTHER_INFO(PARSED_DATA):


    DICT_EX = {'SHAPES'     :'',
               'VBA'        :'',
               'POWER_QUERY':'',
               'DATA_MODEL' :'',
               'STYLE_USED' :'',
               'THEME_USED' :''}



    def EXTRACT_DRAWING(DRAWDING):

        DRAWING_PRESENT = False
        for DRAWING_FILE, SHAPES in DRAWDING.items():
            if len(SHAPES) > 0 and DRAWING_PRESENT == False: DRAWING_PRESENT = True
                
        DRAWING_DATA = {'DRAWING_PRESENT' : DRAWING_PRESENT, "SHAPE_FEATURES" : []}
        for DRAWING_FILE, SHAPES in DRAWDING.items():
            for SHAPE in SHAPES:
                SHAPE_DATA = {  'TYPE'  : SHAPE.get('TYPE', 'N/A'),
                                'NAME'  : SHAPE.get('NAME', 'N/A'),
                                'MACRO' : SHAPE.get('MACRO', 'N/A'),
                                'TEXT'  : SHAPE.get('TEXT', 'N/A')}
                
                DRAWING_DATA['SHAPE_FEATURES'].append(SHAPE_DATA)
        return DRAWING_DATA



    SHAPE_DATA  = EXTRACT_DRAWING(PARSED_DATA['DRAWINGS'])
    VBA_DATA    = {"CODE_PRESENT"  : (False if len(PARSED_DATA['VBA']) == 0 else True), "CODE"          : []}
    PWQ_DATA    = {"CODE_PRESENT"  : (False if len(PARSED_DATA['PWQ']) == 0 else True), "CODE"          : []}


    if VBA_DATA['CODE_PRESENT'] == True:
        for a in range(len(PARSED_DATA['VBA'])):
            VBA_DRILL = {"ORDER"            : "NA",
                         "LOCATION_TYPE"    : "SHEET" if "SHEET" in PARSED_DATA['VBA'].at[a, 'VBA_FILENAME'].upper() else "MODULE",
                         "LOCATION"         : PARSED_DATA['VBA'].at[a, 'VBA_FILENAME'],
                         "SCRIPT"           : PARSED_DATA['VBA'].at[a, 'VBA_CODE']}
        
            VBA_DATA['CODE'].append(VBA_DRILL)


    if PWQ_DATA['CODE_PRESENT'] == True:
        for a in range(len(PARSED_DATA['PWQ'])):
            PWQ_DRILL = {"ORDER"            : "NA",
                         "CONNECTION_NAME"  : PARSED_DATA['PWQ'].at[a, 'NAME'],
                         "REFERENCE_TABLE"  : PARSED_DATA['PWQ'].at[a, 'REFERENCE'],
                         "CONNECTION_TYPE"  : PARSED_DATA['PWQ'].at[a, 'TYPE'],
                         "SCRIPT"           : PARSED_DATA['PWQ'].at[a, 'M_CODE']}
        
            PWQ_DATA['CODE'].append(PWQ_DRILL)




    MODEL_DATA = ''


    DICT_EX['SHAPES']            = SHAPE_DATA
    DICT_EX['VBA']               = VBA_DATA
    DICT_EX['POWER_QUERY']       = PWQ_DATA
    DICT_EX['DATA_MODEL']        = MODEL_DATA
    DICT_EX['STYLE_USED']        = PARSED_DATA['STYLES']
    DICT_EX['THEME_USED']        = {'THEME_APPLIED' : PARSED_DATA['THEME']['color_scheme']}

    return DICT_EX




def CRITICAL_CELLS_SEARCH(GRAPH):
    CRITICAL_CELLS = sorted(GRAPH.in_degree, key=lambda x: x[1], reverse=True)
    return CRITICAL_CELLS












def XML_COMBINATION_DICT(PARSED_DATA, LABELS, OUTPUT_DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS, FILE_NAME, PROCESS_NAME, FREQUENCY, INSTANCE_DATE, INPUT_OPTIONS, INPUT_CHOICE, P3_RESULTS):

    SQL_INPUT_TYPE_VARIABLE     = "STILL TO BE DEVELOPED"
    EX_CON_INPUT_TYPE_VARIABLE  = "STILL TO BE DEVELOPED"
    PASTED_INPUT_TYPE_VARIABLE  = "STILL TO BE DEVELOPED"
    INPUT_SHEETS                = [KEY for KEY, VALUE in LABELS.items() if VALUE == 'INPUT_DATA']

    DESIRED_OUTPUT  = PROCESS_TABLES(PARSED_DATA, LABELS, OUTPUT_DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS, P3_RESULTS)
    DICT_EX         = OTHER_INFO(PARSED_DATA)
    CRITICAL_CELLS  = CRITICAL_CELLS_SEARCH(P3_RESULTS['DEPENDENCY_GRAPH'])



    FINAL_DICT = {"WORKBOOK"            :{  "FILENAME"      : FILE_NAME.split('.xl')[0],
                                            "PROCESS_NAME"  : PROCESS_NAME,
                                            "FREQUENCY"     : FREQUENCY,
                                            "INSTANCE_DATE" : INSTANCE_DATE},
                    "WORKSHEETS"        :DESIRED_OUTPUT,
                    "DEPENDENCY_GRAPH"  :CRITICAL_CELLS,
                    "CALCULATION_ORDER" :PARSED_DATA['CALC_CHAIN'],                         
                    "INPUT_TYPE"        :{  "INPUT_DATA_TYPE"       :INPUT_OPTIONS[INPUT_CHOICE],       
                                            "SQL_FUNCS_CALL"        :SQL_INPUT_TYPE_VARIABLE,
                                            "EXT_CON_LOCATIONS"     :EX_CON_INPUT_TYPE_VARIABLE,
                                            "PASTED_DATA_SOURCE"    :PASTED_INPUT_TYPE_VARIABLE,
                                            "INPUT_DATA_SHEETS"     :INPUT_SHEETS},
                    "SHAPES"            :DICT_EX['SHAPES'],            
                    "VBA"               :DICT_EX['VBA'],            
                    "POWER_QUERY"       :DICT_EX['POWER_QUERY'],       
                    "DATA_MODEL"        :DICT_EX['DATA_MODEL'],                
                    "STYLE_USED"        :DICT_EX['STYLE_USED'],            
                    "THEME_USED"        :DICT_EX['THEME_USED'],      
                }


    return FINAL_DICT, DESIRED_OUTPUT




######## STILL TO DO 
######## ---- DATA MODEL
