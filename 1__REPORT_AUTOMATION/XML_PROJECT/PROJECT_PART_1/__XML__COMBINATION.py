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













def PROCESS_TABLES(PARSED_DATA, LABELS, DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS):

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
        START_COL_LETTER = START_CELL[0]
        START_COL_NUM = COL_TO_NUM(START_COL_LETTER)
        RETURN_COL_NUM = START_COL_NUM + COLUMN_INDEX
        RETURN_COL_LETTER = string.ascii_uppercase[RETURN_COL_NUM - 1]
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

    def POPULATE_CELLS(COLUMN_DICT, APPEND_DICT, UPDATED_WORKSHEETS, SHEET_NAME):
        COLUMN_LETTER = COLUMN_DICT['COLUMN_LETTER']
        for CELL_REF in APPEND_DICT['REF_CELLS']:
            CELL_COLUMN_LETTER = CELL_REF[0]
            if CELL_COLUMN_LETTER == COLUMN_LETTER:
                if CELL_REF in UPDATED_WORKSHEETS[SHEET_NAME]:
                    CELL_DATA = UPDATED_WORKSHEETS[SHEET_NAME][CELL_REF]
                    COLUMN_DICT['CELLS'][CELL_REF] = {
                        'TYPE': CELL_DATA.get('TYPE', None),
                        'STYLE_INDEX': CELL_DATA.get('STYLE_INDEX', None),
                        'FORMULA': CELL_DATA.get('FORMULA', None),
                        'VALUE': CELL_DATA.get('VALUE', None)
                    }

    WORKSHEETS = {}
    for SHEET_NAME, SHEET_DATA in PARSED_DATA['WORKSHEETS'].items():
        SHEET_INFO = {
            "SHEET_TYPE": LABELS.get(SHEET_NAME, "UNKNOWN"),
            "TABLE_PRESENT": False,
            "TABLE_PARAMS": {},
            "PIVOT_PRESENT": False,
            "PIVOT_PARAMS": {},
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
                    POPULATE_CELLS(COLUMN_DICT, APPEND_DICT, UPDATED_WORKSHEETS, SHEET_NAME)
                    APPEND_DICT['COLUMN'].append(COLUMN_DICT)

                SHEET_INFO["CLUSTERS"].append(APPEND_DICT)
                SHEET_INFO = CHECK_CLUSTER_RANGE(SHEET_INFO)

                for J in range(len(SHEET_INFO["CLUSTERS"])):
                    if SHEET_INFO["CLUSTERS"][J]['TABLE_RANGE']:
                        SHEET_INFO["CLUSTERS"][J]['TABLE_TYPE'] = "TABLE"
                    elif SHEET_INFO["CLUSTERS"][J]['PIVOT_RANGE'] == True: SHEET_INFO["CLUSTERS"][J]['TABLE_TYPE'] = "PIVOT"



        WORKSHEETS[SHEET_NAME] = SHEET_INFO

    return {"WORKSHEETS": WORKSHEETS}
























def XML_COMBINATION_DICT(PARSED_DATA, LABELS, OUTPUT_DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS, FILE_NAME, PROCESS_NAME, FREQUENCY, INSTANCE_DATE, INPUT_OPTIONS, INPUT_CHOICE):

    SQL_INPUT_TYPE_VARIABLE     = "STILL TO BE DEVELOPED"
    EX_CON_INPUT_TYPE_VARIABLE  = "STILL TO BE DEVELOPED"
    PASTED_INPUT_TYPE_VARIABLE  = "STILL TO BE DEVELOPED"
    INPUT_SHEETS                = [KEY for KEY, VALUE in LABELS.items() if VALUE == 'INPUT_DATA']

    DESIRED_OUTPUT = PROCESS_TABLES(PARSED_DATA, LABELS, OUTPUT_DETECTED_TABLES, CLUSTERED_NESTED_OUTPUT, NESTED_OUTPUT, UPDATED_WORKSHEETS)




    FINAL_DICT = {"WORKBOOK"            :{  "FILENAME"      : FILE_NAME.split('.xl')[0],
                                            "PROCESS_NAME"  : PROCESS_NAME,
                                            "FREQUENCY"     : FREQUENCY,
                                            "INSTANCE_DATE" : INSTANCE_DATE},
                    "WORKSHEETS"        :DESIRED_OUTPUT,
                    "DEPENDENCY_GRAPH"  :{},
                    "CALCULATION_ORDER" :{},                         
                    "INPUT_TYPE"        :{  "INPUT_DATA_TYPE"       :INPUT_OPTIONS[INPUT_CHOICE],       
                                            "SQL_FUNCS_CALL"        :SQL_INPUT_TYPE_VARIABLE,
                                            "EXT_CON_LOCATIONS"     :EX_CON_INPUT_TYPE_VARIABLE,
                                            "PASTED_DATA_SOURCE"    :PASTED_INPUT_TYPE_VARIABLE,
                                            "INPUT_DATA_SHEETS"     :INPUT_SHEETS},
                    "SHAPES"            :{},
                    "VBA"               :{},                ### DONE
                    "POWER_QUERY"       :{},                ### DONE
                    "DATA_MODEL"        :{},                
                    "EXTERNAL_CON"      :{},
                    "STYLE_USED"        :{},                ### DONE
                    "THEME_USED"        :{},                ### DONE
                }


    return FINAL_DICT