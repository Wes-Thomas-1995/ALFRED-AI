import re







def SHEET_INDEXING(SHEET_NAME, DICT_INPUT):
    for index, NAME in DICT_INPUT['SHEETS'].items():
        if NAME == SHEET_NAME:
            return index
    raise ValueError(f"Sheet name '{SHEET_NAME}' not found in DICT_WORKBOOK['SHEETS']")




def HARDCODED(string):
    return f"{string}"




def VLOOKUP(FEATURES_DICT, WORKBOOK_DICT):

    LOOKUP_VALUE            = FEATURES_DICT['LOOKUP_VALUE']
    TABLE_ARRAY             = FEATURES_DICT['TABLE_ARRAY']
    COL_INDEX_NUM           = FEATURES_DICT['COL_INDEX_NUM']
    RANGE_LOOKUP            = FEATURES_DICT['RANGE_LOOKUP']

    SHEET_NAME, COL_RANGE   = TABLE_ARRAY.split('!')
    SHEET_INDEX             = SHEET_INDEXING(SHEET_NAME, WORKBOOK_DICT)
    
    START_COL, END_COL      = COL_RANGE.split(':')
    COL_MAP_DF              = WORKBOOK_DICT['COLUMN_MAP'][SHEET_INDEX]
    ALL_COL_RANGE           = COL_MAP_DF.set_index('LETTER').loc[START_COL:END_COL, 'HEADER'].values.tolist()
    END_COL_NAME            = ALL_COL_RANGE[(int(COL_INDEX_NUM) - 1)]
    START_COL_NAME          = ALL_COL_RANGE[0]
    

    return f"df.merge(WORKBOOK_DICT['DATAFRAME'][{SHEET_INDEX}][['{START_COL_NAME}', '{END_COL_NAME}']], left_on={LOOKUP_VALUE}, right_on='{START_COL_NAME}', how='inner').drop(columns='{START_COL_NAME}')"





def IF(FEATURES_DICT, WORKBOOK_DICT):

    CONDITION           = FEATURES_DICT['CONDITION']
    TRUE_VALUE          = FEATURES_DICT['TRUE']
    FALSE_VALUE         = FEATURES_DICT['FALSE']

    return f"np.where({CONDITION}, {TRUE_VALUE}, {FALSE_VALUE})"







