

from SQL import SQL_DATA_SAVE, SQL_DATA_READ, CREATE_TABLES, DELETE_TABLES, LOAD_TABLES, DROP_COLUMN


import pandas as pd
import warnings

warnings.filterwarnings("ignore")




def READ_EXCEL(FILE_PATH):

    FUNCTIONS, TABLE_NAMES                          = [], []
    REFERENCE                                       = pd.read_excel(FILE_PATH, sheet_name=0)

    for i in range(1,9):                           FUNCTIONS.append(pd.read_excel(FILE_PATH, sheet_name=i))
    for a in range(len(REFERENCE['TABLE NAME'])):   TABLE_NAMES.append(REFERENCE.at[a, 'TABLE NAME'])

    return TABLE_NAMES, FUNCTIONS









DETAILS                           = {"POSTGRES_HOSTNAME"    : 'ENKIINVESTMENTS-186.postgres.eu.pythonanywhere-services.com',
                                     "POSTGRES_HOST_PORT"   : 10186,
                                     "SSH_USER_NAME"        : 'ENKIINVESTMENTS',
                                     "SSH_PASSWORD"         : 'BRYONYBUTTON0104',
                                     "USERNAME"             : "super",
                                     "PASSWORD"             : 'BRYONYBUTTON0104',
                                     "DB_NAME"              : 'myappdb',
                                     "HOST"                 : '127.0.0.1'}









INPUT_CRITERIA                      = ["TEST VARCHAR(255) NOT NULL"]
EXCEL_FUNCS_FOLDER                  = r'/Users/westhomas/Desktop/ALFRED/1__REPORT_AUTOMATION/REFERENCE_FILES/UPLOAD_TABLES/'
FILE_EXTRA                          = "AST"
EXCEL_FILE_NAME                     = 'AST_RAW_DATA_SQL_UPLOAD'


EXCEL_TEST_PATH                     = EXCEL_FUNCS_FOLDER + '/' + FILE_EXTRA + '/' + EXCEL_FILE_NAME + '.xlsx'
TABLE_NAMES, FUNCTIONS              = READ_EXCEL(EXCEL_TEST_PATH)


TABLE                               = 'tvfgetdt_p_investmentinitialinformation'







#TABLES                              = LOAD_TABLES(DETAILS)
#for DROP in TABLES.TABLES:          DELETE = DELETE_TABLES(DETAILS, DROP)
#for NAME in TABLE_NAMES:            CREATE = CREATE_TABLES(DETAILS, NAME)
#DB                                  = SQL_DATA_READ(TABLE, DETAILS)
#for u in range(len(TABLE_NAMES)):   ERROR = SQL_DATA_SAVE(FUNCTIONS[u], TABLE_NAMES[u], DETAILS)
#for u in range(len(TABLE_NAMES)):   ERROR = DROP_COLUMN(DETAILS, TABLE_NAMES[u], "TEST")