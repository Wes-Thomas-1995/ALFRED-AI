

from SQL import SQL_DATA_SAVE, SQL_DATA_READ, CREATE_TABLES, DELETE_TABLES






DETAILS                           = {"POSTGRES_HOSTNAME"    : 'ENKIINVESTMENTS-186.postgres.eu.pythonanywhere-services.com',
                                     "POSTGRES_HOST_PORT"   : 10186,
                                     "SSH_USER_NAME"        : 'ENKIINVESTMENTS',
                                     "SSH_PASSWORD"         : 'BRYONYBUTTON0104',
                                     "USERNAME"             : "super",
                                     "PASSWORD"             : 'BRYONYBUTTON0104',
                                     "DB_NAME"              : 'myappdb',
                                     "HOST"                 : '127.0.0.1'}

INPUT_DF            = ""
DROP_TABLE          = "algo_list"
TABLE_NAMES         = "algo_list"
TABLE_NAME          = "TEST_DATABASE"
INPUT_CRITERIA      = [ "TIME VARCHAR(255) NOT NULL",
                        "COIN VARCHAR(255) NOT NULL",
                        "LEVERAGE VARCHAR(255) NOT NULL",
                        "TP VARCHAR(255) NOT NULL",
                        "SL VARCHAR(255) NOT NULL",
                        "PORTFOLIO_PERCENT VARCHAR(255) NOT NULL",
                        "DIRECTION VARCHAR(255) NOT NULL",
                        "STATUS VARCHAR(255) NOT NULL"]






DELETE      = DELETE_TABLES(DETAILS, DROP_TABLE)
CREATE      = CREATE_TABLES(DETAILS, TABLE_NAME, INPUT_CRITERIA)

ERROR       = SQL_DATA_SAVE(INPUT_DF, TABLE_NAME, DETAILS)
DB          = SQL_DATA_READ(TABLE_NAMES, DETAILS)




