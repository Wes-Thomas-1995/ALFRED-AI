import xlwings as xw
import pandas as pd





def AIDA_DATA_EXTRACT(INPUT_DATA_REQUIREMENTS):

    DF_LIST                 = []
    PATH                    = r"C:/Users\wesley.thomas/OneDrive - STEPSTONE GROUP LP/Desktop/SQL_TEST/DTP.xlsm"
    EXCEL_APP               = xw.App(visible=False)
    EXCEL_BOOK              = EXCEL_APP.books.open(PATH)

    try:
        if len(INPUT_DATA_REQUIREMENTS) > 1:
            for INDEX, KEY in INPUT_DATA_REQUIREMENTS.items():
                DF_LIST.append(INDEX)

                USAGE               = "INPUT_" + str(int(INDEX[-1])-1)
                try:                TEST = EXCEL_BOOK.sheets(INDEX).activate()
                except:             EXCEL_BOOK.sheets.add(name=INDEX, after=USAGE)



        LIST_OF_SHEETS          = [s.name for s in EXCEL_BOOK.sheets]
        DTP_INDEX               = LIST_OF_SHEETS.index("DT_P")


        for INDEX, KEY in INPUT_DATA_REQUIREMENTS.items():
            SHEETS                      = EXCEL_BOOK.sheets[DTP_INDEX]
            KEYS                        = list(INPUT_DATA_REQUIREMENTS.keys())
            INDEXS                      = KEYS.index(INDEX)
            ROW                         = 2 + INDEXS

            SHEETS.range(ROW, 1).value  = INDEX + "!A1"
            SHEETS.range(ROW, 2).value  = 1
            SHEETS.range(ROW, 3).value  = 1000
            SHEETS.range(ROW, 4).value  = 100
            SHEETS.range(ROW, 5).value  = KEY['REPORT']
            SHEETS.range(ROW, 6).value  = KEY['PID']
            SHEETS.range(ROW, 7).value  = KEY['DATE']


        EXCEL_BOOK.sheets[DTP_INDEX].select()


        MACRO               = EXCEL_BOOK.macro("Sheet7.Evaluation_template")
        MACRO()


        for A in range(len(DF_LIST)):
            DTP_INDEX               = LIST_OF_SHEETS.index(DF_LIST[A])
            DF_LIST[A]          = pd.DataFrame(EXCEL_BOOK.sheets[DTP_INDEX].used_range.value)
            DF_LIST[A].columns  = DF_LIST[A].iloc[0]
            DF_LIST[A]          = DF_LIST[A][1:]


        EXCEL_BOOK.close()

    except: 
        DF_LIST = "FAILED - UNABLE TO PROCESS"
        EXCEL_BOOK.close()
        
    return DF_LIST



