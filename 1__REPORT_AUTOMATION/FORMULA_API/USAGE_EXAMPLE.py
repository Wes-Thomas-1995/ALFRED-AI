import FORMULA_API 
import warnings
import json


warnings.filterwarnings("ignore")





EXCEL_TEST_PATH                     = r'/Users/westhomas/Desktop/ALFRED_AI/REFERENCE_FILES' + '/' + 'EXCL_TEST.xlsx' 
EXCEL_FUNCS_FOLDER                  = r'/Users/westhomas/Desktop/ALFRED_AI/FORMULA_API/REFERENCE_FILES'
EXCEL_FUNCS_EXCEL_PATH              = EXCEL_FUNCS_FOLDER + '/' +  'FUNCTIONS.xlsx'








REFERENCE_INFORMATION               = FORMULA_API.SOURCE_REFERENCE_INFORMATION(EXCEL_FUNCS_EXCEL_PATH)
TEST_INFORMATION                    = FORMULA_API.SOURCE_TEST_INFORMATION(EXCEL_TEST_PATH)




PYTHON                              = FORMULA_API.FORMULA_CONVERSION(TEST_INFORMATION.FORMULA,
                                                                    REFERENCE_INFORMATION.OPERATORS_DF,
                                                                    REFERENCE_INFORMATION.FUNCTIONS_DF,
                                                                    TEST_INFORMATION.DICT_WORKBOOK,
                                                                    EXCEL_FUNCS_FOLDER)




JSON                                = FORMULA_API.FORMULA_JSON(TEST_INFORMATION.FORMULA,
                                                            REFERENCE_INFORMATION.OPERATORS_DF,
                                                            REFERENCE_INFORMATION.FUNCTIONS_DF)






print(' ')
print('ORIGINAL FORMULA  :  ' + str(TEST_INFORMATION.FORMULA))
print(' ')
print('FINAL FORMULA     :  ' + str(PYTHON.PYTHON_SCRIPT))
print(' ')
print(' ')
print('JSON FORMAT OF OUR FORMULA  :')
print(json.dumps(JSON.JSON_FORMULA, indent=4))
print(' ')





