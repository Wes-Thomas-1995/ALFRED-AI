from AIDA import AIDA_DATA_EXTRACT



INPUT_DATA_REQUIREMENTS = { "INPUT_1" : {"REPORT"    :   "tvfGetDT_P_Clientreporting_PortfolioPositions_SPAR",
                                         "PID"       :   128,
                                         "DATE"      :   "31/03/2024"},
                            "INPUT_2" : {"REPORT"    :   "tvfGetDT_P_Clientreporting_PortfolioPositions_SPAR",
                                         "PID"       :   224,
                                         "DATE"      :   "31/03/2024"}
                            }




DF_LIST = AIDA_DATA_EXTRACT(INPUT_DATA_REQUIREMENTS)

print(len(DF_LIST))
print(DF_LIST[0].head())