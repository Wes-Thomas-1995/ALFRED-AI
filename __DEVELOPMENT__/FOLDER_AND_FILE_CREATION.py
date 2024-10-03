


from time import strftime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mplfinance as mpf
from matplotlib.gridspec import GridSpec
from binance.helpers import round_step_size
from binance.client import Client
from binance.enums import *

from datetime import datetime, timedelta, date, time
import pandas as pd
import numpy as np
import warnings
import os.path
import pytz
import time
import os
import ta


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import email, smtplib, ssl
import os



warnings.filterwarnings("ignore")






DATE_INTEREST               = [(datetime.today()).strftime('%Y-%m-%d')]
FOLDER_PATH                 = r'/home/ENKIINVESTMENTS/CRYPTO_TRADER/ALGO_TRADING_3/IMAGES/' + str(DATE_INTEREST[0])
FILE_LIST                   = ['CONTROLLER', 'DATA_MANIPULATION', 'DATA_SOURCING', 'FORMATTING', 'EMAIL']





class FOLDERS_AND_FILES():
    def __init__(self, DATE_INTEREST, FOLDER_PATH, FILE_LIST):
        self.DATE_INTEREST          = DATE_INTEREST
        self.FOLDER_PATH            = FOLDER_PATH
        self.FILE_LIST              = FILE_LIST
        self.FILE_LIST            = self.FOLDERS_AND_FILES()

    def FOLDERS_AND_FILES(self):



        if not os.path.exists(self.FOLDER_PATH):
            os.makedirs(self.FOLDER_PATH)

            for h in range(len(self.FILE_LIST)):
                FILE_PATH = self.FOLDER_PATH + '/' + self.FILE_LIST[h] + '.py'
                with open(FILE_PATH, 'w') as file:  file.write("# This is a placeholder for the Python script created on " + str(DATE_INTEREST[0]))


        else:   print(f"This folder path already exists.")

    
        return self.FILE_LIST