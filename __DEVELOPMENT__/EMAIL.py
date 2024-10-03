
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











class EMAIL():
    def __init__(self, INFO_BUNDLE, ROOT):
        self.INFO_BUNDLE    = INFO_BUNDLE
        self.ROOT           = ROOT
        self.INFO_BUNDLE    = self.EMAIL()

    def EMAIL(self):

        DIRECTION_USE = "SHORT" if self.INFO_BUNDLE[0] == "DOWN" else "LONG"


        mail_content = '''Good Morning Wes,
    We are reccomending you trade {} on {} today. Our confidence level is : {}. Here are some entry information -
    -    Entry Price : {}
    -    0.3% TP and SL : {} TP    -    {} SL
    -    0.5% TP and SL : {} TP    -    {} SL
    -    1.0% TP and SL : {} TP    -    {} SL

    Total portfolio value is - {}

    Kind Regards,
    ENKI Technology Ltd.

        ''' .format(DIRECTION_USE, self.INFO_BUNDLE[1], self.INFO_BUNDLE[2], self.INFO_BUNDLE[3], self.INFO_BUNDLE[4], self.INFO_BUNDLE[5], self.INFO_BUNDLE[6], self.INFO_BUNDLE[7], self.INFO_BUNDLE[8], self.INFO_BUNDLE[9], self.INFO_BUNDLE[10])

        #The mail addresses and password
        sender_address          = 'enki.technologies.ltd@gmail.com'
        sender_pass             = 'ojsfpzrgnneixdeq'
        receiver                = ['wesley.thomas95@gmail.com']
        email_subject           = self.INFO_BUNDLE[1] + ' - ' + DIRECTION_USE + ' - ' + self.INFO_BUNDLE[2] + ' Priority - Trade Reccomendation - ' + self.INFO_BUNDLE[11]
        filename_1                = self.INFO_BUNDLE[1] + '_PLOT_DISTANCE_VIEW_' + self.INFO_BUNDLE[11] + '.png'
        filename_2                = self.INFO_BUNDLE[1] + '_PLOT_NEAR_' + self.INFO_BUNDLE[11] + '.png'

        os.chdir(self.ROOT)

        for y in range(len(receiver)):
            receiver_address        = receiver[y]
            message                 = MIMEMultipart()
            message['From']         = sender_address
            message['To']           = receiver_address
            message['Subject']      = email_subject

            with open(filename_1, "rb") as attachment:
                part_1 = MIMEBase("application", "octet-stream")
                part_1.set_payload(attachment.read())

            with open(filename_2, "rb") as attachment:
                part_2 = MIMEBase("application", "octet-stream")
                part_2.set_payload(attachment.read())


            encoders.encode_base64(part_1)
            encoders.encode_base64(part_2)
            part_1.add_header("Content-Disposition", f"attachment; filename= {filename_1}",)
            part_2.add_header("Content-Disposition", f"attachment; filename= {filename_2}",)
            message.attach(MIMEText(mail_content, 'plain'))
            message.attach(part_1)
            message.attach(part_2)




            #Create SMTP session for sending the mail
            session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
            session.starttls() #enable security
            session.login(sender_address, sender_pass) #login with mail_id and password
            text = message.as_string()
            session.sendmail(sender_address, receiver_address, text)
            session.quit()


        return self.INFO_BUNDLE

