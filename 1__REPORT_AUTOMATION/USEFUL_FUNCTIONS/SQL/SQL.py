



from datetime import datetime, timedelta, date
from sqlalchemy import create_engine
import psycopg2 as pg
import pandas as pd
import sshtunnel
import warnings
import time
import os
import psycopg2

warnings.filterwarnings("ignore")


sshtunnel.SSH_TIMEOUT       = 5.0
sshtunnel.TUNNEL_TIMEOUT    = 5.0









class DELETE_TABLES():
    def __init__(self, DETAILS, DROP_TABLE):
        self.DETAILS                    = DETAILS
        self.DROP_TABLE                 = DROP_TABLE
        self.SQL                        = self.DELETE_TABLES()

    def DELETE_TABLES(self):
        with sshtunnel.SSHTunnelForwarder(
                    ('ssh.eu.pythonanywhere.com'),
                    ssh_username        = self.DETAILS['SSH_USER_NAME'],
                    ssh_password        = self.DETAILS['SSH_PASSWORD'],
                    remote_bind_address = (self.DETAILS['POSTGRES_HOSTNAME'], self.DETAILS['POSTGRES_HOST_PORT'])) as server:
            
            server.start()
            PORT                = server.local_bind_port
            params              = { 'database'  : self.DETAILS['DB_NAME'],
                                    'user'      : self.DETAILS['USERNAME'],
                                    'password'  : self.DETAILS['PASSWORD'],
                                    'host'      : self.DETAILS['HOST'],
                                    'port'      : PORT}

            conn                = psycopg2.connect(**params)
            cursor              = conn.cursor()
            SQL                 = f'''DROP TABLE {self.DROP_TABLE} '''

            cursor.execute(SQL)
            conn.commit() 
            conn.close()

        return SQL









class CREATE_TABLES():

    def __init__(self, DETAILS, TABLE_NAME, INPUT_CRITERIA):
        self.DETAILS                    = DETAILS
        self.TABLE_NAME                 = TABLE_NAME
        self.INPUT_CRITERIA             = INPUT_CRITERIA
        self.SQL                        = self.CREATE_TABLES()


    def CREATE_TABLES(self):

        with sshtunnel.SSHTunnelForwarder(
                    ('ssh.eu.pythonanywhere.com'),
                    ssh_username        = self.DETAILS['SSH_USER_NAME'],
                    ssh_password        = self.DETAILS['SSH_PASSWORD'],
                    remote_bind_address = (self.DETAILS['POSTGRES_HOSTNAME'], self.DETAILS['POSTGRES_HOST_PORT'])) as server:
            
            server.start()

            PORT                = server.local_bind_port
            params              = { 'database'  : self.DETAILS['DB_NAME'],
                                    'user'      : self.DETAILS['USERNAME'],
                                    'password'  : self.DETAILS['PASSWORD'],
                                    'host'      : self.DETAILS['HOST'],
                                    'port'      : PORT
                                }

            conn                = psycopg2.connect(**params)
            cursor              = conn.cursor()
            COLUMNS             = ",\n".join(self.INPUT_CRITERIA)

            SQL                 = F"""
                                CREATE TABLE {self.TABLE_NAME} (
                                {COLUMNS}
                                )"""

            
            cursor.execute(SQL)
            conn.commit() 
            conn.close()
            print('all wrapped up')

        return SQL









class SQL_DATA_READ():

    def __init__(self, TABLE_NAME, DETAILS):
        self.TABLE_NAME                 = TABLE_NAME
        self.DETAILS                    = DETAILS
        self.DATA                       = self.SQL_DATA_READ()


    def SQL_DATA_READ(self):

        with sshtunnel.SSHTunnelForwarder(
                    ('ssh.eu.pythonanywhere.com'),
                    ssh_username        = self.DETAILS['SSH_USER_NAME'],
                    ssh_password        = self.DETAILS['SSH_PASSWORD'],
                    remote_bind_address = (self.DETAILS['POSTGRES_HOSTNAME'], self.DETAILS['POSTGRES_HOST_PORT'])

            ) as tunnel:
                try:
                        PORT                = tunnel.local_bind_port
                        ENGINE              = create_engine(f"postgresql://{self.DETAILS['USERNAME']}:{self.DETAILS['PASSWORD']}@{self.DETAILS['HOST']}:{PORT}/{self.DETAILS['DB_NAME']}")
                        DATA_DETAIL         = pd.read_sql_table(self.TABLE_NAME, ENGINE)

                except (Exception,  pg.DatabaseError) as error: print(error)


        return DATA_DETAIL











class SQL_DATA_SAVE():

    def __init__(self, INPUT_DF, TABLE_NAME, DETAILS):
        self.INPUT_DF               = INPUT_DF
        self.TABLE_NAME             = TABLE_NAME
        self.DETAILS                = DETAILS
        self.ERROR_LIST             = self.SQL_DATA_SAVE()


    def SQL_DATA_SAVE(self):

        ERROR_LIST                  = []

        with sshtunnel.SSHTunnelForwarder(
                    ('ssh.eu.pythonanywhere.com'),
                    ssh_username        = self.DETAILS['SSH_USER_NAME'],
                    ssh_password        = self.DETAILS['SSH_PASSWORD'],
                    remote_bind_address = (self.DETAILS['POSTGRES_HOSTNAME'], self.DETAILS['POSTGRES_HOST_PORT'])

            ) as tunnel:
            try:

                PORT                = tunnel.local_bind_port
                ENGINE              = create_engine(f"postgresql://{self.DETAILS['USERNAME']}:{self.DETAILS['PASSWORD']}@{self.DETAILS['HOST']}:{PORT}/{self.DETAILS['DB_NAME']}")

                if len(self.INPUT_DF) > 0:
                    try:            self.INPUT_DF.to_sql(self.TABLE_NAME, ENGINE, if_exists='replace', index=False)
                    except:         print('UNABLE TO UPLOAD DETAIL')


            except (Exception,  pg.DatabaseError) as error:
                        ERROR_LIST.append(error)

        return ERROR_LIST










