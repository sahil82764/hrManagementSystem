import pyodbc
import logging


serverName = 'LAPTOP-NFRNM2TK'
databaseName = 'igl_client'
driver = '{ODBC Driver 17 for SQL Server}'

class Database:
    def __init__(self):
        # self.serverName = serverName
        # self.databaseName = databaseName
        # self.driver = driver
        pass

    def connectSQL():
        try:
            connectionString = f"""Driver={driver};
                                   Server={serverName};
                                   Database={databaseName};
                                   Trusted_Connection=yes;
                                """
            dbConnection = pyodbc.connect(connectionString)
            return dbConnection
        except Exception as e:
            logging.error("Error while connecting to SQL Server: ", e)
