import pyodbc
import logging


serverName = 'LAPTOP-NFRNM2TK'
databaseName = 'igl_client'
driver = '{ODBC Driver 17 for SQL Server}'

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

def get_vendors_only():
    databaseConnection = connectSQL()
    dbCursor = databaseConnection.cursor()
    dbCursor.execute('SELECT DISTINCT Vendor_Name FROM vendor')
    results = dbCursor.fetchall()
    vendorNames = [row[0] for row in results]
    dbCursor.close()
    databaseConnection.close()
    return vendorNames

def get_vendor_data(vendorName):
    try:
        databaseConnection = connectSQL()
        dbCursor = databaseConnection.cursor()
        dbCursor.execute("SELECT * FROM vendor where Vendor_Name = ?", vendorName)
        result_vendor = dbCursor.fetchall()
        dbCursor.close()
        databaseConnection.close()

        vendorDictionary = {}
        vendorKeys = ['PO No', 'Vendor Code', 'Vendor Name', 'Station Name', 'Contract Date', 'GST No', 'PAN No', 'Operator Name']

        # remove auto incremental id from result_vendor
        result_vendor = [list(t)[1:] for t in result_vendor]

        # assigning values from fetched sql query to dictionary
        vendorDictionary = {vendorKeys[i]: str(result_vendor[0][i]) for i in range(len(vendorKeys)) if i != 3}

        # checking for multiple station name
        vendorDictionary['Station Name'] = [result_vendor[i][3] for i in range(len(result_vendor))]

        #return dictionary
        return vendorDictionary

    except Exception as e:
        print(e)
