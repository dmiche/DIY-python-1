#!/usr/bin/python
# -*- coding: utf-8 -*-
import httplib2
#import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
#
import mysql.connector 
from mysql.connector import errorcode

#from __future__ import print_function
#import httplib2
import os
#import pyodbc
#conn = pyodbc.connect('DRIVER={MySQL};SERVER=192.168.1.5;DATABASE=Budjet;UID=root;PWD=rootpsw')
#conn = pyodbc.connect('DSN=mysql')

#try:
#  cnx = mysql.connector.connect(user='root', password='rootpsw',host='192.168.1.5',
#                                database='Budjet')
#except mysql.connector.Error as err:
#  if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
#    print("Something is wrong with your user name or password")
#  elif err.errno == errorcode.ER_BAD_DB_ERROR:
#    print("Database does not exist")
#  else:
#    print(err)
#else:
#  cnx.close()


from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
#from PythonApplication4 import pymssql


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json

SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:    
    https://docs.google.com/spreadsheets/d/1lzGFfzqAg5H8uoOOKJX_IwNtFvR6HsilHQM85_J9VmY/edit#gid=1656208107
    """
    input_list= [tuple(('','','','','','','','','','','','','','','','','',''))]
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1lzGFfzqAg5H8uoOOKJX_IwNtFvR6HsilHQM85_J9VmY'
    rangeName = 'План!A3:R8'
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])   
    if not values:
            print('No data found.')
    else:
            print('data + OK:')  

    for row in values:            
        #print('%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s' % (row[0], row[1], row[2],row[3], row[4], row[5],row[6], row[7], row[8],row[9], row[10], row[11],row[12], row[13], row[14],row[15], row[16], row[17]))        
        input_list.append((tuple((row[0], row[1], row[2],row[3], row[4], row[5],row[6], row[7], row[8],row[9], row[10], row[11],row[12], row[13], row[14],row[15], row[16], row[17]))))        
            
    rangeName = 'План!A11:R31'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])                            

    if not values:
        print('No data found.')
    else:
        print('data - OK:')    
            
    for row in values:            
        #print('%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s' % (row[0], row[1], row[2],row[3], row[4], row[5],row[6], row[7], row[8],row[9], row[10], row[11],row[12], row[13], row[14],row[15], row[16], row[17]))        
        input_list.append((tuple((row[0], row[1], row[2],row[3], row[4], row[5],row[6], row[7], row[8],row[9], row[10], row[11],row[12], row[13], row[14],row[15], row[16], row[17]))))
             
#############################################################
    try:
        cnx = mysql.connector.connect(user='root', password='rootpsw',host='192.168.1.5',
                                database='Budjet')
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("Database does not exist")
        else:
            print(err)
            exit(1)
    else:
        print('OK')
        

    cursor = cnx.cursor()
    try:
        cursor.execute("truncate table Bud")   

    except mysql.connector.Error as err:
            print("Failed truncate table Bud: {}".format(err))
            exit(1)            
    else:
        print("OK")
        
    
    try:                
        cursor.executemany("INSERT INTO Bud  VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", input_list )   
        #cursor.execute("INSERT INTO Bud  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", input_list )   
        #cursor.executemany("INSERT INTO Bud (ДОХОД, [Column 1], [Column 2], BASIS, [Column 4], Год, Январь, Февраль, Март, Апрель, Май, Июнь, Июль, Август, Сентябрь, Октябрь, Ноябрь, Декабрь) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", input_list )               

    except mysql.connector.Error as err:
            print("Failed INSERT INTO Bud: {}".format(err))
            exit(1)                                       
    else:
        print("OK")
        
    
    cnx.commit()
    cursor.close()
    cnx.close()


if __name__ == '__main__':
    main()
