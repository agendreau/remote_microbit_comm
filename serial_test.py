import serial
import requests
import time
from googleapiclient.discovery import build
import pickle
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

port = "/dev/cu.usbmodem14102" #update to your serial port, to see serial ports use the command ls /dev/cu.* in the command line (for mac only)
conn = serial.Serial(port, 115200)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


sheet_address = "" #insert sheet address 

#Alex read/write
comm_sheet_read = {'b':'AlexMessages!A2:A2','g':'AlexLightControl!A2:A2','p':'AlexIcons!A2:A2','a':'AlexAcceleration!A2:A1000'}
comm_sheet_write = {'b':'LilaMessages!A2:A2','g':'LilaLightControl!A2:A2','p':'LilaIcons!A2:A2','a':'LilaAcceleration!A2:A1000'}

#Lila read/write
#comm_sheet_write = {'m':'AlexMessages!A2:A1000','lc':'AlexLightControl!A2:A1000','i':'AlexIcons!A2:A1000','accel':'AlexAcceleration!A2:A1000'}
#comm_sheet_read = {'m':'LilaMessages!A2:A1000','lc':'LilaLightControl!A2:A1000','i':'LilaIcons!A2:A1000','accel':'LilaAcceleration!A2:A1000'}



def setup():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

    

def write_to_sheet(service,data_range,message):
    print(message)
    #sheet = service.spreadsheets()
    #result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    #values = result.get('values', [])
    #total_values=len(values)
    values = [[message]]
    print(values)
    body = {'values': values}
    #data_range = data_range.split(':')[0]+':'+str(total_values+2)
    print('data range: %s',data_range)
    request = service.spreadsheets().values().update(spreadsheetId=sheet_address, range=data_range, valueInputOption='USER_ENTERED', body=body)
    response = request.execute()


def write_accel_values(service,sheet,data_range,message):
    print(message)
    #sheet = service.spreadsheets()
    #result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    #values = result.get('values', [])
    #total_values=len(values)
    values = [[message]]
    #print(values)
    body = {'values': values}
    #data_range = data_range.split(':')[0]+':'+str(total_values+2)
    #print('data range: %s',data_range)
    request = service.spreadsheets().values().append(spreadsheetId=sheet_address, range=data_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=body)
    response = request.execute()



def read_from_sheet(service,sheet,data_range):
    #sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    values = result.get('values', [])
    if len(values) > 0:
        #total_values=len(values)
        #clear = [[""]]
        '''
        for i in range(0,total_values):
            clear.append([""])
        '''
        body = {'values': [[""]]}
        #update_range = data_range.split(':')[0]+':'+str(total_values+2)
        result = service.spreadsheets().values().update(spreadsheetId=sheet_address, range=data_range,valueInputOption='USER_ENTERED', body=body).execute()
    return values

def read_accel_from_sheet(service,sheet,data_range):
    #sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    values = result.get('values', [])
    return values

def main():
    service = setup()
    sheet = service.spreadsheets()
    while True:
        if(conn.inWaiting()>0):
            print("Message from Micro:bit")
            #print(conn.readline().decode("utf-8").strip())

            m = conn.readline().decode("utf-8")
            print(m)
            decode_m = m.strip().split(',')
            print(decode_m)
            key = decode_m[0]
            value = decode_m[1]
            
            if key == 'a':
                data_range = comm_sheet_write[key]
                write_accel_values(service,sheet,data_range,value)
            elif key == 'ga':
                data_range=comm_sheet_read['a']
                values = read_accel_from_sheet(service,sheet,data_range)
                for v in values:
                    print(v[0])
                    result_bytes = ('ra' + ',' + str(v[0]) + "\n").encode()
                    conn.write(result_bytes)
                    time.sleep(0.1)
            elif key in ['b','g','p']:
                data_range = comm_sheet_write[key]
                write_to_sheet(service,data_range,value)

            else:
                print('key error: '+str(key))
            



        #check for messages
        #iterate over first three rows
        
        for key,value in comm_sheet_read.items():
            if key!='a':
                data_range = value
                values = read_from_sheet(service,sheet,data_range)
                for v in values:
                    result_bytes = ('r' + key+','+str(v[0]) + "\n").encode()
                    print(result_bytes.decode("utf-8"))
                    conn.write(result_bytes)
            #time.sleep(1)
            time.sleep(1)
        
        
        time.sleep(1)

if __name__ == '__main__':
    main()
  
  