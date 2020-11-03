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


sheet_address = #insert sheet address 

#Alex read/write
comm_sheet_read = {'m':'AlexMessages!A2:A1000','lc':'AlexLightControl!A2:A1000','i':'AlexIcons!A2:A1000','accel':'AlexAcceleration!A2:A1000'}
comm_sheet_write = {'m':'LilaMessages!A2:A1000','lc':'LilaLightControl!A2:A1000','i':'LilaIcons!A2:A1000','accel':'LilaAcceleration!A2:A1000'}

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
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    values = result.get('values', [])
    total_values=len(values)
    values = [[message]]
    print(values)
    body = {'values': values}
    data_range = data_range.split(':')[0]+':'+str(total_values+2)
    print('data range: %s',data_range)
    request = service.spreadsheets().values().append(spreadsheetId=sheet_address, range=data_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=body)
    response = request.execute()


def read_from_sheet(service,data_range):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
    values = result.get('values', [])
    if len(values) > 0:
        total_values=len(values)
        clear = []
        for i in range(0,total_values):
            clear.append([""])
        body = {'values': clear}
        update_range = data_range.split(':')[0]+':'+str(total_values+2)
        result = service.spreadsheets().values().update(spreadsheetId=sheet_address, range=update_range,valueInputOption='USER_ENTERED', body=body).execute()
    return values

def main():
    service = setup()
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
            
            if key != 'accel':
                data_range = comm_sheet_write[key]
                write_to_sheet(service,data_range,value)
            else:
                data_range=comm_sheet_read['accel']
                values = read_from_sheet(service,data_range)
                for v in values:
                    result_bytes = (str(v) + "\n".encode)
                    conn.write(result_bytes)
            



        #check for messages
        #iterate over first three rows
        
        for key,value in comm_sheet_read.items():
            data_range = value
            values = read_from_sheet(service,data_range)
            for v in values:
                result_bytes = (key+','+str(v[0]) + "\n").encode()
                print(result_bytes)
                conn.write(result_bytes)
                time.sleep(1)
        

        time.sleep(0.01)

if __name__ == '__main__':
    main()
  
  