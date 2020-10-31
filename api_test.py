import requests
import serial
import time
from googleapiclient.discovery import build
import pickle
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


sheet_address = ""
data_range = 'Sheet1!A2:A3'

values = [[""],[""]]
body = {'values': values}


def main():
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

	service = build('sheets', 'v4', credentials=creds)

	sheet = service.spreadsheets()
	result = sheet.values().get(spreadsheetId=sheet_address,range=data_range).execute()
	values = result.get('values', [])
	if not values:
		print('No data found.')
	else:
		for row in values:
			print('%s' % (row[0]))

	result = service.spreadsheets().values().update(spreadsheetId=sheet_address, range=data_range,valueInputOption='USER_ENTERED', body=body).execute()
	print('{0} cells updated.'.format(result.get('updatedCells')))

if __name__ == '__main__':
	main()