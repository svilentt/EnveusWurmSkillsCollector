from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xlsxwriter

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1eNE6a0kyGMvvTcmAkbIiFipYZacNyVstI5ndiqEbTOk'
SPREADSHEET_NAME = 'Skills Page'

SKILLS_STARTER_ROW = '11'
A1_FORMAT = '{}!{}:{}' # A1 notation https://developers.google.com/sheets/api/guides/concepts

def getSheet():
	"""Shows basic usage of the Sheets API.
	Prints values from a sample spreadsheet.
	"""
	creds = None
	# The file token.pickle stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)

	service = build('sheets', 'v4', credentials=creds)

	# Call the Sheets API
	return service.spreadsheets()
	
def updateSheet(playerName, playerData):
	sheet = getSheet()
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
								range=SPREADSHEET_NAME).execute()
	fetchedValues = result.get('values', [])
	
	playerColumn = ''
	for idx, name in enumerate(fetchedValues[4]):
		if(playerName == name):
			playerColumn = xlsxwriter.utility.xl_col_to_name(idx)
	if(playerColumn == ''):
		print('Player name not found on sheet. Please check, that your in game name and sheet name match exactly.')
		return False
	playerColumnRange = A1_FORMAT.format(SPREADSHEET_NAME, playerColumn + SKILLS_STARTER_ROW, playerColumn)

	values = [
		playerData
	]
	body = {
		'range': playerColumnRange,
		"majorDimension": "COLUMNS",
		'values': values
	}
	result = sheet.values().update(
		spreadsheetId=SPREADSHEET_ID, range=playerColumnRange,
		valueInputOption='USER_ENTERED', body=body).execute()
	
	return True