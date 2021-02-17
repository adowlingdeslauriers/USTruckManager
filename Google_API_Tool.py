from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1GOTmc0vbu4tHtRGkqkgrFGxSazV2XK3fhzcE42qS4mE"
SAMPLE_RANGE_NAME = "MAIN!A2"

def main(in_spreadsheet_id = SAMPLE_SPREADSHEET_ID, in_range = SAMPLE_RANGE_NAME, fedex_in = [], dhl_in = []):
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
    sheet = service.spreadsheets()

    values = []
    for i in range(max(len(dhl_in), len(fedex_in))):
    	try:
    		dhl_gaylord = dhl_in[i]
    	except:
    		dhl_gaylord = None
    	try:
    		fedex_gaylord = fedex_in[i]
    	except:
    		fedex_gaylord = None
    	
    	values.append([fedex_gaylord, dhl_gaylord])

    #values = [
    #	["TEST1", "TEST2"]
    #]
    body = {
    	"values": values
    }
    result = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A1", valueInputOption = "RAW", body = body).execute()
    return result
    #result = sheet.values().get(spreadsheetId=in_spreadsheet_id,range=in_range).execute()
    
    #values = result.get('values', [])
    #return values
    '''
    if not values:
        print('No data found.')
    else:
        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print(row)
    '''

if __name__ == '__main__':
    main()