import config
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account



class GoogleSheetHandler:

    creds = None
    creds = service_account.Credentials.from_service_account_file(config.SERVICE_ACCOUNT_FILE, scopes = config.SCOPES)

    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    def __init__(self, data=None, sheet_name=None):
        self.data = data
        self.sheet_name = sheet_name

    def get_user_password(self):

        """Fetching Username & Password """
        result = self.sheet.values().get(spreadsheetId = config.SAMPLE_SPREADSHEET_ID, range ="Users!A1:B3").execute()
        get_values = result.get('values' , [])
        print('Username & Password Fetched Successfully!')
        return get_values
    
    def getsheet_records(self):
        
        """ Fetching the records from Google Sheet """
        
        result = self.sheet.values().get(spreadsheetId = config.SAMPLE_SPREADSHEET_ID,
                                    range = self.sheet_name).execute()
        get_values = result.get('values', [])
        print(f"GoogleSheet[{self.sheet_name}]: Records Fetched Successfully")
        return get_values

    def updatesheet_records(self, data):
        
        """ Updating the record in Google Sheet """
       
        records_to_update = self.data
        request = self.sheet.values().update(spreadsheetId = config.SAMPLE_SPREADSHEET_ID, range= self.sheet_name, 
        valueInputOption="USER_ENTERED", body={"values":records_to_update}).execute()
        print('Records Updated Successfully!')
        return request

    def appendsheet_records(self):
        
        """ Appending/Inserting record in Google Sheet """

        request = self.sheet.values().append(spreadsheetId = config.SAMPLE_SPREADSHEET_ID, range= self.sheet_name, 
            valueInputOption="USER_ENTERED", body={"values":self.data}).execute()
        
        print("\n\t\t\tRecord Inserted Successfully!")
        return request

    def clearsheet_records(self):
        
        """ Clearing records from Google Sheet """
        request = self.sheet.values().clear(spreadsheetId = config.SAMPLE_SPREADSHEET_ID, range="Sheet1!A3:C9").execute()
        print("Records Cleared Successfully!")
        return request


