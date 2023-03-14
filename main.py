from __future__ import print_function
import collections

import os.path
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1uY8Emah4pRujX4hcfmcgG9rkj-MIGqisutmdZJ5IEZc'

opts = Options()
# opts.set_headless()
browser = Chrome(options=opts)

rows_of_interest = {
    "contratto", "tipologia", "superficie", "locali", "piano", "totale piani edificio", "altre caratteristiche", "prezzo", "spese condominio", "anno di costruzione", "stato", "riscaldamento", "climatizzatore", "efficienza energetica"
}


def fetch_info_from(link: str):
    print("NEW LINK TO EXPLORE:", link)
    result = collections.defaultdict(lambda: None)
    try:
        browser.get(link)

        other_infos = browser.find_elements(By.CSS_SELECTOR,
                                            'dl.in-realEstateFeatures__list')
        for i, info in enumerate(other_infos):
            keys = info.find_elements(
                By.TAG_NAME, 'dt')
            values = info.find_elements(
                By.TAG_NAME, 'dd')
            for i, key in enumerate(keys):
                if key.text.lower() in rows_of_interest:
                    result[key.text.lower()] = values[i].text
        return result
    except Exception as e:
        print(e)
        return collections.defaultdict(lambda: None)


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        links = []
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        RANGE_NAME = 'March 2023!A2:B'
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        for row in values:
            if len(row) == 1:
                links.append(row[0])

        rows_to_write = []
        rows_to_write.append(list(rows_of_interest))
        for link in links:
            info = fetch_info_from(link)
            row = []
            for col in list(rows_of_interest):
                row.append(info[col])
            rows_to_write.append(row)
        RANGE_NAME = 'March 2023!B:Z'
        VALUE_INPUT_OPTION = 'USER_ENTERED'
        body = {
            'values': rows_to_write
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
            valueInputOption=VALUE_INPUT_OPTION, body=body).execute()

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
