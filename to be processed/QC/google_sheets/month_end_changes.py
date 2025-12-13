from google_api.google_sheets.general_sheets_functions import GET_SHEET_ID

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pathlib
import asyncio
from functools import partial

async def ASYNC_ADD_HEADING(value, spreadsheet_id, range_name, api_info):
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, partial(ADD_HEADING, value, spreadsheet_id, range_name, api_info))


def ADD_HEADING(values, spreadsheet_id, range_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    token_file_path = api_info["token_path"]
    credentials_path = api_info["credentials_path"]
    
    creds = Credentials.from_authorized_user_file(token_file_path, scopes)
    
    try:
        service = build("sheets", "v4", credentials=creds)
        
        body = {"values": [[values]]}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error
    

def ADD_TICKBOXES(row_number, spreadsheet_id, tab_name, api_info):
    row_number = int(row_number)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    token_file_path = api_info["token_path"]
    creds = Credentials.from_authorized_user_file(token_file_path, scopes)

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet_id = GET_SHEET_ID(service, spreadsheet_id, tab_name)
        # Tickbox data validation rule
        checkbox_rule = {
            "condition": {
                "type": "BOOLEAN"
            },
            "showCustomUi": True
        }

        # Apply to columns A and B (index 0 and 1), from row_number to bottom
        requests = [
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_number - 1,
                        "startColumnIndex": 8,
                        "endColumnIndex": 10  #I and J
                    },
                    "rule": checkbox_rule
                }
            }
        ]

        body = {
            "requests": requests
        }

        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        return response

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error