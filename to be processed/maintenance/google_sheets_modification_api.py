import os.path
import os
import pandas as pd
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

def GET_SHEET_METADATA(spreadsheet_id, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Fetch the spreadsheet details
    result = sheet.get(spreadsheetId=spreadsheet_id).execute()
    
    # Extract sheet names
    sheets = result.get("sheets", [])
    sheet_names = [s["properties"]["title"] for s in sheets]
    filename = result.get("properties", {}).get("title", "Unknown Spreadsheet")

    return sheet_names, filename


def CREATE_TAB(spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    # Request to add a new sheet named "test"
    add_sheet_request = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": tab_name  # Name of the new sheet
                    }
                }
            }
        ]
    }

    # Execute the request to add the sheet
    try:
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=add_sheet_request
        ).execute()

        print("Sheet 'test' created successfully!")

    except Exception as e:
        print("------------------")
        print("Error: Could not create the sheet 'test'")
        print("Exception:", e)
        print("------------------")
        sys.exit()


def ADD_TICKBOX(spreadsheet_id, tab_name, col_num, row_num, num_items, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)

    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    requests = [
    {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_num,
                "endRowIndex": row_num + 1,
                "startColumnIndex": col_num,
                "endColumnIndex": col_num + num_items
            },
            "rule": {
                "condition": {
                    "type": "BOOLEAN"
                },
                "showCustomUi": True
            }
        }
    }
    ]
    
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()