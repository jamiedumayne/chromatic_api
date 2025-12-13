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


def ATTACH_SCRIPT(spreadsheet_id, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    drive_service = build("drive", "v3", credentials=creds)
    script_service = build("script", "v1", credentials=creds)

    #create appscript project
    create_response = script_service.projects().create(body={
    "title": "Auto-Created Script Project"
    }).execute()

    script_id = create_response["scriptId"]

    test_code = 'console.log("test")'

    script_service.projects().updateContent(
    scriptId=script_id,
    body={
        "files": [
            {
                "name": "Code",
                "type": "SERVER_JS",
                "source": test_code
            },
            {
                "name": "appsscript",
                "type": "JSON",
                "source": '{"timeZone":"Europe/London","exceptionLogging":"STACKDRIVER"}'
            }
        ]
    }
    ).execute()

    # Step 3: Attach Script to Spreadsheet
    drive_service.files().update(
        fileId=script_id,
        addParents=spreadsheet_id
    ).execute()


def DELETE_REMAINING_COLUMNS(spreadsheet_id, tab_name, last_column, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID and total columns
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    total_columns = sheet['properties']['gridProperties']['columnCount']

    # Only delete if there are columns to remove
    if last_column + 1 < total_columns:
        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": last_column + 1,
                    "endIndex": total_columns
                }
            }
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [request]}
        ).execute()