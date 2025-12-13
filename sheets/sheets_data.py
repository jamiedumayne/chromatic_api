import os.path
import os
import pandas as pd
import sys
import pathlib

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

def download_values(spreadsheet_id, range_request, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token_path"]
    credentials_path = api_info["credentials_path"]
    
    token_path = pathlib.Path(token_file_path)

    if os.path.exists(token_file_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
        creds = flow.run_local_server(port=0)
    with open(token_file_path, "w") as token:
        token.write(creds.to_json())

    creds = Credentials.from_authorized_user_file(token_path, scopes)
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    #search for spreadsheet and requested range
    result = (sheet.values().get(spreadsheetId=spreadsheet_id, range=range_request)
            .execute())
    values = result.get("values", [])
    try:
        result = (sheet.values().get(spreadsheetId=spreadsheet_id, range=range_request)
            .execute())
        values = result.get("values", [])
    except:
        print("------------------")
        print("Error: Can't download data from sheet")
        print("Check tab name is correct")
        sys.exit()

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=spreadsheet_id,
                                  range=range_request).execute()
        data = rows.get('values')
        df = pd.DataFrame(data[0:])

        return df


def download_formulas(sheet_info, tab_name):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = sheet_info["api_info"]["token"]
    credentials_path = sheet_info["api_info"]["credentials_path"]

    if os.path.exists(token_file_path):
        creds = Credentials.from_authorized_user_file(token_file_path, scopes)

    service = build("sheets", "v4", credentials=creds)
    sheets = service.spreadsheets().values()
    full_range = tab_name

    params = {
        "spreadsheetId": sheet_info["spreadsheet_id"],
        "range": full_range,
    }
    params["valueRenderOption"] = "FORMULA"

    try:
        result = sheets.get(**params).execute()
        data = result.get("values", [])

        formula_df = pd.DataFrame(data)
        formula_df.columns = formula_df.iloc[0]
        formula_df = formula_df[1:]
    except Exception as e:
        print("---------------")
        print("Error: can't download data from sheet")
        print(e)
        print("---------------")
        sys.exit(1)

    if not data:
        print("No data found.")
        return pd.DataFrame()

    return formula_df


def upload_values(spreadsheet_id, range_name, values, value_input_option, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    token_file_path = api_info["token_path"]
    credentials_path = api_info["credentials_path"]
    
    creds = Credentials.from_authorized_user_file(token_file_path, scopes)
    
    try:
        service = build("sheets", "v4", credentials=creds)
        body = {"values": values}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute()
        )
        #print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error