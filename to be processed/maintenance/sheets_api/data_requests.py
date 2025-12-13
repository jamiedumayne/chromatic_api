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

def DOWNLOAD_DATA(sheet_info, tab_name):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = sheet_info["api_info"]["token"]
    credentials_path = sheet_info["api_info"]["credentials_path"]
    
    token = pathlib.Path(token_file_path)

    if os.path.exists(token_file_path):
        creds = Credentials.from_authorized_user_file(token, scopes)

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    #search for spreadsheet and requested range
    full_range = tab_name

    try:
        result = sheet.values().get(
            spreadsheetId=sheet_info["spreadsheet_id"],
            range=full_range
        ).execute()
        values = result.get("values", [])
    except Exception as e:
        print("------------------")
        print("Error: Can't download data from sheet")
        print("Check tab name is correct")
        print(e)
        print("------------------")
        sys.exit()

    if not values:
        print('No data found.')
        return pd.DataFrame()

    # Use first row as headers
    value_df = pd.DataFrame(values)
    value_df.columns = value_df.iloc[0]
    value_df = value_df[1:].reset_index(drop=True)

    return value_df


def DOWNLOAD_FORMULAS(sheet_info, tab_name):
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


def UPLOAD_TO_SHEET(spreadsheet_id, range_request, cell_value, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    token_file_path = api_info["token"]
    credentials_path = api_info["credentials_path"]
    
    creds = Credentials.from_authorized_user_file(token_file_path, scopes)
    
    try:
        service = build("sheets", "v4", credentials=creds)
        body = {"values": cell_value}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_request,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def CLEAR_TAB(sheet_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = sheet_info["api_info"]["token"]
    credentials_path = sheet_info["api_info"]["credentials_path"]
    token = pathlib.Path(token_file_path)

    if os.path.exists(token_file_path):
        creds = Credentials.from_authorized_user_file(token, scopes)

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    try:
        response = sheet.values().clear(
            spreadsheetId=sheet_info["spreadsheet_id"],
            range=sheet_info["tab_name"]  # This clears the whole tab
        ).execute()
        return response
    except Exception as e:
        print("------------------")
        print("Error: Can't clear data from sheet")
        print("Check tab name is correct")
        print(e)
        print("------------------")
        sys.exit()


def GET_TAB_NAMES(spreadsheet_id, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token"]
    credentials_path = api_info["credentials_path"]
    token = pathlib.Path(token_file_path)

    if os.path.exists(token_file_path):
        creds = Credentials.from_authorized_user_file(token, scopes)
    else:
        raise Exception("Missing token.json file. Run the OAuth flow first.")

    service = build("sheets", "v4", credentials=creds)

    # Get spreadsheet metadata
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    # Extract sheet/tab names
    sheets = spreadsheet.get("sheets", [])
    tab_names = [s["properties"]["title"] for s in sheets]

    return tab_names


def DOWNLOAD_SPECIFC_DATA(spreadsheet_id, range_request, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token"]
    credentials_path = api_info["credentials_path"]
    
    token = pathlib.Path(token_file_path)

    creds = Credentials.from_authorized_user_file(token_file_path, scopes)
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
        print("Check tab has been created")
        print("------------------")
        sys.exit()

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=spreadsheet_id,
                                  range=range_request).execute()
        data = rows.get('values')
        df = pd.DataFrame(data[0:])

        return df