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


def create_new_sheet(filename, api_info):   
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    spreadsheet = {"properties": {"title": filename}}
    spreadsheet = (
      service.spreadsheets()
      .create(body=spreadsheet, fields="spreadsheetId")
      .execute()
    )
   
    spreadsheet_id = spreadsheet.get("spreadsheetId")

    return spreadsheet_id


def add_new_tab(spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    body = {
            "requests":{
                "addSheet":{
                    "properties":{
                        "title":tab_name
                    }
                }
                
            }
        }

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def duplicate_tab(spreadsheet_id, old_tab_name, new_tab_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    
    creds = Credentials.from_authorized_user_file(api_info["token"], scopes)
    service = build("sheets", "v4", credentials=creds)

    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == old_tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    if sheet_id is None:
        raise ValueError(f"Tab '{old_tab_name}' not found in spreadsheet.")

    body = {
        "requests": [
            {
                "duplicateSheet": {
                    "sourceSheetId": sheet_id,
                    "insertSheetIndex": 1,  # Optional: where to insert the new sheet
                    "newSheetName": new_tab_name
                }
            }
        ]
    }

    try:
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        return response
    except Exception as e:
        print("Failed to duplicate tab.")
        print(e)
        return None



def add_new_columns(spreadsheet_id, tab_name, modify_service, num_of_columns_to_add, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)

    sheet_metadata = service.spreadsheets().get(spreadsheetID=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    body = {
            "requests":{
                "insertDimension":{
                    "range":{
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 10,
                        "endIndex": 10 + num_of_columns_to_add
                    }
                }
                
            }
        }

    modify_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def delete_columns(spreadsheet_id, tab_name, first_column, last_column, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID and total columns
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    total_columns = last_column-first_column

    # Only delete if there are columns to remove
    if last_column + 1 < total_columns:
        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": first_column,
                    "endIndex": total_columns
                }
            }
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [request]}
        ).execute()



def delete_tab(spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetID=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    body = {
        "requests": [
            {
                "deleteSheet": {
                    "sheetId": sheet_id
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def delete_rows_after(spreadsheet_id, tab_name, last_row, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID and total columns
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    total_rows = sheet['properties']['gridProperties']['rowCount']

    # Only delete if there are rows to remove
    if last_row + 1 < total_rows:
        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": last_row + 1,
                    "endIndex": total_rows
                }
            }
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [request]}
        ).execute()