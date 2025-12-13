import os.path
import os
import pandas as pd
import sys
import pathlib
from openpyxl.utils.cell import get_column_letter

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def ADD_COLUMNS(spreadsheet_id, tab_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token"]
    token = pathlib.Path(token_file_path)

    creds = Credentials.from_authorized_user_file(token, scopes)
    sheets_service = build("sheets", "v4", credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    if sheet_id is None:
        raise ValueError(f"Tab '{tab_name}' not found in spreadsheet.")

    # Fetch data in the tab
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=tab_name
    ).execute()
    values = result.get('values', [])

    # Determine the last filled column index
    max_col_index = 0
    for row in values:
        if row:  # ignore completely empty rows
            max_col_index = max(max_col_index, len(row))

    body = {
        "requests": [{
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": max_col_index,
                    "endIndex": max_col_index + 1
                }
            }
        }]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    return max_col_index


def SET_COLUMN_HEADING(spreadsheet_id, tab_name, column_number, column_text, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token"]
    token = pathlib.Path(token_file_path)

    creds = Credentials.from_authorized_user_file(token, scopes)
    sheets_service = build("sheets", "v4", credentials=creds)

    #convert number to letter
    column_letter = get_column_letter(column_number + 1)
    range_request = tab_name + "!" + column_letter + "1"

    try:
        body = {"values": [[column_text]]}
        result = (
            sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_request,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        #print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def DUPLICATE_SUMMARY(spreadsheet_id, old_tab_name, new_tab_name, api_info):
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


def SET_BACKGROUND_COLOUR(spreadsheet_id, tab_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_authorized_user_file(api_info["token"], scopes)
    sheets_service = build('sheets', 'v4', credentials=creds)

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    cell_format = {
        "userEnteredFormat": {
            "backgroundColor": {
                "red": 0.878,
                "green": 0.4,
                "blue": 0.4
            },
            "wrapStrategy": "WRAP",
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }
    }

    # Create 3 rows with 5 cells each (columns B-F)
    formatted_rows = []
    for _ in range(2):  # Rows 12, 13, 14
        formatted_rows.append({
            "values": [cell_format.copy() for _ in range(5)]  # 5 columns: B to F
        })

    body = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 11,  # Row 12
                        "endRowIndex": 14,    # Row 15 (exclusive)
                        "startColumnIndex": 1,  # Column B
                        "endColumnIndex": 6     # Column G (exclusive of F+1)
                    },
                    "rows": formatted_rows,
                    "fields": "userEnteredFormat(backgroundColor,wrapStrategy,horizontalAlignment,verticalAlignment)"
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def SET_ALTERNATING_COLOURS(spreadsheet_id, tab_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_authorized_user_file(api_info["token"], scopes)
    sheets_service = build('sheets', 'v4', credentials=creds)

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    requests = [
        {
        "addBanding": {
            "bandedRange": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 13,
                    "startColumnIndex": 1,
                    "endColumnIndex": 6
                },
                "rowProperties": {
                    "firstBandColor": {
                        "red": 1.0,
                        "green": 1.0,
                        "blue": 1.0
                    },
                    "secondBandColor": {
                        "red": 0.988,
                        "green": 0.925,
                        "blue": 0.925
                    }
                }
            }
        }
    }
    ]

    response = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


def REMOVE_BANDING_FROM_RANGE(spreadsheet_id, tab_name, api_info, start_row, end_row, start_col, end_col):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_authorized_user_file(api_info["token"], scopes)
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            sheet_data = sheet  # Save it!
            break

    if sheet_id is None:
        raise ValueError(f"Sheet '{tab_name}' not found.")

    # Get all banded ranges in the sheet
    banded_ranges = sheet_data.get('bandedRanges', [])

    requests = []

    for band in banded_ranges:
        range_data = band['range']
        band_id = band['bandedRangeId']

        # Check for overlap with target range
        if range_data.get('sheetId') != sheet_id:
            continue

        r1_start = range_data.get('startRowIndex', 0)
        r1_end = range_data.get('endRowIndex', float('inf'))
        c1_start = range_data.get('startColumnIndex', 0)
        c1_end = range_data.get('endColumnIndex', float('inf'))

        # If the ranges intersect
        rows_overlap = not (r1_end <= start_row or r1_start >= end_row)
        cols_overlap = not (c1_end <= start_col or c1_start >= end_col)

        if rows_overlap and cols_overlap:
            requests.append({
                "deleteBanding": {
                    "bandedRangeId": band_id
                }
            })

    # Send batch request to remove the banding
    if requests:
        body = {"requests": requests}
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
    else:
        print("No banded ranges overlapped with the specified range.")


def REMOVE_DATA(spreadsheet_id, tab_name, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_authorized_user_file(api_info["token"], scopes)
    sheets_service = build('sheets', 'v4', credentials=creds)

    range_request = f"{tab_name}!B13:F19"

    # Get the current values
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_request
    ).execute()
    values = result.get('values', [])

    # Clean only text values
    updated_values = []
    for row in values:
        new_row = []
        for cell in row:
            new_row.append('')
        updated_values.append(new_row)

    # Fill in missing empty rows or columns to match range size
    num_rows = 7
    num_cols = 5
    for i in range(len(updated_values), num_rows):
        updated_values.append([''] * num_cols)
    for row in updated_values:
        while len(row) < num_cols:
            row.append('')

    # Write cleaned values back
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_request,
        valueInputOption="USER_ENTERED",
        body={"values": updated_values}
    ).execute()


def UPLOAD_CELL_VALUE(spreadsheet_id, range_request, cell_text_list, api_info):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    token_file_path = api_info["token"]
    token = pathlib.Path(token_file_path)

    creds = Credentials.from_authorized_user_file(token, scopes)
    sheets_service = build("sheets", "v4", credentials=creds)


    try:
        body = {"values": [cell_text_list]}
        result = (
            sheets_service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_request,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        #print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error