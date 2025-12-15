import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from openpyxl.utils.cell import get_column_letter


def SET_COLUMN_DATA_TYPE(column_type, column_type_formatting, column_number, spreadsheet_id, tab_name, api_info, start_row_index = 2):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    first_row = start_row_index + 1
    col_letter = get_column_letter(column_number+1)
    date_validation_formula = f"=OR(ISBLANK(${col_letter}{first_row}),${col_letter}{first_row}<=TODAY())"
    unique_data_formula = f"=COUNTIF(${col_letter}${first_row}:${col_letter},${col_letter}{first_row})>1"
    sheet_row_count = 1000

    if column_type == "Dropdown":
        requests = [
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": column_number,
                        "endColumnIndex": column_number + 1,
                        "startRowIndex": start_row_index,  # skip header row
                        "endRowIndex": sheet_row_count
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_RANGE",
                            "values": [
                                {"userEnteredValue": column_type_formatting}
                            ]
                        },
                        "showCustomUi": True,
                        "strict": True
                    }
                }
            }
        ]
    elif column_type == "DATE":
        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": column_number,
                        "endColumnIndex": column_number + 1,
                        "startRowIndex": start_row_index,  # skip header row
                        "endRowIndex": sheet_row_count
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern": column_type_formatting   # e.g. "dd/mm/yyyy"
                            }
                        },
                        "dataValidation": {
                            "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [
                                {"userEnteredValue": date_validation_formula}
                            ]
                            },
                            "showCustomUi": True,
                            "strict": False
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat,dataValidation"
                }
            }
        ]
    elif column_type == "Unique":
        requests = [
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": sheet_id,
                                "startRowIndex": start_row_index,
                                "endRowIndex": sheet_row_count,
                                "startColumnIndex": column_number,
                                "endColumnIndex": column_number + 1
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [
                                    {"userEnteredValue": unique_data_formula}
                                ]
                            },
                            "format": {
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 0.8,
                                    "blue": 0.8
                                }
                            }
                        }
                    },
                    "index": 0
                }
            }
        ]
    else:
        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": column_number,
                        "endColumnIndex": column_number + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": column_type,
                                "pattern": column_type_formatting
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            }
        ]

    response = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


def SET_MULTIPLE_CELLS_FORMATTING(spreadsheet_id, tab_name, start_col_index, end_col_index, start_row_index, end_row_index, colours, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    rows = []
    for _ in range(start_row_index, end_row_index):
        row = {
            "values": [
                {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": colours["red"],
                            "green": colours["green"],
                            "blue": colours["blue"]
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": colours["font"],
                                "green": colours["font"],
                                "blue": colours["font"]
                            },
                            "fontFamily": "Montserrat"
                        }
                    }
                }
                for _ in range(start_col_index, end_col_index)
            ]
        }
        rows.append(row)

    body = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row_index,
                        "endRowIndex": end_row_index,
                        "startColumnIndex": start_col_index,
                        "endColumnIndex": end_col_index
                    },
                    "rows": rows,
                    "fields": "userEnteredFormat(backgroundColor,textFormat.foregroundColor,textFormat.fontFamily)"
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()





