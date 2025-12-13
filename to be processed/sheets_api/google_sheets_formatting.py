import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from openpyxl.utils.cell import get_column_letter


def SET_FONT_COLOUR_COLUMN(spreadsheet_id, tab_name, start_col_index, end_col_index, start_row_index, colours, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    color_value = colours["font"]
    text_format = {
        "foregroundColor": {
            "red": color_value,
            "green": color_value,
            "blue": color_value,
            "alpha": 1.0
        },
        "fontFamily": "Montserrat"
    }

    body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row_index,
                        "startColumnIndex": start_col_index,
                        "endColumnIndex": end_col_index
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": text_format
                        }
                    },
                    "fields": "userEnteredFormat.textFormat"
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


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



"""
    elif column_type == "DATE":
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
                            "type": "CUSTOM_FORMULA",
                            "values": [
                                {"userEnteredValue": date_validation_formula}
                            ]
                        },
                        "showCustomUi": True,
                        "strict": False
                    }
                }
            }
        ]
"""



def SET_DROPDOWN_CELL(dropdown_values, column_number, row_number, spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    requests = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": column_number,
                    "endColumnIndex": column_number + 1,
                    "startRowIndex": row_number,
                    "endRowIndex": row_number+1
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in dropdown_values]
                    },
                    "showCustomUi": True,
                    "strict": True
                }
            }
        }
    ]


    response = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


def SET_COLUMN_ALTERNATING_COLOURS(column_number, spreadsheet_id, tab_name, row_start_index, column_colours, api_info,
                                    row_end_index=None, swap_colours=False):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    banding_range = {
        "sheetId": sheet_id,
        "startRowIndex": row_start_index,
        "startColumnIndex": column_number,
        "endColumnIndex": column_number + 1
    }

    # Add endRowIndex if user provided it
    if row_end_index is not None:
        banding_range["endRowIndex"] = row_end_index
    
    # Default colours
    first_band = {"red": 1.0, "green": 1.0, "blue": 1.0}
    second_band = {
        "red": column_colours["red"],
        "green": column_colours["green"],
        "blue": column_colours["blue"]
    }

    # Swap if requested
    if swap_colours:
        first_band, second_band = second_band, first_band
    
    requests = [
        {
            "addBanding": {
                "bandedRange": {
                    "range": banding_range,
                    "rowProperties": {
                        "firstBandColor": first_band,
                        "secondBandColor": second_band
                    }
                }
            }
        }
    ]

    response = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


def SET_MULTIPLE_COLUMN_ALTERNATING_COLOURS(col_start, col_end, row_start, spreadsheet_id, tab_name, column_colours, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

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
                    "startRowIndex": row_start,
                    "startColumnIndex": col_start,
                    "endColumnIndex": col_end
                },
                "rowProperties": {
                    "firstBandColor": {
                        "red": 1.0,
                        "green": 1.0,
                        "blue": 1.0
                    },
                    "secondBandColor": {
                        "red": column_colours["red"],
                        "green": column_colours["green"],
                        "blue": column_colours["blue"]
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


def SET_FORMATTING(spreadsheet_id, tab_name, col_index, row_index, colours, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)
    end_col_index = col_index + 1
    end_row_index = row_index + 1

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    body = {
    "requests": [
        {
        "updateCells": {
            "range": {
            "sheetId": sheet_id,
            "startRowIndex": row_index,
            "endRowIndex": end_row_index,
            "startColumnIndex": col_index,
            "endColumnIndex": end_col_index
            },
            "rows": [
            {
                "values": [
                {
                    "userEnteredFormat": {
                    "backgroundColor": {
                        "red": colours["red"],
                        "green" : colours["green"],
                        "blue" : colours["blue"]
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
                ]
            }
            ],
            "fields": "userEnteredFormat(backgroundColor,textFormat.foregroundColor,textFormat.fontFamily)"
        }
        }
    ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


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


def SET_ALL_BORDERS_WHITE(spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID and dimensions
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    row_count = sheet['properties']['gridProperties']['rowCount']
    column_count = sheet['properties']['gridProperties']['columnCount']

    white = {
        "red": 1.0,
        "green": 1.0,
        "blue": 1.0
    }

    border_style = {
        "style": "SOLID",
        "width": 1,
        "color": white
    }

    request = {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": row_count,
                "startColumnIndex": 0,
                "endColumnIndex": column_count
            },
            "top": border_style,
            "bottom": border_style,
            "left": border_style,
            "right": border_style,
            "innerHorizontal": border_style,
            "innerVertical": border_style
        }
    }

    body = {"requests": [request]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def WRAP_ROW(spreadsheet_id, tab_name, row_index, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Get sheet_id and column count
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            column_count = sheet['properties']['gridProperties']['columnCount']
            break

    body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_index,
                        "endRowIndex": row_index + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": column_count
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP",
                            "textFormat": {
                                "bold": True
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(wrapStrategy,textFormat.bold,horizontalAlignment,verticalAlignment)"
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def SET_WIDTH(spreadsheet_id, tab_name, column_number, width, api_info, width_type="COLUMNS"):
    # Authenticate
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Find the sheet ID from the tab name
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    if sheet_id is None:
        raise ValueError(f"Tab '{tab_name}' not found in spreadsheet.")

    # Request to update column width
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": width_type,
                    "startIndex": column_number,
                    "endIndex": column_number + 1
                },
                "properties": {
                    "pixelSize": width
                },
                "fields": "pixelSize"
            }
        }
    ]

    # Execute batch update
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


def MERGE_CELLS_BATCH(spreadsheet_id, tab_name, merge_ranges, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    if sheet_id is None:
        raise ValueError(f"Sheet {tab_name} not found")

    # Prepare batch requests
    requests = []
    for start_row, end_row, start_col, end_col in merge_ranges:
        requests.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col
                },
                "mergeType": "MERGE_ALL"
            }
        })

    # Execute batch update
    body = {"requests": requests}
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()




def MERGE_MULTIPLE_CELLS(spreadsheet_id, tab_name, start_col_index, end_col_index, start_row_index, api_info, end_row_index_num=1):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)
    end_row_index = start_row_index + end_row_index_num

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    body = {
    "requests": [
        {
           "mergeCells": {
              "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row_index,
                "endRowIndex": end_row_index,
                "startColumnIndex": start_col_index,
                "endColumnIndex": end_col_index

           },
           "mergeType": "MERGE_ALL"
        }
        }
    ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


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