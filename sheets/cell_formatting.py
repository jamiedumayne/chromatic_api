from sheets.sheets_general_functions import get_sheet_id

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def add_tickbox(row_number, spreadsheet_id, tab_name, api_info):
    row_number = int(row_number)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    token_file_path = api_info["token_path"]
    creds = Credentials.from_authorized_user_file(token_file_path, scopes)

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet_id = get_sheet_id(service, spreadsheet_id, tab_name)
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


def cell_font_colour_background_size(spreadsheet_id, tab_name, col_index, row_index, colours, api_info, bold=False, num_rows = 1):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)
    end_col_index = col_index + 1
    end_row_index = row_index + num_rows

    #get sheet_id
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    body = {
    "requests": [
        {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_index,
                "endRowIndex": end_row_index,
                "startColumnIndex": col_index,
                "endColumnIndex": end_col_index
                },
            "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": colours["background_colour"]["red"],
                                "green": colours["background_colour"]["green"],
                                "blue": colours["background_colour"]["blue"]
                            },
                            "textFormat": {
                                "foregroundColor": {
                                    "red": colours["font_colour"]["red"],
                                    "green": colours["font_colour"]["green"],
                                    "blue": colours["font_colour"]["blue"]
                                },
                                "fontFamily": "Montserrat",
                                "fontSize": colours["font_size"],
                                "bold": bold
                            },
                            "wrapStrategy": "WRAP",
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    # only update formatting fields (no values)
                    "fields": "userEnteredFormat(backgroundColor,textFormat.foregroundColor,textFormat.fontFamily,textFormat.fontSize,textFormat.bold,wrapStrategy,horizontalAlignment,verticalAlignment)"
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def set_column_alternating_colours(column_number, spreadsheet_id, tab_name, row_start_index, column_colours, api_info,
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


def set_multiple_column_banding(spreadsheet_id, tab_name, col_start, col_end, row_start, column_colours, api_info):
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


def remove_banding(spreadsheet_id, tab_name, api_info, start_row, end_row, start_col, end_col):
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


def set_border_style(spreadsheet_id, tab_name, border_colour, api_info, width=1, style="SOLID"):
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

    colour = {
        "red": border_colour["red"],
        "green": border_colour["red"],
        "blue": border_colour["red"]
    }

    border_style = {
        "style": style,
        "width": width,
        "color": colour
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


def create_dropdown(dropdown_values, column_number, row_number, spreadsheet_id, tab_name, api_info):
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


def set_font(spreadsheet_id, tab_name, start_col_index, end_col_index, start_row_index, colours, api_info, font_family = "Montserrat"):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    text_format = {
        "foregroundColor": {
            "red": colours["font"]["red"],
            "green": colours["font"]["green"],
            "blue": colours["font"]["blue"],
            "alpha": 1.0
        },
        "fontFamily": font_family
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


def wrap_row_text(spreadsheet_id, tab_name, row_index, api_info):
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


def set_column_width(spreadsheet_id, tab_name, column_number, width, api_info, width_type="COLUMNS"):
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


def merge_cells_batch(spreadsheet_id, tab_name, merge_ranges, api_info):
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


def merge_multiple_cells(spreadsheet_id, tab_name, start_col_index, end_col_index, start_row_index, api_info, end_row_index_num=1):
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