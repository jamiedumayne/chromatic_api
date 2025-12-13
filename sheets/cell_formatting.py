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