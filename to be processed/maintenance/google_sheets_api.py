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


def CREATE_SHEET(form_answers, filename, column_information, api_info):   
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    spreadsheet = {"properties": {"title": filename}}
    spreadsheet = (
      service.spreadsheets()
      .create(body=spreadsheet, fields="spreadsheetId")
      .execute()
    )
   
    spreadsheet_id = spreadsheet.get("spreadsheetId")
    #folder_id = "1vHY71llhNBYXhR1uxkBKdka9bedGbbkb"
    folder_id = "1A2HpJxXF2xf_e_LcKnHUDC_7t4vWx6Rr"
   
    MOVE_FILE(spreadsheet_id, folder_id, creds)
   
    modify_service = build('sheets', 'v4', credentials=creds)

    #set report tab names
    if form_answers["client"]  == "Trafigura":
        if form_answers["file type"] == 'Transhipment' or form_answers["file type"] == 'Loadport':
           loading_tab = "Traf_Loading_Report"
           reception_tab = "Traf_Reception_Report"
        else:
            if form_answers["services"]["si_required"] == "Yes":
                if form_answers["file type"] == 'Source':
                    reception_tab = "blank"
                    loading_tab = "Traf_Loading_Report"
                else:
                    reception_tab = "Traf_Reception_Report"
                    loading_tab = "blank"
            else:
                loading_tab = "blank"
                reception_tab = "blank"
        
        report_tab = "Traf_Report"
        summary_tab = "Traf_Summary"
        si_tab = "Traf_Report_SI_HC"
    else:
        if form_answers["file type"] == 'Transhipment' or form_answers["file type"] == 'Loadport':
           loading_tab = "Client_Loading_Report"
           reception_tab = "Client_Reception_Report"
        else:
            if form_answers["services"]["si_required"] == "Yes":
                if form_answers["file type"] == 'Source':
                    reception_tab = "blank"
                    loading_tab = "Client_Loading_Report"
                else:
                    reception_tab = "Client_Reception_Report"
                    loading_tab = "blank"
            else:
                loading_tab = "blank"
                reception_tab = "blank"
        
        report_tab = "Client_Report"
        summary_tab = "Client_Summary"
        si_tab = "Client_Report_SI_HC"
    
    reporting_tab_names = {"summary": summary_tab, "report":report_tab, "loading":loading_tab, "reception":reception_tab, "si":si_tab}

    #add tabs
    ADD_TAB(spreadsheet_id, "Export", api_info)
    ADD_TAB(spreadsheet_id, reporting_tab_names["summary"], api_info)
    
    if form_answers["file type"] == 'Transhipment' or form_answers["file type"] == 'Loadport':
        ADD_TAB(spreadsheet_id, reporting_tab_names["report"], api_info)
        ADD_TAB(spreadsheet_id, reporting_tab_names["reception"], api_info)
        ADD_TAB(spreadsheet_id, reporting_tab_names["loading"], api_info)
        if form_answers["services"]["si_required"] == "Yes":
            ADD_TAB(spreadsheet_id, reporting_tab_names["si"], api_info)
    else:
        if form_answers["services"]["si_required"] == "Yes":
            if form_answers["file type"] == 'Source':
                ADD_TAB(spreadsheet_id, reporting_tab_names["report"], api_info)
                ADD_TAB(spreadsheet_id, reporting_tab_names["loading"], api_info)
                ADD_TAB(spreadsheet_id, reporting_tab_names["si"], api_info)
            else:
                ADD_TAB(spreadsheet_id, reporting_tab_names["report"], api_info)
                ADD_TAB(spreadsheet_id, reporting_tab_names["reception"], api_info)
                ADD_TAB(spreadsheet_id, reporting_tab_names["si"], api_info)
        else:
            ADD_TAB(spreadsheet_id, reporting_tab_names["report"], api_info)

    ADD_SUMMARY_TABS(form_answers, spreadsheet_id, modify_service)

    if form_answers["file type"] == 'Transhipment' or form_answers["file type"] == 'Loadport' or reporting_tab_names["si"] != "blank":
        ADD_TAB(spreadsheet_id, "Observations", api_info)
        ADD_TAB(spreadsheet_id, "Merge Log", api_info)

    ADD_TAB(spreadsheet_id, "Validation", api_info)
   
    #get sheet_id
    sheet_metadata = modify_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    
    #add columns to report, reception and loading tabs
    if form_answers["file type"] == 'Transhipment' or form_answers["file type"] == 'Loadport':
        FIND_TAB_AND_ADD_COLUMNS(column_information["num_of_loading_cols"], reporting_tab_names["loading"], sheet_metadata, spreadsheet_id, modify_service)
        FIND_TAB_AND_ADD_COLUMNS(column_information["num_of_reception_cols"], reporting_tab_names["reception"], sheet_metadata, spreadsheet_id, modify_service)

    FIND_TAB_AND_ADD_COLUMNS(column_information["num_of_report_cols"], reporting_tab_names["report"], sheet_metadata, spreadsheet_id, modify_service)

    return spreadsheet_id, reporting_tab_names


def ADD_SUMMARY_TABS(form_answers, spreadsheet_id, api_info):
    if form_answers["summaries"]["truck_summary"] == 1:
        ADD_TAB(spreadsheet_id, "Truck_Summary", api_info)
    
    if form_answers["summaries"]["train_summary"] == 1:
        ADD_TAB(spreadsheet_id, "Train_Summary", api_info)
    
    if form_answers["summaries"]["lot_summary"] == 1:
        ADD_TAB(spreadsheet_id, "Lot_Summary", api_info)
    
    if form_answers["summaries"]["container_summary"] == 1:
        ADD_TAB(spreadsheet_id, "Container_Summary", api_info)
    
    if form_answers["summaries"]["hc_summary"] == 1:
        ADD_TAB(spreadsheet_id, "HC_Summary", api_info)


def DOWNLOAD_DATA(spreadsheet_id, range_request, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    try:
        result = (sheet.values().get(spreadsheetId=spreadsheet_id, range=range_request)
            .execute())
        values = result.get("values", [])
    except:
        print("------------------")
        print("Error: Can't download data from sheet")
        print("Check tab name is correct")
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


def UPDATE_VALUES(spreadsheet_id, range_name, values, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    # pylint: disable=maybe-no-member
    try:
        service = build("sheets", "v4", credentials=creds)
        body = {"values": values}
        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="USER_ENTERED",
                body=body,
                )
            .execute()
            )
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def MOVE_FILE(spreadsheet_id, folder_id, creds):
    drive_service = build("drive", "v3", credentials=creds)

    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=spreadsheet_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))

    # Move the file to the new folder
    file = drive_service.files().update(
        fileId=spreadsheet_id,
        addParents=folder_id,
        removeParents=previous_parents,
        fields='id, parents',
        supportsAllDrives=True
    ).execute()

    print("File created")


def ADD_TAB(spreadsheet_id, tab_name, api_info):
    print(api_info)
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


def FIND_TAB_AND_ADD_COLUMNS(number_of_columns, tab_name, sheet_metadata, spreadsheet_id, modify_service):
    if number_of_columns <= 0:
        number_of_columns = 0
            
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == tab_name:
            sheet_id = sheet['properties']['sheetId']
            break

    ADD_COLUMNS(spreadsheet_id, sheet_id, modify_service, number_of_columns)


def ADD_COLUMNS(spreadsheet_id, sheet_id, modify_service, num_of_columns_to_add):
    body = {
            "requests":{
                "insertDimension":{
                    "range":{
                        "sheetId": sheet_id, # <--- Please set the sheet ID of "Class Data" sheet.
                        "dimension": "COLUMNS",
                        "startIndex": 10,
                        "endIndex": 10 + num_of_columns_to_add
                    }
                }
                
            }
        }

    modify_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def DELETE_TAB(spreadsheet_id, tab_name, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    sheets_service = build('sheets', 'v4', credentials=creds)

    #sheet_metadata = sheets_service.spreadsheets().get(spreadsheetID=spreadsheet_id).execute()
    #for sheet in sheet_metadata['sheets']:
    #    if sheet['properties']['title'] == tab_name:
    #        sheet_id = sheet['properties']['sheetId']
    #        break
    
    body = {
        "requests": [
            {
                "deleteSheet": {
                    "sheetId": 0
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


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


def DELETE_REMAINING_ROWS(spreadsheet_id, tab_name, last_row, api_info):
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


def GET_FILENAME(spreadsheet_id, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)

    try:
        # only request the title to keep the response small
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="properties/title"
        ).execute()
        title = spreadsheet.get("properties", {}).get("title")
    except Exception as e:
        print("------------------")
        print("Error: Can't retrieve spreadsheet metadata")
        print("Exception:", e)
        print("------------------")
        sys.exit()

    if not title:
        print("Spreadsheet title not found.")
        return None

    return title




def INSERT_IMAGE(spreadsheet_id, range, image_url, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build('sheets', 'v4', credentials=creds)

    image_formula = f'=IMAGE("{image_url}")'

    # Prepare request body
    body = {
        'values': [[image_formula]]
    }

    # Update cell with formula
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

