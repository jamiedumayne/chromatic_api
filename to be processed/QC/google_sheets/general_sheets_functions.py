

def GET_SHEET_ID(service, spreadsheet_id, tab_name):
    # Get spreadsheet metadata
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get("sheets", [])

    # Look for the sheet with the specified name
    for sheet in sheets:
        if sheet["properties"]["title"] == tab_name:
            return sheet["properties"]["sheetId"]
    
    sheet_id = sheet["properties"]["sheetId"]

    return sheet_id