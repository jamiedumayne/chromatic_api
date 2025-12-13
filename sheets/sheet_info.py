import sys

def get_filename(spreadsheet_id, api_info):
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


def get_tab_names(spreadsheet_id, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info["scopes"])
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Fetch the spreadsheet details
    result = sheet.get(spreadsheetId=spreadsheet_id).execute()
    
    # Extract sheet names
    sheets = result.get("sheets", [])
    sheet_names = [s["properties"]["title"] for s in sheets]
    filename = result.get("properties", {}).get("title", "Unknown Spreadsheet")

    return sheet_names, filename