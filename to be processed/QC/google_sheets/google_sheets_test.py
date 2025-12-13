import os.path
import os
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#google sheets api overview
#https://developers.google.com/sheets/api/guides/concepts

#os.chdir('C:/Users/jamie.dumayne/Documents/python/scripts/trafigura_checks/google_api/google_sheets/')
os.chdir('C:/Users/jamie.dumayne/Documents/python/repos/test/logistic_checks_v0.1/google_api/google_sheets')


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1Hkkkcbw4eLKZFYoL0PpLcNSYpK1vAIpLtKGmVCwJvLo"
SAMPLE_RANGE_NAME = "Sheet1!A2:B"


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                  range=SAMPLE_RANGE_NAME).execute()
        data = rows.get('values')
        df = pd.DataFrame(data[0:], columns=["A", "B"])
        print(df)
        print("COMPLETE: Data copied")
  except HttpError as err:
    print(err)

def update_values(spreadsheet_id, range_name, value_input_option):

  creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # pylint: disable=maybe-no-member
  try:
    service = build("sheets", "v4", credentials=creds)
    values = [["Test"], ["Test2"]]
    body = {"values": values}
    result = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            body=body,
        )
        .execute()
    )
    print(f"{result.get('updatedCells')} cells updated.")
    return result
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error

#example of updating a spreadsheet
update_values(
      "1Hkkkcbw4eLKZFYoL0PpLcNSYpK1vAIpLtKGmVCwJvLo",
      "Sheet1!C1:C2",
      "USER_ENTERED",
  )


if __name__ == "__main__":
  main()