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

def move_file(spreadsheet_id, folder_id, creds):
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


def copy_file(file_id, destination_folder_url, new_name, api_info):
    scopes = ["https://www.googleapis.com/auth/drive"]
    
    token_file_path = api_info["token"]
    token = pathlib.Path(token_file_path)

    creds = Credentials.from_authorized_user_file(token, scopes)
    drive_service = build("drive", "v3", credentials=creds)

    # Extract folder ID from URL
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", destination_folder_url)
    if not match:
        raise ValueError("Invalid Google Drive folder URL.")
    destination_folder_id = match.group(1)

    # Step 1: Get the original file name
    original_file = drive_service.files().get(
        fileId=file_id,
        fields="name"
    ).execute()
    old_file_name = original_file["name"]

    # Prepare metadata for the copy
    file_metadata = {
        "name": new_name,
        "parents": [destination_folder_id]
    }

    # Create the copy
    copied_file = drive_service.files().copy(
        fileId=file_id,
        body=file_metadata,
        fields="id, name, parents"
    ).execute()

    return copied_file["id"]