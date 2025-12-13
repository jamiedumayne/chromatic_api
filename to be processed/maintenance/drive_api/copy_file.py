from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pathlib
import re

def COPY_AND_RENAME_FILE(file_id, destination_folder_url, api_info, new_name_prefix):
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

    new_file_name = new_name_prefix + "_" + old_file_name

    # Prepare metadata for the copy
    file_metadata = {
        "name": new_file_name,
        "parents": [destination_folder_id]
    }

    # Create the copy
    copied_file = drive_service.files().copy(
        fileId=file_id,
        body=file_metadata,
        fields="id, name, parents"
    ).execute()

    return copied_file["id"]