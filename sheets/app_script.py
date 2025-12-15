import os
import sys

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


def attach_app_script(spreadsheet_id, filename, script_id, api_info):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info.get("scopes"))
    script_service = build("script", "v1", credentials=creds)

    src = script_service.projects().getContent(scriptId=script_id).execute()
    files = src.get("files", []) if src else []


    create_body = {"title": filename, "parentId": spreadsheet_id}
    proj = script_service.projects().create(body=create_body).execute()
    bound_script_id = proj.get("scriptId")

    update_body = {"files": files}
    resp = script_service.projects().updateContent(body=update_body, scriptId=bound_script_id).execute()