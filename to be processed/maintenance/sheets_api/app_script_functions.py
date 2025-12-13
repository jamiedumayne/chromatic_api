import os
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def ATTACH_SCRIPT(spreadsheet_id, filename, api_info, merge_script=False):
    creds = Credentials.from_authorized_user_file(api_info["token"], api_info.get("scopes"))
    script_service = build("script", "v1", credentials=creds)

    export_script_id = "1blzOHOZ1zKGdXSyI-iAL8R3Nw3IAYimIrBzfrC8TBjF-Qe9DycLrrzfv" #example sheet with export script in
    #SOURCE_SCRIPT_ID = "1hU68IauO9oY5z31GC_Q60JmQEldyhjPTPzt7XotzboB7wtp-3NGd7Qsa" #actual centralised export script

    #merge_secript_id = "" #centralised merge script
    merge_secript_id = "" #example sheet with centralised merge script

    
    src_export = script_service.projects().getContent(scriptId=export_script_id).execute()
    files_export = src_export.get("files", [])

    if merge_script:
        src_merge = script_service.projects().getContent(scriptId=merge_secript_id).execute()
        files_merge = src_merge.get("files", []) if src_merge else []
    
        files_all = files_export + files_merge
    else:
        files_all = files_export

    create_body = {"title": filename, "parentId": spreadsheet_id}
    proj = script_service.projects().create(body=create_body).execute()
    bound_script_id = proj.get("scriptId")

    # 3) copy files into bound project (updateContent replaces all files)
    update_body = {"files": files_all}
    resp = script_service.projects().updateContent(body=update_body, scriptId=bound_script_id).execute()

