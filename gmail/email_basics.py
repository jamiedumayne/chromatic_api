import os.path
import os
import base64
from email.message import EmailMessage
import pathlib

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def send_email(reciepent_id, subject, message, gmail_path, sender_id, api_info):
  os.chdir(gmail_path)
  
  SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  token_file_path = api_info["token_path"]
  credentials_path = api_info["credentials_path"]
    
  token_path = pathlib.Path(token_file_path)

  if os.path.exists(token_path):
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)

    # Save the credentials for future use
    with open(token_path, "w") as token:
        token.write(creds.to_json())

  service = build('gmail', 'v1', credentials=creds)
  msg = str(message)

  try:
    # Sets up message parameters.
    message = EmailMessage()
    message.set_content(msg)

    message['To'] = reciepent_id
    message['From'] = sender_id
    message['Subject'] = subject

    # encodes message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()) \
      .decode()

    #creating message
    create_message = {'raw': encoded_message}
    # sending message
    send_message = (service.users().messages().send
                    (userId="me", body=create_message).execute())
    print("Email sent")

  except HttpError as error:
    print("problem sending email")
    send_message = None


def send_email_with_attachment(email_to, subject, message, attachement_name, attachment_path, gmail_path, sender_id, api_info):
  os.chdir(gmail_path)
  
  scopes = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  token_file_path = api_info["token_path"]
  credentials_path = api_info["credentials_path"]
    
  token_path = pathlib.Path(token_file_path)
  creds = Credentials.from_authorized_user_file(token_path, scopes)

  service = build('gmail', 'v1', credentials=creds)
  msg = str(message)

  try:
    # Sets up message parameters.
    message = EmailMessage()
    message.set_content(msg)

    os.chdir(attachment_path)

    with open(attachement_name, 'rb') as content_file:
        content = content_file.read()
        message.add_attachment(content, maintype='application', subtype= (attachement_name.split('.')[1]), filename=attachement_name)
    
    os.chdir(gmail_path)

    message['To'] = email_to
    message['From'] = sender_id
    message['Subject'] = subject

    # encodes message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()) \
      .decode()

    #creating message
    create_message = {'raw': encoded_message}
    # sending message
    send_message = (service.users().messages().send
                    (userId="me", body=create_message).execute())
    print("Email sent")

  except HttpError as error:
    print("problem sending email")
    send_message = None


def gmail_create_draft(reciepent_id, subject, body, sender, api_info):
  scopes = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  token_file_path = api_info["token_path"]
  credentials_path = api_info["credentials_path"]
    
  token_path = pathlib.Path(token_file_path)
  creds = Credentials.from_authorized_user_file(token_path, scopes)

  try:
    # create gmail api client
    service = build("gmail", "v1", credentials=creds)

    msg = str(body)
    message = EmailMessage()
    message.set_content(msg)

    message["To"] = reciepent_id
    message["From"] = sender
    message["Subject"] = subject

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {"message": {"raw": encoded_message}}
    # pylint: disable=E1101
    draft = (
        service.users()
        .drafts()
        .create(userId="me", body=create_message)
        .execute()
    )

    print("Draft email created")

  except HttpError as error:
    print(f"An error occurred: {error}")
    draft = None

  return draft


def download_last_email(api_info):
  scopes = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  token_file_path = api_info["token_path"]
  credentials_path = api_info["credentials_path"]
    
  token_path = pathlib.Path(token_file_path)
  creds = Credentials.from_authorized_user_file(token_path, scopes)

  try:
    service = build("gmail", "v1", credentials=creds)
    result = service.users().messages().list(maxResults=2, userId='me').execute() 
  
    # We can also pass maxResults to get any number of emails. Like this: 
    # result = service.users().messages().list(maxResults=200, userId='me').execute() 
    messages = result.get('messages')

    for msg in messages:
        txt = service.users().messages().get(userId='me', id=msg['id']).execute()
        
        #payload is the bulk of the message
        payload = txt["payload"]

        headers = payload["headers"]

        for d in headers:
            if d['name'] == 'Subject':
                subject = d['value']
                print("Subject: " + subject)
            if d['name'] == 'From': 
                sender = d['value']
                print("From: " + sender)


  except HttpError as err:
    print(err)