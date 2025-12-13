import os.path
import os
import pandas as pd
import base64
from email.message import EmailMessage
import pathlib

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def SEND_EMAIL(reciepent_id, subject, message, gmail_path, sender_id="jamie.dumayne@ahkgroup.com"):
  os.chdir(gmail_path)
  
  SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']

  token_path = pathlib.Path("token.json")
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)

    # Save the credentials for future use
    with open("token.json", "w") as token:
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

  return send_message


def SEND_EMAIL_ATTACHMENT(email_to, subject, message, attachement_name, attachment_path, gmail_path, api_info, sender_id):
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

  return send_message