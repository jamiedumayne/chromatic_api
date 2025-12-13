import os.path
import os
import pandas as pd
import base64
from email.message import EmailMessage

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#google sheets api overview
#https://developers.google.com/sheets/api/guides/concepts

os.chdir('C:/Users/jamie.dumayne/Documents/python/repos/test/logistic_checks_v0.1/google_api/gmail')


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']

def GMAIL_CONNECT(SCOPES):
  """Shows basic usage of the Gmail API.
  Gets users gmail labels
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
    service = build("gmail", "v1", credentials=creds)
    results = service.users().labels().list(userId="me").execute()
    labels = results.get("labels", [])

    if not labels:
      print("No labels found")
      return
    print("Labels:")
    for label in labels:
      print(label["name"])

  except HttpError as err:
    print(err)
    #trafigura checks label has id Label_3608611773423964260, also want labelIds 'INBOX'

#guide: https://www.geeksforgeeks.org/how-to-read-emails-from-gmail-using-gmail-api-in-python/
def READ_EMAIL(SCOPES):
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)

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

def SEND_EMAIL(reciepent_id, subject, message, sender_id="jamie.dumayne@ahkgroup.com"):
  os.chdir('C:/Users/jamie.dumayne/Documents/python/repos/test/logistic_checks_v0.1/google_api/gmail')
  
  SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)
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

def gmail_create_draft():
  os.chdir('C:/Users/jamie.dumayne/Documents/python/repos/test/logistic_checks_v0.1/google_api/gmail')
  SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  try:
    # create gmail api client
    service = build("gmail", "v1", credentials=creds)

    message = EmailMessage()

    message.set_content("This is automated draft mail")

    message["To"] = "jamie.dumayne@ahkgroup.com"
    message["From"] = "jamie.dumayne@ahkgroup.com"
    message["Subject"] = "Automated draft"

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



#GMAIL_CONNECT(SCOPES)
#READ_EMAIL(SCOPES)
SEND_EMAIL("jamie.dumayne@ahkgroup.com", "Email API Test", "Test message")
#gmail_create_draft()