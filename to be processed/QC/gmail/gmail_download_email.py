import os.path
import os
import pandas as pd
import base64
from base64 import urlsafe_b64decode
import email
from email.message import EmailMessage

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def GMAIL_CONNECT():
  #make connection and generate token
  #return labels

  os.chdir('C:/Users/jamie.dumayne/Documents/python/scripts/trafigura_checks/google_api/gmail/')
  SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
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

#GMAIL_CONNECT(SCOPES)

def GET_TRAF_EMAIL(file_path, num_emails):
    os.chdir('C:/Users/jamie.dumayne/Documents/python/scripts/trafigura_checks/google_api/gmail/')
    SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", 'https://www.googleapis.com/auth/gmail.compose']
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    service = build("gmail", "v1", credentials=creds)
    result = service.users().messages().list(maxResults=num_emails, userId='me').execute() 
  
    # We can also pass maxResults to get any number of emails. Like this: 
    # result = service.users().messages().list(maxResults=200, userId='me').execute() 
    messages = result.get('messages')

    for msg in messages:
        txt = service.users().messages().get(userId='me', id=msg['id']).execute()
        try:
            get_label = txt['labelIds']

        except:
           print("No label") 

        traf_label = "Label_3608611773423964260"

        if traf_label in get_label:
            #payload is the bulk of the message
            payload = txt["payload"]
            headers = payload["headers"]

            for d in headers:
                if d['name'] == 'Subject':
                    subject = d['value']
                if d['name'] == 'From': 
                    sender = d['value']

            
            #get data from email
            for part in payload["parts"]:
                if part['filename']:
                    #get filename
                    email_attachment_filename  = part['filename']

                    if 'xlsx' in email_attachment_filename:
                        if 'data' in part['body']:
                            data = part['body']['data']
                        else:
                            att_id = part['body']['attachmentId']
                            att = service.users().messages().attachments().get(userId='me', messageId=msg['id'],id=att_id).execute()
                            data = att['data']
                        
                        #Take the data and convert it into something readable
                        bytesFile = urlsafe_b64decode(data)

                        #check if file is an excel file, b'PK is what an excel file should start with?
                        if bytesFile[0:2] != b'PK':
                            raise ValueError('The attachment is not a XLSX file!')
                        message = email.message_from_bytes(bytesFile)


                        run_traf_check = CHECK_FILE_IS_ON_COMPUTER(file_path, email_attachment_filename, message)

                        return run_traf_check

def CHECK_FILE_IS_ON_COMPUTER(file_path, email_attachment_filename, message):
    os.chdir(file_path)

    all_filenames_in_archive = pd.DataFrame({'Filenames':os.listdir()})
    filtered_filenames = all_filenames_in_archive[all_filenames_in_archive["Filenames"] == email_attachment_filename]
    print("File in email: " + email_attachment_filename)

    run_traf_check = False
   
    if len(filtered_filenames) == 0 and 'Surveyor_Logistics' in email_attachment_filename: 
        print("File not on computer, downloading")

        email_filename = email_attachment_filename

        #export the file as an excel file
        open(email_attachment_filename, 'wb').write(message.get_payload(decode=True))
        run_traf_check = True
        #elif len(filtered_filenames) == 1 and 'Surveyor_Logistics' in email_attachment_filename:
        #print("File already on computer, not downloaded")

        #renamed_email_attachment = email_attachment_filename.rstrip(".xlsx") + " (1).xlsx"
        #email_filename = renamed_email_attachment
        #export the file as an excel file
        #open(email_filename, 'wb').write(message.get_payload(decode=True))

        #os.rename(email_attachment_filename, renamed_email_attachment)
        #run_traf_check = True
    else:
      print("File already downloaded")
    
    return run_traf_check

#send email
#https://medium.com/@developervick/sending-mail-using-gmail-api-and-detailed-guide-to-google-apis-documentation-41c000706b50
def SEND_EMAIL(reciepent_id, subject, message, sender_id="jamie.dumayne@ahkgroup.com"):
  os.chdir('C:/Users/jamie.dumayne/Documents/python/scripts/trafigura_checks/google_api/gmail/')
  
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

#SEND_EMAIL("jamie.dumayne@ahkgroup.com", "Email API Test", "Test message")

def gmail_create_draft():
  os.chdir('C:/Users/jamie.dumayne/Documents/python/scripts/trafigura_checks/google_api/gmail/')
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

#gmail_create_draft()