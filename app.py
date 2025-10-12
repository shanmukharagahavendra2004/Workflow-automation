import os
import base64
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from time import sleep
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
SHEET_NAME = os.getenv("SHEET_NAME", "Emails")
POLL_INTERVAL = 30  # seconds

# OAuth scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/spreadsheets'
]


# Authenticate Gmail API
flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
creds = flow.run_local_server(port=0)
gmail_service = build('gmail', 'v1', credentials=creds)

# Authenticate Google Sheets
gc = gspread.service_account(filename='service_account.json')  # Use your service account
sheet = gc.open(SHEET_NAME).sheet1


# Keep track of sent rows
sent_rows = []

while True:
    rows = sheet.get_all_records()
    for i, row in enumerate(rows):
        if i in sent_rows:
            continue

        to_email = row.get("Email")
        subject = row.get("Subject")
        body = row.get("Body")

        if not to_email or not body:
            continue

        # Create email
        message = MIMEText(body)
        message['to'] = to_email
        message['subject'] = subject or "No Subject"
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

        try:
            gmail_service.users().messages().send(
                userId="me",
                body={'raw': raw}
            ).execute()
            print(f"Sent email to {to_email}")
        except Exception as e:
            print(f"Error sending to {to_email}: {e}")

        sent_rows.append(i)

    sleep(POLL_INTERVAL)
