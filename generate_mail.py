import os
import base64
import calendar
import datetime as dt
import pickle

from dotenv import load_dotenv

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parsedate_tz, mktime_tz

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

import pandas as pd

load_dotenv()


# The Gmail API scope needed for sending emails
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.readonly"]

class HetCAKEmail():
    def __init__(self, sender: str, to: str):
        """

        :param sender:
        :param to:
        """
        self.service = self.authenticate_gmail()
        self.sender = sender
        self.to = to

    @staticmethod
    def authenticate_gmail():
        """Authenticate and create a service object for sending emails."""
        creds = None
        # The file token.pickle stores the user's access and refresh tokens.
        # It is created automatically when the authorization flow completes for the first time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        # Build the Gmail API service
        service = build('gmail', 'v1', credentials=creds)
        return service

    @staticmethod
    def get_previous_month_start_end():
        today = dt.datetime.today()
        current_year = today.year
        current_month = today.month

        # If we're in January, go back to December of the previous year
        if current_month == 1:
            current_month = 12
            current_year -= 1
        else:
            current_month -= 1

        # Get the first day of the previous month
        first_day_of_previous_month = dt.datetime(current_year, current_month, 1)

        # Get the last day of the previous month using calendar.monthrange
        _, last_day_of_previous_month = calendar.monthrange(current_year, current_month)
        last_day_of_previous_month = dt.datetime(current_year, current_month, last_day_of_previous_month, 23, 59, 59)

        # Convert to timestamps
        start_timestamp = first_day_of_previous_month.date().strftime("%a, %d %b %Y")
        end_timestamp = last_day_of_previous_month.date().strftime("%a, %d %b %Y")

        return start_timestamp, end_timestamp

    def create_message(self, sender, to, subject, body = None):
        """Create an email message."""
        message = MIMEMultipart()
        message['to'] = to
        message['from'] = sender
        message['subject'] = subject

        date_start, date_end = self.get_previous_month_start_end()
        sent_emails = self.get_emails(date_start, date_end)
        sent_emails_df = pd.DataFrame(sent_emails, columns=["Date", "Event"])

        body = body or f"""
        <body>
        Beste, <br><br>
        
        
        Over de afgelopen periode ({date_start} t/m {date_end}) zijn de volgende {len(sent_emails)} mails verzonden: 
        <br><br>
    
        {sent_emails_df.to_html()}
        <br>
        
        Met vriendelijke groet,<br>
        CINQ ICT
        </body>
        """

        msg = MIMEText(body, 'html')
        message.attach(msg)

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        return {'raw': raw_message}

    @staticmethod
    def date_is_within_range(date_start, date_end, date_to_check) -> bool:
        date_format = "%a, %d %b %Y"  # Format: Weekday, Day Month Year
        date_to_check_tz = dt.datetime.fromtimestamp(mktime_tz(parsedate_tz(date_to_check)))
        date_to_check = date_to_check_tz.date()

        # Parse the start and end date (without time and timezone)
        start_date = dt.datetime.strptime(date_start, date_format).date()
        end_date = dt.datetime.strptime(date_end, date_format).date()

        # Check if the date falls within the range
        if start_date <= date_to_check <= end_date:
            print(f"{date_to_check} is within the range.")
            return True
        else:
            print(f"{date_to_check} is outside the range.")
            return False

    def get_emails(self, date_start: dt.datetime.date = None, date_end: dt.datetime.date = None, folders: list = ["INBOX"]) -> list[tuple]:
        try:
            excluded_subject = "Mailrapportage"

            results = self.service.users().messages().list(userId="me", labelIds=folders).execute()
            ids = [ids["id"] for ids in results.get("messages", [])]

            sent_emails = []
            for id in ids:
                message_data = self.service.users().messages().get(userId="me", id=id, format="full",
                                                              metadataHeaders=None).execute()
                message_headers = message_data["payload"]["headers"]

                date, subject = "", ""
                for kvpair in message_headers:
                    if kvpair["name"] == "Date":
                        date = kvpair["value"]
                    if kvpair["name"] == "Subject":
                        subject = kvpair["value"]
                if self.date_is_within_range(date_start, date_end, date):
                    sent_emails.append((date, subject))

            return sent_emails

        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f"An error occurred: {error}")


    def send_message(self, subject, body = None):
        """Send an email message."""
        try:
            message = self.create_message(self.sender, self.to, subject, body)
            send_message = self.service.users().messages().send(userId="me", body=message).execute()
            print(f'Message Id: {send_message["id"]}')
        except Exception as error:
            print(f'An error occurred: {error}')


mail = HetCAKEmail(sender=os.getenv('SENDER_MAIL'), to=os.getenv('TO_MAIL'))
mail.send_message(subject="Mailonderwerpen afgelopen maand")
