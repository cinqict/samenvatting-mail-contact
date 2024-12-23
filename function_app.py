import logging
import azure.functions as func

import os
import base64
import calendar
import datetime as dt
import pickle
import json

from dotenv import load_dotenv

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parsedate_tz, mktime_tz

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

import pandas as pd

load_dotenv()

app = func.FunctionApp()

# TODO set run_on_startup to false when going to production/using cak mails
@app.timer_trigger(schedule="0 0 8 1 * *", arg_name="mailTimer", run_on_startup=True,
              use_monitor=False) 
def cak_communicatie_mail(mailTimer: func.TimerRequest) -> None:
    mail = Mail(os.environ["MAIL_SENDER"], os.environ["MAIL_TO"])
    subject = "Overzicht mailuitwisseling afgelopen maand"
    logging.info(f"Sending mail from {str(mail.sender)} to {str(mail.to)} with subject {subject}")
    mail.send_message(subject=subject)


# The Gmail API scope needed for sending emails
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.readonly"]

class Mail:
    def __init__(self, sender_mail: str, to_mail: str):
        """
        Create and send an e-mail with an overview of the emails that were found in a given folder of the user's e-mail.

        The script looks at the contents of the e-mail from the authenticated user. This user's email must be the same
        as the sender_mail. Default the script looks at the content of the "INBOX" folder.

        Given an example sender_mail and a to_mail, and a subject, the following mail could be sent:
        E.g.:
        ```
        Beste,

        Over de afgelopen periode (Tue, 01 Oct 2024 t/m Thu, 31 Oct 2024) zijn de volgende 2 mails verzonden.

        Datum	                        Onderwerp
        Mon, 28 Oct 2024 14:13:56 +0100	Formulier ingevuld
        Mon, 14 Oct 2024 13:04:14 +0200	Re: Welkomstmail

        Met vriendelijke groet,
        CINQ ICT
        ```

        It is possible to override the default content of the e-mail that is sent with a custom text. However, you must
        then also construct your own table of information on the contents of the selected e-mail folder(s).

        Parameters
        ----------
        sender_mail :    str
                    E-mail address from which the folders will be inspected, and from which the mail will be sent from.
        to_mail :        str
                    E-mail address to send the constructed email to.
        """
        self.service = self.authenticate_gmail()
        self.sender = sender_mail
        self.to = to_mail

    @staticmethod
    def authenticate_gmail():
        """
        Authenticate and create a Gmail object for sending emails.

        Returns
        -------
        service : object through which can be interacted with the Gmail API.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens.
        # It is created automatically when the authorization flow completes for the first time.
        access_token = os.getenv('GMAIL_ACCESS_TOKEN')
        refresh_token= os.getenv('GMAIL_REFRESH_TOKEN')
        client_id = os.getenv('GMAIL_CLIENT_ID')
        client_secret= os.getenv('GMAIL_CLIENT_SECRET')
        if access_token and refresh_token and client_id and client_secret:
            token_uri = "https://oauth2.googleapis.com/token"  # Google's token endpoint
 
            # Create a Credentials object manually using the stored access_token and refresh_token
            creds = Credentials(
                token=access_token,
                refresh_token=refresh_token,
                client_id=client_id,
                client_secret=client_secret,
                token_uri=token_uri
            )

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
        """
        Get the start and end dates from the previous month, based on the current month.

        Returns
        -------
        start_date :   datetime.datetime
                            The end date of the previous month, of the format e.g. "Mon, 01 Jan 2024"
        end_date :     datetime.datetime
                            The end date of the previous month, of the format e.g. "Mon, 01 Jan 2024"
        """
        today = dt.datetime.today()
        current_year = today.year
        current_month = today.month

        # If we're in January, go back to_mail December of the previous year
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

        # Convert to_mail timestamps of the format e.g. "Mon, 01 Jan 2024"
        start_date = first_day_of_previous_month.date().strftime("%a, %d %b %Y")
        end_date = last_day_of_previous_month.date().strftime("%a, %d %b %Y")

        return start_date, end_date

    def create_message(self, sender: str, to: str, subject: str, body: str = None):
        """
        Create an email message.

        Parameters
        ----------
        sender :    str
                    The email address to_mail send the mail from.
        to :        str
                    The email address to_mail send the mail to_mail.
        subject :
                    The email address to_mail send the mail to_mail.
        body :      str, optional
                    The text to_mail use as the body of the e-mail. See the contents of the function for the default body.

        Returns
        -------
        dict :  dictionary of a single base64 encoded string holding the message

        """
        message = MIMEMultipart()
        message['to'] = to
        message['from'] = sender
        message['subject'] = subject

        date_start, date_end = self.get_previous_month_start_end()
        sent_emails = self.get_emails(date_start, date_end)
        sent_emails_df = pd.DataFrame(sent_emails, columns=["Datum", "Onderwerp"])
        sent_emails_df = sent_emails_df.set_index("Datum")
        sent_emails_df.sort_index(inplace=True)

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
    def date_is_within_range(date_start: dt.datetime.date, date_end: dt.datetime.date, date_to_check:  dt.datetime.date) -> bool:
        """
        Check if a given date is within the passed daterange.

        Parameters
        ----------
        date_start :    dt.datetime.date
                        The start of the daterange.
        date_end :      dt.datetime.date
                        The end of the daterange.
        date_to_check : dt.datetime.date
                        The date to_mail be checked.

        Returns
        -------
        bool : True if the date_to_check falls in the given daterange. Otherwise False.
        """
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
        """
        Retrieve all emails from the given daterange that are present in the requested folder(s).

        Parameters
        ----------
        date_start :    datetime.datetime.date, optional
                        The start of the date-range within the e-mail's timestamp must fall in.
        date_end :      datetime.datetime.date, optional
                        The end of the date-range within the e-mail's timestamp must fall in.
        folders :       list of str, default=["INBOX"]
                        The folders to_mail scan to_mail retrieve the mails.

        Returns
        -------
        sent_emails :   list of tuples
                        A list of tuples, where each tuple contains the e-mail's timestamp and the subject.
        """
        try:
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


    def send_message(self, subject: str, body: str = None) -> None:
        """
        Send an e-mail message.

        Parameters
        ----------
        subject :   str
                    The text to_mail use as the subject of the e-mail.
        body :  str
                The text to_mail use as the body of the e-mail.
        """
        try:
            message = self.create_message(self.sender, self.to, subject, body)
            send_message = self.service.users().messages().send(userId="me", body=message).execute()
            print(f'Message Id: {send_message["id"]}')
            logging.info("Message sent!")
        except Exception as error:
            print(f'An error occurred: {error}')
