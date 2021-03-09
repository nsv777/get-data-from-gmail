import base64
import os
import pickle
import re

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from classes.excel_file import ExcelFile


class Gmail(object):
    # https://developers.google.com/gmail/api/quickstart/python
    def __init__(self, api_scopes=None, credentials_file="credentials.json"):
        if not api_scopes:
            api_scopes = ["https://www.googleapis.com/auth/gmail.readonly"]
        self.service = build(
            serviceName='gmail',
            version='v1',
            credentials=self._get_credentials(api_scopes, credentials_file=credentials_file)
        )

    @staticmethod
    def _get_credentials(api_scopes, credentials_file="credentials.json"):
        credentials = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                credentials = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not credentials or not credentials.valid:
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, api_scopes)
                # credentials = flow.run_local_server(port=0)
                credentials = flow.run_console()
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(credentials, token)

        return credentials

    def get_messages(self, query=None):
        if query is None:
            return None

        messages = list()
        results = self.service.users().messages().list(
            userId='me',
            q=query,
            maxResults=300
        ).execute()
        messages += results.get("messages", None)

        page_token = results.get('nextPageToken', None)
        while page_token:
            results = self.service.users().messages().list(
                userId='me',
                q=query,
                maxResults=300,
                pageToken=page_token).execute()
            messages += results.get("messages")
            page_token = results.get('nextPageToken', None)
        return messages

    def get_message_body(self, message_id):
        message = self.service.users().messages().get(userId='me', id=message_id).execute()
        try:
            body_data = message.get("payload").get("parts")[0].get("body").get("data")
        except TypeError:
            # Not forwarded message
            body_data = message.get("payload").get("body").get("data")
        body = base64.urlsafe_b64decode(body_data.encode("ASCII")).decode("utf-8")
        # for body_part in message.get("payload").get("parts"):
        #     body += base64.urlsafe_b64decode(body_part.get("body").get("data").encode("ASCII")).decode("utf-8")
        return body

    @staticmethod
    def get_row(body, body_patterns):
        if '<br>' in body:  # html email
            soup = BeautifulSoup(body, "lxml")
            body = soup.get_text(separator='\n')
        message_data = dict()
        for line in body.splitlines():
            for body_pattern in body_patterns:
                if not message_data.get(body_pattern[0], None):
                    match = re.search(body_pattern[1], line, flags=re.IGNORECASE)
                    if match:
                        message_data[body_pattern[0]] = match.group(1)
                    else:
                        message_data[body_pattern[0]] = ""
        return message_data

    def process_messages(self, query, body_patterns):
        excel_file = ExcelFile()
        is_header = False
        for message in self.get_messages(query=query):
            row = self.get_row(self.get_message_body(message_id=message["id"]), body_patterns=body_patterns)
            if not is_header:
                excel_file.add_row(list(row.keys()))
                is_header = True
            excel_file.add_row(list(row.values()))
            print(row)
        excel_file.save()
