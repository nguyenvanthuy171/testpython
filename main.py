from __future__ import print_function
import base64
import os
import pandas as pd

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================
# CONFIG
# ============================================
QUERY = "has:attachment filename:xlsx"     # Gmail search
DRIVE_FOLDER_ID = "12t79mueDBK6F7wtfRHIputY0IQ2jiJ9w"   # Folder Google Drive
SERVICE_ACCOUNT_FILE = "service_account.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# ============================================

def get_services():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gmail = build("gmail", "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return gmail, drive


def download_latest_excel(gmail):
    results = gmail.users().messages().list(
        userId='me',
        q=QUERY,
        maxResults=1
    ).execute()

    messages = results.get('messages', [])
    if not messages:
        print("Không tìm thấy email có file Excel")
        return None

    msg_id = messages[0]["id"]
    message = gmail.users().messages().get(userId="me", id=msg_id).execute()

    parts = message["payload"].get("parts", [])
    for part in parts:
        if part["filename"].endswith(".xlsx"):
            att_id = part["body"]["attachmentId"]
            att = gmail.users().messages().attachments().get(
                userId="me",
                messageId=msg_id,
                id=att_id
            ).execute()

            data = base64.urlsafe_b64decode(att["data"])
            filename = part["filename"]

            with open(filename, "wb") as f:
                f.write(data)

            print("Đã tải file:", filename)
            return filename

    print("Không có file .xlsx trong email")
    return None


def upload_to_drive(drive, filename):
    file_metadata = {
        "name": filename,
        "parents": [DRIVE_FOLDER_ID]
    }

    media = MediaFileUpload(filename, resumable=True)
    uploaded = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    print("Đã upload vào Google Drive, file ID:", uploaded.get("id"))
    return uploaded.get("id")


def main():
    gmail, drive = get_services()

    filename = download_latest_excel(gmail)

    if filename:
        upload_to_drive(drive, filename)


if __name__ == "__main__":
    main()
