from __future__ import print_function
import base64
import os
import pandas as pd

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================
# CONFIG
# ============================================
QUERY = "has:attachment filename:xlsx"  # Gmail search
DRIVE_FOLDER_ID = "12t79mueDBK6F7wtfRHIputY0IQ2jiJ9w"   # Folder Google Drive

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ============================================

def get_services():
    # Lấy thông tin OAuth từ biến môi trường (đã set qua GitHub Secrets)
    client_id = os.environ["GCP_CLIENT_ID"]
    client_secret = os.environ["GCP_CLIENT_SECRET"]
    refresh_token = os.environ["GCP_REFRESH_TOKEN"]

    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=SCOPES,
    )

    # Refresh để lấy access token mới
    creds.refresh(Request())

    gmail = build("gmail", "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return gmail, drive


def download_latest_excel(gmail):
    results = gmail.users().messages().list(
        userId="me",
        q=QUERY,
        maxResults=1,
    ).execute()

    messages = results.get("messages", [])
    if not messages:
        print("Không tìm thấy email có file Excel")
        return None

    msg_id = messages[0]["id"]
    message = gmail.users().messages().get(userId="me", id=msg_id).execute()

    parts = message["payload"].get("parts", [])
    for part in parts:
        filename = part.get("filename", "")
        if filename.endswith(".xlsx"):
            body = part.get("body", {})
            att_id = body.get("attachmentId")
            if not att_id:
                continue

            att = gmail.users().messages().attachments().get(
                userId="me",
                messageId=msg_id,
                id=att_id,
            ).execute()

            data = base64.urlsafe_b64decode(att["data"])
            with open(filename, "wb") as f:
                f.write(data)

            print("Đã tải file:", filename)
            return filename

    print("Không có file .xlsx trong email")
    return None


def upload_to_drive(drive, filename):
    file_metadata = {
        "name": filename,
        "parents": [DRIVE_FOLDER_ID],
    }

    media = MediaFileUpload(filename, resumable=True)
    uploaded = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
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
