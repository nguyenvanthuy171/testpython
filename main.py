from __future__ import print_function
import base64
import os

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================
# CONFIG
# ============================================

# Gmail search query
QUERY = "has:attachment filename:xlsx"

# Google Drive folder ID (nơi sẽ upload file)
DRIVE_FOLDER_ID = "12t79mueDBK6F7wtfRHIputY0IQ2jiJ9w"

# OAuth scopes (đủ quyền Gmail + Drive)
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


# ============================================
# AUTH
# ============================================

def get_services():
    """Load credentials từ GitHub Secrets và tạo service Gmail + Drive."""

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

    # Refresh để lấy access token
    creds.refresh(Request())

    gmail = build("gmail", "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)

    return gmail, drive


# ============================================
# GMAIL DOWNLOAD
# ============================================

def download_latest_excel(gmail):
    """
    Tìm email mới nhất có file .xlsx và tải file về.
    Trả về: (filename, msg_id)
    """

    results = gmail.users().messages().list(
        userId="me",
        q=QUERY,
        maxResults=1
    ).execute()

    messages = results.get("messages", [])

    if not messages:
        print("Không tìm thấy email có file Excel")
        return None, None

    msg_id = messages[0]["id"]
    message = gmail.users().messages().get(userId="me", id=msg_id).execute()

    parts = message["payload"].get("parts", [])

    for part in parts:
        filename = part.get("filename", "")

        if filename.endswith(".xlsx"):
            att_id = part["body"].get("attachmentId")
            if not att_id:
                continue

            att = gmail.users().messages().attachments().get(
                userId="me",
                messageId=msg_id,
                id=att_id
            ).execute()

            data = base64.urlsafe_b64decode(att["data"])

            # Lưu file
            with open(filename, "wb") as f:
                f.write(data)

            print("Đã tải file:", filename)
            return filename, msg_id

    print("Không có file .xlsx trong email")
    return None, None


# ============================================
# DRIVE UPLOAD
# ============================================

def upload_to_drive(drive, filename):
    """Upload file lên Google Drive folder chỉ định."""

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

    file_id = uploaded.get("id")
    print("Đã upload file vào Drive, ID:", file_id)

    return file_id


# ============================================
# DELETE EMAIL
# ============================================

def delete_email_with_excel(gmail, msg_id):
    """Chuyển email có file Excel vào Trash."""

    if not msg_id:
        return

    gmail.users().messages().trash(
        userId="me",
        id=msg_id
    ).execute()

    print(f"Đã chuyển email {msg_id} vào Trash.")


# ============================================
# MAIN
# ============================================

def main():
    gmail, drive = get_services()

    filename, msg_id = download_latest_excel(gmail)

    if filename:
        upload_to_drive(drive, filename)
        delete_email_with_excel(gmail, msg_id)


if __name__ == "__main__":
    main()
