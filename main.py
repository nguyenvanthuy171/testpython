from __future__ import print_function 
import base64
import os
import pandas as pd  # hiện chưa dùng nhưng cứ để nếu sau này xử lý excel

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
    # ĐỔI scope read-only -> modify để được phép xóa/move to trash
    "https://www.googleapis.com/auth/gmail.modify",
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


# ============================================
# HÀM XÓA TẤT CẢ FILE TRONG FOLDER TRÊN DRIVE
# ============================================

def delete_all_files_in_drive_folder(drive, folder_id):
    """
    Xóa (move to trash) toàn bộ file trong Google Drive folder.
    """
    query = f"'{folder_id}' in parents"
    results = drive.files().list(q=query, fields="files(id, name)").execute()

    files = results.get("files", [])
    if not files:
        print("📂 Thư mục rỗng — không có gì để xóa.")
        return

    for f in files:
        drive.files().update(
            fileId=f["id"],
            body={"trashed": True}
        ).execute()
        print(f"🗑️ Đã xóa file: {f['name']} ({f['id']})")

    print("🎉 Đã xóa toàn bộ file trong thư mục.")


# ============================================
# DOWNLOAD EXCEL FROM GMAIL
# ============================================

def download_latest_excel(gmail):
    """
    Tìm email mới nhất có file .xlsx, tải file về,
    TRẢ VỀ: (filename, msg_id)
    """
    results = gmail.users().messages().list(
        userId="me",
        q=QUERY,
        maxResults=1,
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
            return filename, msg_id

    print("Không có file .xlsx trong email")
    return None, None


# ============================================
# UPLOAD TO DRIVE
# ============================================

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


# ============================================
# DELETE GMAIL MESSAGE
# ============================================

def delete_email_with_excel(gmail, msg_id):
    if not msg_id:
        return

    gmail.users().messages().trash(userId="me", id=msg_id).execute()
    print(f"Đã chuyển email {msg_id} vào Trash.")


# ============================================
# MAIN
# ============================================

def main():
    gmail, drive = get_services()

    # 1) Xóa toàn bộ file trong folder trước khi import file mới
    print("🔄 Đang dọn thư mục Drive trước khi xử lý...")
    delete_all_files_in_drive_folder(drive, DRIVE_FOLDER_ID)

    # 2) Tải file Excel mới nhất từ Gmail
    filename, msg_id = download_latest_excel(gmail)

    if filename:
        # 3) Upload Excel lên Drive
        upload_to_drive(drive, filename)

        # 4) Xóa email chứa file Excel sau khi xử lý xong
        delete_email_with_excel(gmail, msg_id)


if __name__ == "__main__":
    main()
