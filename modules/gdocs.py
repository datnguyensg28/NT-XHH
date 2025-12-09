# modules/gdocs.py (PRO VERSION)
import io
import time
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

# === CONFIG ===
TARGET_FOLDER_ID = "1r0NCx4cIDDQ6bfS2dQPfz2zio9VYN9FH"  # thư mục Drive của bạn

# Cache service để tăng tốc
@st.cache_resource
def get_service(service_name, version="v1"):
    info = st.secrets["gcp_service_account"]
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/documents",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return build(service_name, version, credentials=creds)


# === Retry wrapper (Google API hay lag) ===
def api_retry(func, max_attempts=4, wait=0.6):
    for attempt in range(max_attempts):
        try:
            return func()
        except Exception as e:
            if attempt == max_attempts - 1:
                raise e
            time.sleep(wait)


# === MAIN FUNCTION: COPY + REPLACE ===
def copy_template_and_replace(template_doc_id: str, user_data: dict, title: str):

    drive = get_service("drive", "v3")
    docs = get_service("docs", "v1")

    # === STEP 1: COPY TEMPLATE VÀO FOLDER CỦA BẠN ===
    copied = api_retry(lambda: drive.files().copy(
        fileId=template_doc_id,
        body={
            "name": title,
            "parents": [TARGET_FOLDER_ID]
        },
        supportsAllDrives=True
    ).execute())

    new_id = copied.get("id")

    # ==== STEP 2: TÍNH TOÁN TRƯỜNG PHỤ ====
    user_data["Danh_gia_cot"] = (
        "Đạt" if user_data.get("Loai_cot") == "cột dây co" else "Không đánh giá"
    )
    user_data["Danh_gia_PM"] = (
        "Đạt" if user_data.get("Phong_may") != "Không thuê" else "Không đánh giá"
    )
    user_data["Danh_gia_DH"] = (
        "Đạt" if user_data.get("Dieu_hoa") != "Không thuê" else "Không đánh giá"
    )

    # ==== STEP 3: BATCH REPLACE PLACEHOLDERS ====
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": f"${k}", "matchCase": True},
                "replaceText": str(v),
            }
        }
        for k, v in user_data.items()
    ]

    api_retry(lambda: docs.documents().batchUpdate(
        documentId=new_id,
        body={"requests": requests}
    ).execute())

    return new_id


# === EXPORT DOCX — TỐI ƯU TẢI FILE ===
def export_docx_and_download(doc_id: str, suggested_filename: str):

    drive = get_service("drive", "v3")

    request = drive.files().export_media(
        fileId=doc_id,
        mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    fh.seek(0)
    return fh.read()


# === AUTO DELETE TEMP FILE ON DRIVE ===
def delete_drive_file(file_id: str):

    drive = get_service("drive", "v3")

    try:
        api_retry(lambda: drive.files().delete(fileId=file_id).execute())
    except Exception as e:
        st.warning(f"Không thể xóa file tạm: {e}")
