# modules/gdocs.py
import io, time
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import streamlit as st
from google.oauth2 import service_account

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

def copy_template_and_replace(template_doc_id: str, user_data: dict, title: str):
    drive = get_service("drive", "v3")
    docs = get_service("docs", "v1")
    copied = drive.files().copy(
        fileId=template_doc_id,
        body={"name": title},
        supportsAllDrives=True
    ).execute()
    new_id = copied.get("id")

    # prepare computed fields (mirror logic from original)
    if user_data.get("Loai_cot") == "cột dây co":
        user_data["Danh_gia_cot"] = "Đạt"
    else:
        user_data["Danh_gia_cot"] = "Không đánh giá"

    user_data["Danh_gia_PM"] = "Đạt" if user_data.get("Phong_may") != "Không thuê" else "Không đánh giá"
    user_data["Danh_gia_DH"] = "Đạt" if user_data.get("Dieu_hoa") != "Không thuê" else "Không đánh giá"

    # batch replace
    requests = []
    for k, v in user_data.items():
        requests.append({
            "replaceAllText": {
                "containsText": {"text": f"${k}", "matchCase": True},
                "replaceText": str(v)
            }
        })
    docs.documents().batchUpdate(documentId=new_id, body={"requests": requests}).execute()
    return new_id

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
    # Return bytesIO for Streamlit to serve
    return fh.read()

def delete_drive_file(file_id: str):
    drive = get_service("drive", "v3")
    try:
        drive.files().delete(fileId=file_id).execute()
    except Exception as e:
        st.warning(f"Không xóa được file trên Drive: {e}")
