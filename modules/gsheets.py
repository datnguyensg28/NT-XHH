import gspread
import pandas as pd
import streamlit as st
from google.oauth2 import service_account

# Load từ block secrets
gcp = st.secrets["gcp_service_account"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]

def get_creds_from_secrets():
    creds = service_account.Credentials.from_service_account_info(
        gcp,
        scopes=SCOPES
    )
    return creds

def open_spreadsheet():
    spreadsheet_url = gcp["SPREADSHEET_URL"]
    creds = get_creds_from_secrets()
    client = gspread.authorize(creds)
    sh = client.open_by_url(spreadsheet_url)
    return sh

def load_dataframes():
    sh = open_spreadsheet()

    sheet_csdl = sh.worksheet("CSDL")
    sheet_taichinh = sh.worksheet("Taichinh")

    df_csdl = pd.DataFrame(sheet_csdl.get_all_records())
    df_taichinh = pd.DataFrame(sheet_taichinh.get_all_records())

    df_csdl.columns = [c.strip() for c in df_csdl.columns]
    df_taichinh.columns = [c.strip() for c in df_taichinh.columns]

    return df_csdl, df_taichinh, sh
