import streamlit as st
from modules import gsheets, auth
from modules.docx_engine import build_docx

import pandas as pd
from datetime import datetime, date
from PIL import Image
import io

st.set_page_config(page_title="BBNT - XHH", layout="wide")
st.title("BBNT - Xã Hội Hóa (Stable)")

# ================= LOAD DATA =================
@st.cache_data(ttl=300)
def load_data():
    df_csdl, df_taichinh, _ = gsheets.load_dataframes()
    return df_csdl, df_taichinh

df_csdl, df_taichinh = load_data()
ma_tram_list = df_csdl["ma_tram"].astype(str).str.upper().tolist()

# ================= SESSION =================
st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("images_bytes", {})

# ================= LOGIN =================
if not st.session_state.logged_in:
    with st.form("login"):
        ma_tram = st.text_input("Mã trạm").upper()
        password = st.text_input("Mật khẩu", type="password")
        submit = st.form_submit_button("Đăng nhập")

    if submit:
        if ma_tram not in ma_tram_list:
            st.error("Sai mã trạm")
            st.stop()

        idx = ma_tram_list.index(ma_tram)
        if not auth.verify_password(password, str(df_csdl["Password"].iloc[idx])):
            st.error("Sai mật khẩu")
            st.stop()

        st.session_state.logged_in = True
        st.session_state.ma_tram = ma_tram
        st.rerun()

st.success(f"✅ Đăng nhập thành công: {st.session_state.ma_tram}")

# ================= INPUT DATE =================
col1, col2 = st.columns(2)
with col1:
    ngaybatdau = st.date_input("Ngày bắt đầu", value=date.today())
with col2:
    ngayketthuc = st.date_input("Ngày kết thúc", value=date.today())

# ================= UPLOAD IMAGES =================
st.subheader("📸 Upload ảnh (Anh1 → Anh8)")

for i in range(1, 9):
    file = st.file_uploader(f"Ảnh {i}", type=["jpg", "png", "jpeg"])
    if file:
        img = Image.open(file).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        st.session_state.images_bytes[f"Anh{i}"] = buf.getvalue()
        st.image(img, width=300)

# ================= CREATE DOCX =================
if st.button("📄 Tạo & tải biên bản"):
    text_map = {
        "ngaybatdau": ngaybatdau.strftime("%d/%m/%Y"),
        "ngayketthuc": ngayketthuc.strftime("%d/%m/%Y"),
    }

    docx_bytes = build_docx(
        template_path="template.docx",
        text_map=text_map,
        image_map=st.session_state.images_bytes
    )

    st.download_button(
        "📥 Tải file DOCX",
        data=docx_bytes,
        file_name=f"BBNT_{st.session_state.ma_tram}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
