# app.py
import streamlit as st
from modules import gsheets, gdocs, auth, docx_image, utils
import io
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="BBNT - Xã Hội Hóa", layout="centered")

st.title("BBNT - Xã Hội Hóa (Web)")

# Load spreadsheet data (cached)
@st.cache_data(ttl=300)
def load_data():
    df_csdl, df_taichinh, sh = gsheets.load_dataframes(st.secrets["SPREADSHEET_URL"])
    return df_csdl, df_taichinh

try:
    df_csdl, df_taichinh = load_data()
except Exception as e:
    st.error(f"Không thể kết nối Google Sheets: {e}")
    st.stop()

# Preprocess
ma_tram_list = [str(x).strip().upper() for x in df_csdl.get("ma_tram", [])]
password_hashes = df_csdl.get("Password", [])  # in original file they are plaintext; recommend migration

# UI
with st.form("login_form"):
    col1, col2 = st.columns(2)
    with col1:
        ma_tram = st.text_input("Mã Nhà Trạm").upper().strip()
        thang_list = sorted(df_taichinh["Thang"].astype(str).unique().tolist())
        thang = st.selectbox("Tháng thanh toán", [""] + thang_list)
    with col2:
        password = st.text_input("Mật khẩu", type="password")
        submit = st.form_submit_button("Đăng nhập & Tạo biên bản")

if submit:
    if not ma_tram or not password or not thang:
        st.warning("Vui lòng nhập đầy đủ thông tin.")
    else:
        # find index
        if ma_tram in ma_tram_list:
            idx = ma_tram_list.index(ma_tram)
            stored_pw = df_csdl["Password"].iloc[idx]
            # If stored_pw looks hashed (length 64 hex) assume it's hashed; else advise migration
            if len(str(stored_pw)) == 64:
                ok = auth.verify_password(password, stored_pw)
            else:
                # legacy: compare plaintext -> recommend hashing migration
                ok = (password == str(stored_pw))
                if ok:
                    st.info("Lưu ý: mật khẩu hiện lưu plaintext trong Sheet. Nên migrate sang hash để bảo mật.")
            if not ok:
                st.error("Mật khẩu không chính xác.")
            else:
                st.success("Đăng nhập thành công!")
                # build user_data
                csdl_dict = df_csdl.iloc[idx].to_dict()
                match = df_taichinh[
                    (df_taichinh["Ma_vi_tri"].astype(str).str.upper() == ma_tram)
                    & (df_taichinh["Thang"].astype(str) == thang)
                ]
                if match.empty:
                    st.error(f"Không tìm thấy dữ liệu thanh toán cho tháng {thang}")
                else:
                    user_data = csdl_dict.copy()
                    user_data.update(match.iloc[0].to_dict())
                    user_data["Thang"] = thang

                    # show preview
                    st.subheader("Thông tin trạm")
                    st.write(pd.Series(user_data))

                    # Upload images (multiple)
                    st.info("Upload tối đa 8 ảnh (theo thứ tự ${Anh1} ... ${Anh8}).")
                    uploaded_files = st.file_uploader("Upload ảnh", type=["jpg","jpeg","png"], accept_multiple_files=True)

                    if st.button("Tạo & Tải biên bản"):
                        with st.spinner("Đang tạo tài liệu..."):
                            title = f"BBNT_{ma_tram}_{thang}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                            template_id = st.secrets["TEMPLATE_DOC_ID"]
                            # create doc on Drive and replace tags
                            try:
                                doc_id = gdocs.copy_template_and_replace(template_id, user_data, title)
                                docx_bytes = gdocs.export_docx_and_download(doc_id, f"{title}.docx")
                                # insert images into docx bytes
                                placeholders = [f"${{Anh{i}}}" for i in range(1,9)]
                                # map uploaded files by order
                                for i, file in enumerate(uploaded_files[:8]):
                                    try:
                                        img_bytes = file.read()
                                        docx_bytes = docx_image.insert_image_into_docx_bytes(docx_bytes, placeholders[i], img_bytes)
                                    except Exception as e:
                                        st.warning(f"Lỗi chèn ảnh {i+1}: {e}")
                                # provide download
                                st.download_button(
                                    label="📥 Tải biên bản (docx)",
                                    data=docx_bytes,
                                    file_name=f"{title}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            except Exception as e:
                                st.error(f"Lỗi tạo tài liệu: {e}")
                            finally:
                                # delete temp doc on drive to avoid clutter
                                try:
                                    gdocs.delete_drive_file(doc_id)
                                except Exception:
                                    pass
