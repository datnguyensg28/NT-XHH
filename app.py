# app.py
import streamlit as st
from modules import gsheets, auth, docx_image
import pandas as pd
from datetime import datetime
from PIL import Image
import io
import re
import zipfile

st.set_page_config(page_title="BBNT - Xã Hội Hóa V4", layout="wide")
st.title("BBNT - Xã Hội Hóa (Web V4)")

# ---------- load data ----------
@st.cache_data(ttl=300)
def load_data():
    df_csdl, df_taichinh, _ = gsheets.load_dataframes()
    return df_csdl, df_taichinh

try:
    df_csdl, df_taichinh = load_data()
except Exception as e:
    st.error(f"Không thể kết nối Google Sheets: {e}")
    st.stop()

ma_tram_list = [str(v).strip().upper() for v in df_csdl["ma_tram"]]

# session
st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("images", {})
st.session_state.setdefault("images_bytes", {})

def bytes_from_pil(img: Image.Image):
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()

# ---------- helpers ----------
from modules.docx_image import _merge_xml  # vẫn dùng

def extract_placeholders_from_docx_bytes(docx_bytes: bytes):
    """
    Trả về set placeholder dạng 'ngaybatdau', 'anh1', 'tienthangtruocthue', ...
    tìm cả $xxx và ${xxx} và trả về tên không đổi (giữ nguyên case có trong file).
    """
    bio = io.BytesIO(docx_bytes)
    with zipfile.ZipFile(bio, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8")

    xml = _merge_xml(xml)

    # tìm $xxx và ${xxx}
    s = set()
    for m in re.finditer(r"\$\{?\s*([A-Za-z0-9_]+)\s*\}?", xml):
        s.add(m.group(1))

    return s

# ---------- login ----------
if not st.session_state.logged_in:
    with st.form("login_form"):
        col1, col2 = st.columns(2)
        with col1:
            ma_tram = st.text_input("Mã Nhà Trạm").upper().strip()
            list_thang = sorted(df_taichinh["Thang"].astype(str).unique().tolist())
            thang = st.selectbox("Tháng thanh toán", [""] + list_thang)
        with col2:
            password = st.text_input("Mật khẩu", type="password")
        submit = st.form_submit_button("Đăng nhập")

    if submit:
        if not ma_tram:
            st.warning("Nhập mã trạm!")
            st.stop()
        if ma_tram not in ma_tram_list:
            st.error("Sai mã trạm!")
            st.stop()
        idx = ma_tram_list.index(ma_tram)
        stored_pw = str(df_csdl["Password"].iloc[idx])
        ok = (auth.verify_password(password, stored_pw) if len(stored_pw) == 64 else stored_pw == password)
        if not ok:
            st.error("Sai mật khẩu.")
            st.stop()
        st.session_state.logged_in = True
        st.session_state.ma_tram = ma_tram
        st.session_state.thang = thang
        st.session_state.images = {}
        st.session_state.images_bytes = {}
        st.rerun()

if not st.session_state.logged_in:
    st.stop()

ma_tram = st.session_state.ma_tram
thang = st.session_state.thang
idx = ma_tram_list.index(ma_tram)
csdl_dict = df_csdl.iloc[idx].to_dict()

match = df_taichinh[
    (df_taichinh["Ma_vi_tri"].astype(str).str.upper() == ma_tram)
    &
    (df_taichinh["Thang"].astype(str) == thang)
]

if match.empty:
    st.error("Không tìm thấy dữ liệu tháng.")
    st.stop()

user_data = csdl_dict.copy()
user_data.update(match.iloc[0].to_dict())
user_data["Thang"] = thang

# auto fields
loai_cot = str(user_data.get("Loai_cot", "")).strip().lower()
user_data["Danh_gia_cot"] = "Đạt" if loai_cot == "cột dây co" else "Không đánh giá"
user_data["Danh_gia_PM"] = "Đạt" if str(user_data.get("Phong_may","")) != "Không thuê" else "Không đánh giá"
user_data["Danh_gia_DH"] = "Đạt" if str(user_data.get("Dieu_hoa","")) != "Không thuê" else "Không đánh giá"

st.subheader("Thông tin trạm")
st.write(pd.Series(user_data))
st.markdown("---")

# upload & rotate
st.subheader("📸 Upload & Xoay ảnh (1–8)")
labels = [
    "Anh1 – Toàn cảnh cột anten",
    "Anh2 – Móng M0",
    "Anh3 – Móng M1",
    "Anh4 – Móng M2",
    "Anh5 – Móng M3",
    "Anh6 – Anten & RRU",
    "Anh7 – Phòng máy ngoài→vào",
    "Anh8 – Phòng máy trong→ra"
]

def do_rotate(idx, angle):
    key = f"img{idx}"
    if key in st.session_state.images:
        img = st.session_state.images[key]
        rotated = img.rotate(angle, expand=True)
        st.session_state.images[key] = rotated
        st.session_state.images_bytes[key] = bytes_from_pil(rotated)

for i, label in enumerate(labels, start=1):
    key = f"img{i}"
    st.markdown(f"### {label}")
    file = st.file_uploader(label, type=["jpg","jpeg","png"], key=f"u{i}")
    if file and key not in st.session_state.images:
        img = Image.open(file).convert("RGB")
        img.thumbnail((1600,1600))
        st.session_state.images[key] = img
        st.session_state.images_bytes[key] = bytes_from_pil(img)
    if key in st.session_state.images:
        col1, col2, col3 = st.columns([4,1,1])
        with col1:
            st.image(st.session_state.images[key], width=450)
        with col2:
            st.button("⟲", key=f"L{i}", on_click=do_rotate, args=(i, 90))
        with col3:
            st.button("⟳", key=f"R{i}", on_click=do_rotate, args=(i, -90))
    st.markdown("---")

# ---------- CREATE REPORT ----------
# ---------- CREATE REPORT ----------
from modules import docx_image_safe as ds

def convert_date_value(value):
    """Tự nhận dạng:
    - DD/MM/YYYY
    - YYYY-MM-DD
    - Excel serial number
    """
    import pandas as pd

    if value is None:
        return ""

    # Excel serial numbers
    if isinstance(value, (int, float)) and value > 25000:
        try:
            base = pd.to_datetime("1899-12-30")
            dt = base + pd.to_timedelta(int(value), "D")
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%Y")
    except:
        pass

    return str(value)


if st.button("📄 Tạo & Tải biên bản"):
    try:
        with st.spinner("Đang tạo biên bản..."):

            # Load template docx
            with open("template.docx", "rb") as f:
                docx_bytes = f.read()

            doc = ds.load_docx_bytes(docx_bytes)

            # Chuẩn hoá dữ liệu user_data
            normalized_map = {}

            for k, v in user_data.items():
                key = k.lower().replace("_","")  # ví dụ ngaybatdau
                normalized_map[key] = convert_date_value(v)

            # Danh sách placeholders trong template:
            holders = [
                # text placeholders
                "ngaybatdau", "ngayketthuc", "ngay_ky",
                "Chu_ha_tang", "Chuc_vu",
                "Danh_gia_DH","Danh_gia_PM","Danh_gia_cot",
                "Dia_chi","Ky_thanh_toan",
                "Loai_cot","Loai_tram",
                "Ma_HD","Ten_GD_VT","Ten_don_vi_XHH",
                "ma_tram","tien_bang_chu",
                "tienthangtruocthue","tienthueky",
                "tientruocthueky","tongtienky"
            ]

            # --- Replace text placeholders ---
            for holder in holders:
                key_norm = holder.lower().replace("_","")
                value = normalized_map.get(key_norm, "")

                # replace 4 dạng
                ds.replace_text(doc, f"${holder}", value)
                ds.replace_text(doc, f"${{{holder}}}", value)
                ds.replace_text(doc, f"${holder};", value)
                ds.replace_text(doc, f"${{{holder}}};", value)

            # --- Insert images ---
            for i in range(1, 9):
                key = f"img{i}"
                if key in st.session_state.images_bytes:
                    img_bytes = st.session_state.images_bytes[key]

                    # 4 dạng placeholder
                    for ph in [
                        f"${{Anh{i}}}",
                        f"$Anh{i}",
                        f"${{anh{i}}}",
                        f"$anh{i}"
                    ]:
                        ds.insert_image(doc, ph, img_bytes, 12)

            # Save lại DOCX
            out_bytes = ds.save_docx(doc)

            title = f"BBNT_{ma_tram}_{thang}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

            st.download_button(
                "📥 Tải DOCX",
                data=out_bytes,
                file_name=title + ".docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        import traceback
        st.error(f"Lỗi tạo biên bản: {e}")
        st.text(traceback.format_exc())
