# app.py
import streamlit as st
from modules import gsheets, auth
from modules import docx_image_safe as ds   # SAFE MODE DOCX
import pandas as pd
from datetime import datetime
from PIL import Image
import io
import re
import zipfile

# -------------------- SETTINGS --------------------
st.set_page_config(page_title="BBNT - Xã Hội Hóa V3", layout="wide")
st.title("BBNT - Xã Hội Hóa (Web V3)")

# -------------------- LOAD DATA --------------------
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

# -------------------- SESSION --------------------
st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("images", {})
st.session_state.setdefault("images_bytes", {})

def bytes_from_pil(img: Image.Image):
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()

# -------------------- LOGIN --------------------
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
        ok = (
            auth.verify_password(password, stored_pw)
            if len(stored_pw) == 64
            else stored_pw == password
        )
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

# -------------------- GET USER ROW --------------------
ma_tram = st.session_state.ma_tram
thang = st.session_state.thang
idx = ma_tram_list.index(ma_tram)
csdl_dict = df_csdl.iloc[idx].to_dict()

match = df_taichinh[
    (df_taichinh["Ma_vi_tri"].astype(str).str.upper() == ma_tram) &
    (df_taichinh["Thang"].astype(str) == thang)
]

if match.empty:
    st.error("Không tìm thấy dữ liệu tháng.")
    st.stop()

user_data = csdl_dict.copy()
user_data.update(match.iloc[0].to_dict())
user_data["Thang"] = thang

# -------------------- AUTO FIELDS --------------------
loai_cot = str(user_data.get("Loai_cot", "")).strip().lower()
user_data["Danh_gia_cot"] = "Đạt" if loai_cot == "cột dây co" else "Không đánh giá"
user_data["Danh_gia_PM"] = "Đạt" if str(user_data.get("Phong_may","")) != "Không thuê" else "Không đánh giá"
user_data["Danh_gia_DH"] = "Đạt" if str(user_data.get("Dieu_hoa","")) != "Không thuê" else "Không đánh giá"

st.subheader("Thông tin trạm")
st.write(pd.Series(user_data))
st.markdown("---")

# -------------------- UPLOAD & ROTATE --------------------
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

# -------------------------------------------------------
# -------- SAFE FORMATTING ENGINE (DỮ LIỆU NGÀY & TIỀN) -----
# -------------------------------------------------------

DATE_KEYS = {"ngaybatdau", "ngayketthuc", "ngay_ky", "ngayky"}

def is_likely_money_key(key: str):
    k = key.lower()
    return any(tok in k for tok in ["tien","tong","gia","giatri","phi","tax","thue"])

def format_number_vn(x):
    """Chuyển số thành dạng 5.500.000"""
    try:
        if isinstance(x, str):
            s = x.replace(".", "").replace(",", "").strip()
            if not s.isdigit():
                return x
            x = float(s)
        val = float(x)
    except:
        return str(x)

    if abs(val - int(val)) < 1e-9:
        return f"{int(val):,}".replace(",", ".")
    else:
        s = f"{val:,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")

def convert_if_date(key_norm: str, value):
    """Chỉ convert nếu key nằm trong DATE_KEYS."""
    import pandas as pd
    if key_norm not in DATE_KEYS:
        return None

    if isinstance(value, (int, float)) and value > 25000:
        base = pd.to_datetime("1899-12-30")
        dt = base + pd.to_timedelta(int(value), "D")
        return dt.strftime("%d/%m/%Y")

    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%Y")
    except:
        pass

    return None


# -------------------------------------------------------
# -------------------- CREATE REPORT ---------------------
# -------------------------------------------------------
if st.button("📄 Tạo & Tải biên bản"):
    try:
        with st.spinner("Đang tạo biên bản..."):

            # Load template
            with open("template.docx", "rb") as f:
                docx_bytes = f.read()

            doc = ds.load_docx_bytes(docx_bytes)

            # Build normalized_map (AN TOÀN)
            normalized_map = {}

            for k, v in user_data.items():
                key_norm = k.lower().replace("_","")
                sval = "" if v is None else str(v).strip()

                # 1) DATE
                date_val = convert_if_date(key_norm, v)
                if date_val:
                    normalized_map[key_norm] = date_val
                    continue

                # 2) MONEY
                if is_likely_money_key(key_norm):
                    normalized_map[key_norm] = format_number_vn(v)
                    continue

                # 3) OTHER NUMBERS
                if isinstance(v, (int, float)):
                    normalized_map[key_norm] = format_number_vn(v)
                    continue

                # 4) TRY parse date-like strings
                if any(x in sval for x in ["/","-"]) and re.search(r"\d{4}", sval):
                    try:
                        dt = pd.to_datetime(sval, dayfirst=True, errors="coerce")
                        if not pd.isna(dt):
                            normalized_map[key_norm] = dt.strftime("%d/%m/%Y")
                            continue
                    except:
                        pass

                # 5) DEFAULT
                normalized_map[key_norm] = sval

            # Replace TEXT placeholders
            for holder, val in normalized_map.items():
                for ph in [f"${holder}", f"${{{holder}}}", f"${holder};", f"${{{holder}}};"]:
                    ds.replace_text(doc, ph, val)

            # Insert IMAGES
            for i in range(1, 9):
                key = f"img{i}"
                if key in st.session_state.images_bytes:
                    img_bytes = st.session_state.images_bytes[key]
                    for ph in [f"${{Anh{i}}}", f"$Anh{i}", f"${{anh{i}}}", f"$anh{i}"]:
                        ds.insert_image(doc, ph, img_bytes, 12)

            # Save DOCX
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
