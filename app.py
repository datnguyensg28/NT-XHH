# app.py
import streamlit as st
from modules import gsheets, auth, docx_image
import pandas as pd
from datetime import datetime
from pathlib import Path
from PIL import Image
import io
import json
import re
import zipfile

PROJECT_DIR = Path(__file__).resolve().parent
SECRETS_PATH = PROJECT_DIR / ".streamlit" / "secrets.toml"
SECRETS_EXAMPLE_PATH = PROJECT_DIR / ".streamlit" / "secrets.example.toml"

st.set_page_config(page_title="BBNT - Xã Hội Hóa", layout="wide")
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.25rem; max-width: 1200px; }
    .app-shell {
        background: linear-gradient(135deg, #f6f8fb 0%, #ffffff 55%, #eef5f3 100%);
        border: 1px solid #e4e8ee;
        border-radius: 8px;
        padding: 1.25rem 1.4rem;
        margin-bottom: 1.1rem;
    }
    .app-title { color: #17202a; font-size: 2rem; font-weight: 780; margin-bottom: .2rem; }
    .app-subtitle { color: #5f6368; margin-bottom: 0; max-width: 780px; }
    .quick-steps {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: .75rem;
        margin: .5rem 0 1rem 0;
    }
    .quick-step {
        background: #ffffff;
        border: 1px solid #e4e8ee;
        border-radius: 8px;
        padding: .85rem 1rem;
    }
    .quick-step strong { color: #17202a; display: block; margin-bottom: .15rem; }
    .quick-step span { color: #667085; font-size: .9rem; }
    .friendly-panel {
        background: #ffffff;
        border: 1px solid #e4e8ee;
        border-radius: 8px;
        padding: 1rem 1.15rem;
        margin-bottom: .85rem;
    }
    .friendly-title { font-weight: 750; color: #17202a; font-size: 1.08rem; margin-bottom: .2rem; }
    .friendly-muted { color: #667085; font-size: .92rem; line-height: 1.45; }
    .required-badge {
        display: inline-block;
        color: #b42318;
        background: #fff1f0;
        border: 1px solid #ffccc7;
        border-radius: 999px;
        padding: .1rem .5rem;
        font-size: .78rem;
        font-weight: 700;
        margin-left: .35rem;
    }
    .done-badge {
        display: inline-block;
        color: #067647;
        background: #ecfdf3;
        border: 1px solid #abefc6;
        border-radius: 999px;
        padding: .1rem .5rem;
        font-size: .78rem;
        font-weight: 700;
        margin-left: .35rem;
    }
    .setup-card {
        background: #ffffff;
        border: 1px solid #e1e6ef;
        border-radius: 8px;
        padding: 1rem 1.1rem;
        height: 100%;
    }
    .setup-title { color: #17202a; font-weight: 700; font-size: 1.05rem; margin-bottom: .25rem; }
    .setup-muted { color: #5f6368; font-size: .92rem; line-height: 1.45; }
    .status-pill {
        display: inline-block;
        border-radius: 999px;
        padding: .25rem .65rem;
        font-size: .8rem;
        font-weight: 700;
        background: #fff3cd;
        color: #7a4d00;
        border: 1px solid #ffe08a;
        margin-bottom: .6rem;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e6e8ec;
        border-radius: 8px;
        padding: .75rem 1rem;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #c9ced6;
        border-radius: 8px;
        padding: .35rem .75rem;
    }
    @media (max-width: 760px) {
        .quick-steps { grid-template-columns: 1fr; }
        .app-title { font-size: 1.55rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div class="app-shell">
        <div class="app-title">BBNT - Xã Hội Hóa</div>
        <div class="app-subtitle">Cảm ơn quý vị đã phối hợp và hỗ trợ Viettel trong suốt thời gian qua</div>
    </div>
    """,
    unsafe_allow_html=True,
)


def render_connection_error(error):
    st.markdown('<span class="status-pill">Chưa kết nối Google Sheets</span>', unsafe_allow_html=True)
    st.error(f"Không thể kết nối Google Sheets: {error}")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            f"""
            <div class="setup-card">
                <div class="setup-title">1. Vị trí file cấu hình</div>
                <div class="setup-muted">Tạo file <b>secrets.toml</b> trong thư mục <b>.streamlit</b> của project.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.code(str(SECRETS_PATH), language="text")
    with col2:
        st.markdown(
            """
            <div class="setup-card">
                <div class="setup-title">2. Điền service account</div>
                <div class="setup-muted">Dùng nội dung JSON credential từ Google Cloud, đổi sang TOML theo mẫu.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.code(str(SECRETS_EXAMPLE_PATH), language="text")
    with col3:
        st.markdown(
            """
            <div class="setup-card">
                <div class="setup-title">3. Chia sẻ Google Sheet</div>
                <div class="setup-muted">Share sheet cho email service account và điền đúng SPREADSHEET_URL.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Tải lại sau khi cấu hình", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    st.markdown("#### Mẫu `.streamlit/secrets.toml`")
    st.code(
        """[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/..."
""",
        language="toml",
    )

    st.markdown("#### Tạo cấu hình trực tiếp")
    with st.form("setup_secrets_form"):
        spreadsheet_url = st.text_input(
            "Google Sheet URL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
        service_account_json = st.text_area(
            "Dán toàn bộ nội dung file JSON service account",
            height=220,
            placeholder='{"type": "service_account", "project_id": "...", ...}',
        )
        submitted = st.form_submit_button("Lưu cấu hình và tải lại", use_container_width=True)

    if submitted:
        try:
            config = json.loads(service_account_json)
            if not spreadsheet_url.strip():
                st.error("Vui lòng nhập Google Sheet URL.")
                return
            config["SPREADSHEET_URL"] = spreadsheet_url.strip()
            write_streamlit_secrets(config)
            st.success("Đã tạo `.streamlit/secrets.toml`. Đang tải lại ứng dụng...")
            st.cache_data.clear()
            st.rerun()
        except json.JSONDecodeError:
            st.error("Nội dung service account không phải JSON hợp lệ.")
        except Exception as exc:
            st.error(f"Không thể lưu cấu hình: {exc}")


def toml_quote(value):
    text = "" if value is None else str(value)
    return json.dumps(text)


def write_streamlit_secrets(config):
    required_keys = [
        "type",
        "project_id",
        "private_key_id",
        "private_key",
        "client_email",
        "client_id",
        "auth_uri",
        "token_uri",
        "auth_provider_x509_cert_url",
        "client_x509_cert_url",
        "SPREADSHEET_URL",
    ]
    missing = [key for key in required_keys if not config.get(key)]
    if missing:
        raise ValueError("Thiếu trường: " + ", ".join(missing))

    SECRETS_PATH.parent.mkdir(parents=True, exist_ok=True)
    lines = ["[gcp_service_account]"]
    for key in required_keys:
        lines.append(f"{key} = {toml_quote(config[key])}")
    SECRETS_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")

# ---------- load data ----------
@st.cache_data(ttl=300)
def load_data():
    df_csdl, df_taichinh, _ = gsheets.load_dataframes()
    return df_csdl, df_taichinh

try:
    df_csdl, df_taichinh = load_data()
except Exception as e:
    render_connection_error(e)
    st.stop()

ma_tram_list = [str(v).strip().upper() for v in df_csdl["ma_tram"]]

with st.sidebar:
    st.markdown("### Trạng thái")
    st.success("Google Sheets đã kết nối")
    st.caption(f"CSDL: {len(df_csdl)} dòng")
    st.caption(f"Tài chính: {len(df_taichinh)} dòng")

# session
st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("images", {})
st.session_state.setdefault("images_bytes", {})
st.session_state.setdefault("image_upload_mode", "Upload ảnh ngay")

def bytes_from_pil(img: Image.Image):
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


DATE_FIELDS = {"ngaybatdau", "ngayketthuc", "ngayky", "tungay", "denngay"}
MONEY_FIELDS = {
    "tienthangtruocthue",
    "tienthueky",
    "tientruocthueky",
    "tongtienky",
}


def normalize_key(value):
    return str(value).lower().replace("_", "").replace(" ", "")


def format_vn_number(value, decimals=0):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        number = pd.to_numeric(text.replace(".", "").replace(",", "."), errors="raise")
    except Exception:
        return text
    if decimals == 0:
        return f"{int(round(float(number))):,}".replace(",", ".")
    formatted = f"{float(number):,.{decimals}f}"
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def format_vn_money(value):
    return format_vn_number(value, decimals=0)


def format_vn_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    if isinstance(value, (int, float)) and value > 25000:
        try:
            base = pd.to_datetime("1899-12-30")
            dt = base + pd.to_timedelta(int(value), "D")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return str(value)


def format_value_for_field(key, value):
    key_norm = normalize_key(key)
    if key_norm in MONEY_FIELDS:
        return format_vn_money(value)
    if key_norm in DATE_FIELDS:
        return format_vn_date(value)
    if key_norm == "kythanhtoan":
        return format_vn_number(value)
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)


def build_formatted_data(data):
    return {k: format_value_for_field(k, v) for k, v in data.items()}

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
    left, center, right = st.columns([1, 1.35, 1])
    with center:
        st.markdown(
            """
            <div class="friendly-panel">
                <div class="friendly-title">Đăng nhập thông tin trạm</div>
                <div class="friendly-muted">Chọn tháng thanh toán, nhập mã trạm và mật khẩu để bắt đầu tạo biên bản.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            ma_tram = st.text_input("Mã nhà trạm", placeholder="Ví dụ: AGG001").upper().strip()
            list_thang = sorted(df_taichinh["Thang"].astype(str).unique().tolist())
            thang = st.selectbox("Tháng thanh toán", [""] + list_thang)
            password = st.text_input("Mật khẩu", type="password", placeholder="Nhập mật khẩu")
            submit = st.form_submit_button("Đăng nhập", use_container_width=True)

    if submit:
        if not ma_tram:
            st.warning("Nhập mã trạm!")
            st.stop()
        if not thang:
            st.warning("Chọn tháng thanh toán!")
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
formatted_user_data = build_formatted_data(user_data)

overview_tab, image_tab, report_tab = st.tabs(["Thông tin trạm", "Hình ảnh nghiệm thu", "Tạo biên bản"])

with overview_tab:
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Mã trạm", formatted_user_data.get("ma_tram", ma_tram))
    col_b.metric("Tháng", formatted_user_data.get("Thang", thang))
    col_c.metric("Loại cột", formatted_user_data.get("Loai_cot", ""))
    col_d.metric("Tổng tiền kỳ", formatted_user_data.get("tongtienky", ""))

    display_fields = [
        "Ma_HD", "Dia_chi", "Ten_don_vi_XHH", "Chu_ha_tang", "Chuc_vu",
        "Loai_cot", "Loai_tram", "Phong_may", "Dieu_hoa",
        "Ky_thanh_toan", "tienthangtruocthue", "tientruocthueky",
        "tienthueky", "tongtienky", "tu_ngay", "den_ngay", "ngaybatdau", "ngayketthuc", "ngay_ky",
    ]
    rows = []
    for field in display_fields:
        if field in formatted_user_data:
            rows.append({"Trường dữ liệu": field, "Giá trị": formatted_user_data[field]})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# upload & rotate
IMAGE_RULES = [
    {"no": 1, "title": "BẢNG TÊN TRẠM", "conditions": []},
    {"no": 2, "title": "TỔNG THỂ CỘT ANTEN", "conditions": []},
    {"no": 3, "title": "Móng M0", "conditions": []},
    {"no": 4, "title": "Móng co 1", "conditions": ["guyed_tower"]},
    {"no": 5, "title": "Móng co 2", "conditions": ["guyed_tower"]},
    {"no": 6, "title": "Móng co 3", "conditions": ["guyed_tower"]},
    {"no": 7, "title": "Hình ảnh thể hiện lực căng chỉnh lực dây co", "conditions": ["guyed_tower"]},
    {"no": 8, "title": "Hình ảnh thể hiện đo lực siết khóa cáp", "conditions": ["guyed_tower"]},
    {"no": 9, "title": "Vị trí lắp anten và RRU", "conditions": []},
    {"no": 10, "title": "Tổng thể bên ngoài phòng máy", "conditions": ["rented_station"]},
    {"no": 11, "title": "Cốt cấp AC nhập trạm", "conditions": []},
    {"no": 12, "title": "Hình ảnh điều hòa", "conditions": ["aircon"]},
]


def normalize_text(value):
    import unicodedata

    text = "" if value is None else str(value)
    text = text.strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.replace("đ", "d")


def is_rented(value):
    text = normalize_text(value)
    return text not in {"", "nan", "none", "khong", "khong thue", "khong co"}


def get_required_image_rules(data):
    checks = {
        "guyed_tower": "day co" in normalize_text(data.get("Loai_cot", "")),
        "rented_station": is_rented(data.get("Loai_tram", "")) or is_rented(data.get("Phong_may", "")),
        "aircon": is_rented(data.get("Dieu_hoa", "")),
    }

    required = []
    for rule in IMAGE_RULES:
        if all(checks.get(cond, False) for cond in rule["conditions"]):
            required.append(rule)
    return required


def image_rule_status(rule, data):
    if not rule["conditions"]:
        return "Bắt buộc *", "Tất cả các loại"

    descriptions = {
        "guyed_tower": "Nếu là loại cột dây co",
        "rented_station": "Nếu có thuê trạm/phòng máy",
        "aircon": "Nếu có thuê điều hòa",
    }
    required_rules = get_required_image_rules(data)
    status = "Bắt buộc *" if rule in required_rules else "Không áp dụng"
    reason = ", ".join(descriptions.get(cond, cond) for cond in rule["conditions"])
    return status, reason


def save_uploaded_image(slot_no, uploaded_file):
    img = Image.open(uploaded_file).convert("RGB")
    img.thumbnail((1600, 1600))
    key = f"img{slot_no}"
    st.session_state.images[key] = img
    st.session_state.images_bytes[key] = bytes_from_pil(img)


required_image_rules = get_required_image_rules(user_data)
uploaded_required_count = sum(
    1 for rule in required_image_rules
    if f"img{rule['no']}" in st.session_state.images_bytes
)

with st.sidebar:
    st.markdown("---")
    st.markdown("### Phiên làm việc")
    st.caption(f"Mã trạm: **{ma_tram}**")
    st.caption(f"Tháng: **{thang}**")
    st.progress(
        uploaded_required_count / len(required_image_rules)
        if required_image_rules else 1.0
    )
    st.caption(f"Ảnh bắt buộc: {uploaded_required_count}/{len(required_image_rules)}")
    if st.button("Đăng xuất", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.images = {}
        st.session_state.images_bytes = {}
        st.session_state.image_upload_mode = "Upload ảnh ngay"
        st.rerun()

st.markdown(
    """
    <div class="quick-steps">
        <div class="quick-step"><strong>1. Kiểm tra thông tin</strong><span>Xem nhanh dữ liệu trạm và giá trị thanh toán.</span></div>
        <div class="quick-step"><strong>2. Chọn ảnh</strong><span>Upload ngay hoặc để bổ sung ảnh sau.</span></div>
        <div class="quick-step"><strong>3. Tải biên bản</strong><span>Tạo DOCX theo lựa chọn ảnh của bạn.</span></div>
    </div>
    """,
    unsafe_allow_html=True,
)

def do_rotate(idx, angle):
    key = f"img{idx}"
    if key in st.session_state.images:
        img = st.session_state.images[key]
        rotated = img.rotate(angle, expand=True)
        st.session_state.images[key] = rotated
        st.session_state.images_bytes[key] = bytes_from_pil(rotated)

def render_image_picker(rule, in_dialog=False):
    no = rule["no"]
    key = f"img{no}"
    uploaded = st.file_uploader(
        "Chọn ảnh mới để thay thế" if key in st.session_state.images else "Chọn ảnh",
        type=["jpg", "jpeg", "png"],
        key=f"u{no}_{'dialog' if in_dialog else 'inline'}",
    )
    if uploaded:
        save_uploaded_image(no, uploaded)
        st.success("Đã lưu ảnh.")

    if key in st.session_state.images:
        st.image(st.session_state.images[key], width=450)
        col_l, col_r = st.columns(2)
        with col_l:
            st.button("⟲", key=f"L{no}_{'dialog' if in_dialog else 'inline'}", on_click=do_rotate, args=(no, 90))
        with col_r:
            st.button("⟳", key=f"R{no}_{'dialog' if in_dialog else 'inline'}", on_click=do_rotate, args=(no, -90))


if hasattr(st, "dialog"):
    @st.dialog("Chọn hình ảnh")
    def image_dialog(rule):
        st.markdown(f"### Ảnh {rule['no']} - {rule['title']}")
        render_image_picker(rule, in_dialog=True)


with image_tab:
    st.subheader("Hình ảnh nghiệm thu")
    st.caption("Bạn có thể upload ảnh ngay hoặc tạo biên bản trước rồi bổ sung ảnh sau.")

    upload_mode = st.radio(
        "Cách xử lý hình ảnh",
        ["Upload ảnh ngay", "Để upload sau"],
        horizontal=True,
        key="image_upload_mode",
    )
    upload_later = upload_mode == "Để upload sau"

    progress_value = (
        uploaded_required_count / len(required_image_rules)
        if required_image_rules else 1.0
    )
    st.progress(progress_value)
    st.caption(f"Đã chọn {uploaded_required_count}/{len(required_image_rules)} ảnh bắt buộc.")
    if upload_later:
        st.info("Bạn đang chọn tạo biên bản trước. Ảnh chưa upload sẽ không chặn quá trình tạo file.")

    rule_rows = []
    for rule in IMAGE_RULES:
        status, reason = image_rule_status(rule, user_data)
        rule_rows.append({
            "Ảnh": f"Ảnh {rule['no']}",
            "Hạng mục": rule["title"],
            "Trạng thái": status,
            "Rule": reason,
        })
    with st.expander("Xem bảng rule ảnh", expanded=False):
        st.dataframe(pd.DataFrame(rule_rows), use_container_width=True, hide_index=True)

    if upload_later:
        missing_titles = [
            f"Ảnh {rule['no']} - {rule['title']}"
            for rule in required_image_rules
            if f"img{rule['no']}" not in st.session_state.images_bytes
        ]
        if missing_titles:
            with st.expander("Danh sách ảnh sẽ bổ sung sau", expanded=True):
                st.write(missing_titles)
    else:
        for rule in required_image_rules:
            i = rule["no"]
            key = f"img{i}"
            has_image = key in st.session_state.images
            status_text = "Đã chọn" if has_image else "Chưa chọn"
            label = f"{'✓' if has_image else '•'} Ảnh {i} - {rule['title']} * (bắt buộc) - {status_text}"
            with st.expander(label, expanded=not has_image):
                if has_image:
                    col1, col2, col3, col4 = st.columns([4, 1, 1, 1])
                    with col1:
                        st.image(st.session_state.images[key], width=450)
                    with col2:
                        st.button("⟲", key=f"L{i}", on_click=do_rotate, args=(i, 90), use_container_width=True)
                    with col3:
                        st.button("⟳", key=f"R{i}", on_click=do_rotate, args=(i, -90), use_container_width=True)
                    with col4:
                        if hasattr(st, "dialog"):
                            if st.button("Thay thế", key=f"open{i}", use_container_width=True):
                                image_dialog(rule)
                        else:
                            with st.expander("Thay thế ảnh"):
                                render_image_picker(rule)
                else:
                    st.markdown(
                        '<span class="required-badge">Bắt buộc</span>',
                        unsafe_allow_html=True,
                    )
                    if hasattr(st, "dialog"):
                        if st.button("Chọn ảnh", key=f"open{i}", use_container_width=True):
                            image_dialog(rule)
                    else:
                        render_image_picker(rule)

# ---------- CREATE REPORT ----------
# ---------- CREATE REPORT ----------
from modules import docx_image_safe as ds


with report_tab:
    st.subheader("Tạo biên bản")
    st.caption("Kiểm tra nhanh trước khi xuất file Word.")
    col_ready_1, col_ready_2, col_ready_3 = st.columns(3)
    col_ready_1.metric("Mã trạm", ma_tram)
    col_ready_2.metric("Tháng", thang)
    col_ready_3.metric("Ảnh bắt buộc", f"{uploaded_required_count}/{len(required_image_rules)}")
    upload_later = st.session_state.get("image_upload_mode") == "Để upload sau"
    if uploaded_required_count == len(required_image_rules):
        st.success("Đã đủ ảnh bắt buộc. Có thể tạo biên bản.")
    elif upload_later:
        st.info("Bạn đã chọn upload ảnh sau. Có thể tạo biên bản trước.")
    else:
        st.warning("Chưa đủ ảnh bắt buộc. Vui lòng hoàn tất ở tab Hình ảnh nghiệm thu.")

if report_tab.button("📄 Tạo & Tải biên bản", use_container_width=True):
    try:
        missing_images = [
            f"Ảnh {rule['no']} - {rule['title']} * (bắt buộc)"
            for rule in required_image_rules
            if f"img{rule['no']}" not in st.session_state.images_bytes
        ]
        upload_later = st.session_state.get("image_upload_mode") == "Để upload sau"
        if missing_images and not upload_later:
            st.error("Vui lòng chọn đủ ảnh trước khi tạo biên bản:")
            st.write(missing_images)
            st.stop()
        if missing_images and upload_later:
            st.warning("Biên bản sẽ được tạo trước, các ảnh sau đây cần bổ sung sau:")
            st.write(missing_images)

        with st.spinner("Đang tạo biên bản..."):

            # Load template docx
            with open("template.docx", "rb") as f:
                docx_bytes = f.read()

            doc = ds.load_docx_bytes(docx_bytes)

            # Chuẩn hoá dữ liệu user_data
            normalized_map = {}

            for k, v in user_data.items():
                key = normalize_key(k)
                normalized_map[key] = format_value_for_field(k, v)

            normalized_map["tungay"] = normalized_map.get("ngaybatdau", "")
            normalized_map["denngay"] = normalized_map.get("ngayketthuc", "")


            # Danh sách placeholders trong template:
            holders = [
                # text placeholders
                "ngaybatdau", "ngayketthuc", "ngay_ky", "tu_ngay", "den_ngay",
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
            for rule in required_image_rules:
                i = rule["no"]
                key = f"img{i}"
                if key not in st.session_state.images_bytes:
                    continue
                img_bytes = st.session_state.images_bytes[key]

                inserted = False
                for ph in [
                    f"${{Anh{i}}}",
                    f"$Anh{i}",
                    f"${{anh{i}}}",
                    f"$anh{i}",
                    f"Ảnh {i}",
                    f"ảnh {i}",
                    f"Anh {i}",
                    f"anh {i}",
                ]:
                    inserted = ds.insert_image(doc, ph, img_bytes, 12) or inserted

                if not inserted:
                    ds.insert_image_in_final_table(doc, i, rule["title"], img_bytes, 12)

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
