# modules/docx_image.py
import io
import zipfile
import re
from docx import Document
from docx.shared import Cm

def _merge_xml(xml: str) -> str:
    """
    Merge các text node bị split trong DOCX:
    </w:t><w:t ...>  -> gộp lại (loại bỏ việc split run)
    """
    return re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)

def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
    """
    Replace text trong DOCX (làm việc trực tiếp trên bytes ZIP để tránh lỗi 'zip closed').
    placeholder: chính xác chuỗi cần thay (ví dụ "$ma_tram" hoặc "${ma_tram}" hoặc "$ma_tram;")
    value: chuỗi thay thế
    Trả về docx bytes mới.
    """
    bio = io.BytesIO(docx_bytes)

    # đọc toàn bộ file zip
    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    # xử lý document.xml
    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)
    # thay thế mọi occurrence
    xml = xml.replace(placeholder, value)

    # ghi ZIP mới
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for name, content in files.items():
            if name == "word/document.xml":
                zout.writestr(name, xml.encode("utf-8"))
            else:
                zout.writestr(name, content)

    return out.getvalue()


def insert_image_into_docx_bytes(docx_bytes: bytes, placeholder: str, img_bytes: bytes, width_cm: float = 10):
    """
    Chèn hình vào đúng đoạn chứa placeholder (đã merge XML trước).
    placeholder: chính xác chuỗi cần tìm (ví dụ "${Anh1}" hoặc "$Anh1").
    img_bytes: bytes của ảnh (JPEG/PNG).
    width_cm: chiều rộng ảnh (cm).
    Trả về docx bytes mới.
    """
    # --- 1) Đầu tiên MERGE XML trong bản docx_bytes ---
    bio = io.BytesIO(docx_bytes)
    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)
    files["word/document.xml"] = xml.encode("utf-8")

    # --- 2) Ghi lại DOCX đã merge (rồi mới dùng python-docx đọc) ---
    merged = io.BytesIO()
    with zipfile.ZipFile(merged, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    # --- 3) Giờ python-docx có thể tìm chính xác placeholder ---
    merged.seek(0)
    doc = Document(merged)

    # normalize placeholder for checking: remove spaces and lowercase to be tolerant
    ph_norm = placeholder.replace(" ", "").lower()

    for p in doc.paragraphs:
        # python-docx ghép run bằng dấu cách khi hiển thị -> loại bỏ spaces để so sánh
        clean_p = p.text.replace(" ", "").lower()

        if ph_norm in clean_p:
            # XÓA sạch mọi run chứa text trong đoạn paragraph
            for r in list(p.runs):
                try:
                    r._element.getparent().remove(r._element)
                except Exception:
                    # nếu có lỗi thì vẫn tiếp tục
                    pass

            # THÊM ẢNH
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    # --- 4) Lưu kết quả ---
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
