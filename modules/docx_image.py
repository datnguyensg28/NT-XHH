import io
import zipfile
import re
from docx import Document
from docx.shared import Cm


def _merge_xml(xml: str) -> str:
    """
    Merge các text node bị split trong DOCX.
    Hỗ trợ cả:
        </w:t><w:t>
    và:
        </w:t></w:r><w:r><w:t>
    """

    # Merge kiểu phổ biến nhất
    xml = re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)

    # Merge trường hợp có thêm w:r
    xml = re.sub(r"</w:t>\s*</w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>", "", xml)

    return xml


def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
    """
    Replace text trong DOCX
    """
    bio = io.BytesIO(docx_bytes)

    # đọc toàn bộ file zip
    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    # xử lý document.xml
    xml = files["word/document.xml"].decode("utf-8")

    # merge text
    xml = _merge_xml(xml)

    # replace
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
    """
    # --- 1) Đọc docx và merge XML ---
    bio = io.BytesIO(docx_bytes)

    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)
    files["word/document.xml"] = xml.encode("utf-8")

    # --- 2) Ghi lại DOCX đã merge ---
    merged = io.BytesIO()
    with zipfile.ZipFile(merged, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    merged.seek(0)
    doc = Document(merged)

    # --- 3) Tìm placeholder & chèn ảnh ---
    for p in doc.paragraphs:

        # python-docx thêm dấu cách giữa các run → remove space
        clean_p = p.text.replace(" ", "")

        # so sánh không space để tìm chính xác
        if placeholder.replace(" ", "") in clean_p:

            # XÓA sạch các run cũ
            for r in list(p.runs):
                try:
                    r._element.getparent().remove(r._element)
                except:
                    pass

            # CHÈN ẢNH
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    # --- 4) Lưu kết quả ---
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
