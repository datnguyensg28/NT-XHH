import io
import zipfile
import re
from docx import Document
from docx.shared import Cm


def merge_xml_text(xml: str) -> str:
    """
    Gom toàn bộ text node bị split trong Word:
    <w:t>A</w:t><w:t>B</w:t> => AB
    """
    xml = re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)
    return xml


def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
    """
    Replace text placeholder trong file DOCX mà vẫn giữ nguyên cấu trúc ZIP.
    """
    bio = io.BytesIO(docx_bytes)

    # --------- ĐỌC TOÀN BỘ ZIP TRƯỚC WHEN zin còn mở ---------
    with zipfile.ZipFile(bio, "r") as zin:

        # Lấy nội dung file document.xml
        xml = zin.read("word/document.xml").decode("utf-8")
        xml = merge_xml_text(xml)
        xml = xml.replace(placeholder, value)

        # Load tất cả file khác vào bộ nhớ
        files = {
            item.filename: zin.read(item.filename)
            for item in zin.infolist()
        }

    # --------- GHI ZIP MỚI ---------
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:

        for filename, content in files.items():
            if filename == "word/document.xml":
                zout.writestr(filename, xml.encode("utf-8"))
            else:
                zout.writestr(filename, content)

    return out.getvalue()


def insert_image_into_docx_bytes(docx_bytes: bytes, placeholder: str, img_bytes: bytes, width_cm=10):
    """
    Chèn hình vào DOCX tại vị trí ${AnhX}.
    Dùng python-docx để xử lý ảnh chuẩn hơn.
    """
    bio = io.BytesIO(docx_bytes)
    doc = Document(bio)

    for p in doc.paragraphs:
        if placeholder in p.text:
            p.clear()
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
