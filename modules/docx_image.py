import io
import zipfile
import re
from docx import Document
from docx.shared import Cm


def merge_xml_text(xml: str) -> str:
    """
    Gom toàn bộ text node bị split trong Word:
    <w:t>ABC</w:t><w:t>DEF</w:t> => ABCDEF
    """
    xml = re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)
    return xml


def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
    """
    Replace text placeholder trong file DOCX.
    Không dùng python-docx vì nó khó xử lý XML split.
    """
    bio = io.BytesIO(docx_bytes)

    with zipfile.ZipFile(bio, "r") as zin:
        xml = zin.read("word/document.xml").decode("utf-8")

    xml = merge_xml_text(xml)
    xml = xml.replace(placeholder, value)

    # Ghi lại DOCX mới
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr("word/document.xml", xml.encode("utf-8"))
            else:
                zout.writestr(item.filename, zin.read(item.filename))

    return out.getvalue()


def insert_image_into_docx_bytes(docx_bytes: bytes, placeholder: str, img_bytes: bytes, width_cm=10):
    """
    Chèn hình vào DOCX tại vị trí ${AnhX}.
    Dùng python-docx vì xử lý ảnh chuẩn hơn.
    """
    bio = io.BytesIO(docx_bytes)
    doc = Document(bio)

    for p in doc.paragraphs:
        if placeholder in p.text:
            p.clear()  # xoá placeholder
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
