import io
import zipfile
import re
from docx import Document
from docx.shared import Cm

def _merge_xml(xml: str) -> str:
    """
    Hợp nhất các đoạn text bị Word tách nhỏ trong document.xml
    """
    # Kiểu phổ biến: </w:t><w:t>
    xml = re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)

    # Kiểu bị tách run: </w:t></w:r><w:r><w:t>
    xml = re.sub(r"</w:t>\s*</w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>", "", xml)

    return xml


def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
    """
    Thay text đơn giản trong document.xml
    """
    bio = io.BytesIO(docx_bytes)

    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)
    xml = xml.replace(placeholder, value)

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
    Tìm đoạn chứa ${AnhX} hoặc $AnhX và chèn ảnh vào đúng vị trí.
    Không quan trọng placeholder bị Word tách run.
    """
    bio = io.BytesIO(docx_bytes)

    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)
    files["word/document.xml"] = xml.encode("utf-8")

    merged = io.BytesIO()
    with zipfile.ZipFile(merged, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    merged.seek(0)
    doc = Document(merged)

    norm_target = placeholder.replace(" ", "")

    for p in doc.paragraphs:
        clean_p = p.text.replace(" ", "")
        if norm_target in clean_p:

            # Xóa toàn bộ run cũ trong đoạn
            for r in list(p.runs):
                try:
                    r._element.getparent().remove(r._element)
                except:
                    pass

            # Thêm ảnh
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
