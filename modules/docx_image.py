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
    bio = io.BytesIO(docx_bytes)

    with zipfile.ZipFile(bio, "r") as zin:
        files = {f.filename: zin.read(f.filename) for f in zin.infolist()}

    # ✅ MERGE XML TRƯỚC
    xml = files["word/document.xml"].decode("utf-8")
    xml = _merge_xml(xml)

    # ✅ replace sau khi merge
    xml = xml.replace(placeholder, value)

    files["word/document.xml"] = xml.encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    return out.getvalue()



ddef insert_image_into_docx_bytes(docx_bytes, placeholder, img_bytes, width_cm=12):
    import io
    from docx import Document
    from docx.shared import Cm

    bio = io.BytesIO(docx_bytes)
    doc = Document(bio)

    ph_norm = placeholder.replace(" ", "").lower()

    for p in doc.paragraphs:
        full_text = "".join(r.text for r in p.runs).replace(" ", "").lower()

        if ph_norm in full_text:
            # ❌ XÓA SẠCH RUN
            for r in p.runs:
                r.text = ""

            # ✅ CHÈN ẢNH
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
