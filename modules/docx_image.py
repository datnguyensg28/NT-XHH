import io
import zipfile
import re
import uuid
from docx import Document
from docx.shared import Cm
#verify that the required libraries are installed
def _merge_xml(xml: str) -> str:
    xml = re.sub(r"</w:t>\s*<w:t[^>]*>", "", xml)
    xml = re.sub(r"</w:t><w:t[^>]*>", "", xml)
    return xml

def replace_text_bytes(docx_bytes: bytes, placeholder: str, value: str) -> bytes:
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
    # 1) merge XML
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

    # 2) try replace directly with python-docx
    doc = Document(merged)
    norm_placeholder = placeholder.lower().replace(" ", "")

    for p in doc.paragraphs:
        txt = "".join(r.text for r in p.runs).lower().replace(" ", "")
        if norm_placeholder in txt:
            for r in list(p.runs):
                try:
                    r._element.getparent().remove(r._element)
                except:
                    pass
            r2 = p.add_run()
            r2.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

            out = io.BytesIO()
            doc.save(out)
            return out.getvalue()

    # 3) fallback → token an toàn
    token = f"IMGTOKEN_{uuid.uuid4().hex}"

    xml = files["word/document.xml"].decode("utf-8")
    variants = {
        placeholder,
        placeholder.replace("${", "$").replace("}", ""),
        placeholder.replace("$", "${") + "}",
    }

    replaced = False
    for pat in variants:
        if pat in xml:
            xml = xml.replace(pat, token)
            replaced = True

    if not replaced:
        return docx_bytes

    files["word/document.xml"] = xml.encode("utf-8")

    temp = io.BytesIO()
    with zipfile.ZipFile(temp, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)
    temp.seek(0)

    # read again
    doc2 = Document(temp)
    for p in doc2.paragraphs:
        if token in p.text:
            for r in list(p.runs):
                try:
                    r._element.getparent().remove(r._element)
                except:
                    pass
            r2 = p.add_run()
            r2.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    out2 = io.BytesIO()
    doc2.save(out2)
    return out2.getvalue()
