import zipfile
from io import BytesIO
from docx import Document
from docx.shared import Inches
import xml.sax.saxutils as saxutils
from PIL import Image


# =============================
# Escape XML để tránh lỗi format
# =============================
def _xml_escape(s: str) -> str:
    return saxutils.escape(s)


# ======================================================
# 1) REPLACE TEXT (hỗ trợ split-runs)
# ======================================================
def replace_text_bytes(docx_bytes: bytes, placeholder: str, replacement: str) -> bytes:
    # ---------- Phase 1: Replace trực tiếp trong XML ----------
    bio_in = BytesIO(docx_bytes)
    bio_out = BytesIO()

    with zipfile.ZipFile(bio_in, "r") as zin:
        try:
            xml_doc = zin.read("word/document.xml").decode("utf-8")
        except KeyError:
            xml_doc = None

        if xml_doc and placeholder in xml_doc:
            new_xml = xml_doc.replace(placeholder, _xml_escape(str(replacement)))
            with zipfile.ZipFile(bio_out, "w") as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, new_xml.encode("utf-8"))
                    else:
                        zout.writestr(item, zin.read(item.filename))
            return bio_out.getvalue()

    # ---------- Phase 2: python-docx fallback ----------
    doc = Document(BytesIO(docx_bytes))

    def replace_in_paragraph(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder not in full_text:
            return False

        new_text = full_text.replace(placeholder, str(replacement))

        for r in paragraph.runs:
            r.text = ""

        paragraph.add_run(new_text)
        return True

    changed = False

    # Replace trong paragraphs
    for p in doc.paragraphs:
        if replace_in_paragraph(p):
            changed = True

    # Replace trong tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_in_paragraph(p):
                        changed = True

    # Save nếu có thay đổi
    if changed:
        out = BytesIO()
        doc.save(out)
        return out.getvalue()

    return docx_bytes


# ======================================================
# 2) XOAY ẢNH TRƯỚC KHI CHÈN
# ======================================================
def rotate_image_bytes(image_bytes: bytes, angle: int):
    """
    Trả về ảnh đã xoay (bytes)
    """
    img = Image.open(BytesIO(image_bytes))
    rotated = img.rotate(angle, expand=True)

    out = BytesIO()
    rotated.save(out, format="JPEG")
    return out.getvalue()


# ======================================================
# 3) CHÈN ẢNH VÀO DOCX TỪ BYTES
# ======================================================
def insert_image_into_docx_bytes(docx_bytes, placeholder, image_bytes):
    doc = Document(BytesIO(docx_bytes))

    def insert_in_paragraph(paragraph):
        if placeholder in paragraph.text:

            # Xóa placeholder
            for run in paragraph.runs:
                run.text = run.text.replace(placeholder, "")

            img_stream = BytesIO(image_bytes)
            try:
                paragraph.add_run().add_picture(img_stream, width=Inches(3.5))
            except Exception:
                paragraph.add_run().add_picture(img_stream)

            return True

        return False

    replaced = False

    # Insert vào paragraphs
    for p in doc.paragraphs:
        if insert_in_paragraph(p):
            replaced = True

    # Insert vào tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if insert_in_paragraph(p):
                        replaced = True

    out = BytesIO()
    doc.save(out)
    return out.getvalue()
