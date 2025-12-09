# modules/docx_image_safe.py
import io
from docx import Document
from docx.shared import Cm

# ============================
# Load & Save DOCX
# ============================

def load_docx_bytes(docx_bytes: bytes):
    """Load docx từ bytes."""
    bio = io.BytesIO(docx_bytes)
    return Document(bio)

def save_docx(doc: Document) -> bytes:
    """Save docx về bytes."""
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ============================
# Replace TEXT (Safe Mode)
# ============================

def replace_text_in_paragraph(paragraph, placeholder, value):
    """Replace placeholder trong 1 paragraph bất chấp Word split runs."""
    full = ''.join(run.text for run in paragraph.runs)

    if placeholder not in full:
        return False

    # replace toàn bộ
    new_text = full.replace(placeholder, value)

    # xoá run cũ
    for run in list(paragraph.runs):
        run.clear()

    paragraph.add_run(new_text)
    return True


def replace_text_in_doc(doc, placeholder, value):
    """Replace toàn doc (paragraph + table)."""
    replaced = False

    # paragraphs
    for p in doc.paragraphs:
        if replace_text_in_paragraph(p, placeholder, value):
            replaced = True

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_text_in_paragraph(p, placeholder, value):
                        replaced = True

    return replaced


# ============================
# Insert IMAGE (Safe Mode)
# ============================

def insert_image(doc, placeholder, img_bytes, width_cm=10):
    """Chèn hình tại vị trí placeholder."""
    found = False

    def process_para(p):
        nonlocal found
        full = ''.join(run.text for run in p.runs)
        if placeholder in full:
            for run in list(p.runs):
                run.clear()

            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
            found = True

    # paragraphs
    for p in doc.paragraphs:
        process_para(p)

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_para(p)

    return found
