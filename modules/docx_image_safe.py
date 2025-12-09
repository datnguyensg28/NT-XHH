import io
from docx import Document
from docx.shared import Cm


def load_docx_bytes(docx_bytes: bytes):
    bio = io.BytesIO(docx_bytes)
    return Document(bio)


def save_docx(doc: Document) -> bytes:
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def _replace_in_paragraph(paragraph, placeholder, value):
    full_text = "".join(run.text for run in paragraph.runs)

    if placeholder not in full_text:
        return False

    # GIỮ FORMAT – không xóa paragraph
    new_text = full_text.replace(placeholder, value)

    # Xóa toàn bộ run cũ
    for r in list(paragraph.runs):
        r.clear()

    # Thêm lại 1 run DUY NHẤT giữ format của run đầu tiên
    run = paragraph.add_run(new_text)
    return True


def replace_text(doc: Document, placeholder: str, value: str):
    replaced = False

    # Check paragraph
    for p in doc.paragraphs:
        if _replace_in_paragraph(p, placeholder, value):
            replaced = True

    # Check table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _replace_in_paragraph(p, placeholder, value):
                        replaced = True

    return replaced


def _insert_img_to_paragraph(paragraph, placeholder, img_bytes, width_cm):
    full_text = "".join(run.text for run in paragraph.runs)

    if placeholder not in full_text:
        return False

    # Remove all runs
    for r in list(paragraph.runs):
        r.clear()

    run = paragraph.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    return True


def insert_image(doc: Document, placeholder: str, img_bytes: bytes, width_cm=12):
    inserted = False

    for p in doc.paragraphs:
        if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
            inserted = True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
                        inserted = True

    return inserted
