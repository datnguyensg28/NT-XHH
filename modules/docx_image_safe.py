import io
from docx import Document
from docx.shared import Cm

def load_docx_bytes(docx_bytes: bytes):
    return Document(io.BytesIO(docx_bytes))

def save_docx(doc: Document) -> bytes:
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def _replace_in_paragraph(paragraph, placeholder, value):
    # Lấy toàn bộ text gốc
    runs = paragraph.runs
    if not runs:
        return False

    full_text = "".join(r.text for r in runs)
    if placeholder not in full_text:
        return False

    # tạo text mới sau khi replace
    new_text = full_text.replace(placeholder, value)

    # phân bổ text mới vào các run theo chiều dài run cũ
    original_lens = [len(r.text) for r in runs]
    pos = 0

    for i, r in enumerate(runs):
        take = original_lens[i]

        # run cuối lấy phần còn lại
        if i == len(runs) - 1:
            r.text = new_text[pos:]
            break

        r.text = new_text[pos:pos + take]
        pos += take

    return True


def replace_text(doc: Document, placeholder: str, value: str):
    replaced = False

    # paragraphs
    for p in doc.paragraphs:
        if _replace_in_paragraph(p, placeholder, value):
            replaced = True

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _replace_in_paragraph(p, placeholder, value):
                        replaced = True

    return replaced


def _insert_img_to_paragraph(paragraph, placeholder, img_bytes, width_cm):
    runs = paragraph.runs
    full_text = "".join(r.text for r in runs)

    if placeholder not in full_text:
        return False

    # clear all runs → chèn ảnh
    for r in runs:
        r.text = ""

    run = paragraph.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    return True


def insert_image(doc: Document, placeholder: str, img_bytes: bytes, width_cm=12):
    inserted = False

    # paragraphs
    for p in doc.paragraphs:
        if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
            inserted = True

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
                        inserted = True

    return inserted
