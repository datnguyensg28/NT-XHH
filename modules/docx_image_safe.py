# modules/docx_image_safe.py
import io
from docx import Document
from docx.shared import Cm


def load_docx_bytes(docx_bytes: bytes):
    """Load DOCX từ bytes."""
    bio = io.BytesIO(docx_bytes)
    return Document(bio)


def save_docx(doc: Document) -> bytes:
    """Save DOCX về bytes."""
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def _replace_in_paragraph(paragraph, placeholder, value):
    """
    Replace text nhưng GIỮ NGUYÊN định dạng run.
    Không xóa run – chỉ chỉnh nội dung trong run.
    """
    # full text
    full = "".join(r.text for r in paragraph.runs)
    if placeholder not in full:
        return False

    new_full = full.replace(placeholder, value)

    # Giữ nguyên style → chỉ update text
    idx = 0
    for r in paragraph.runs:
        run_len = len(r.text)
        r.text = new_full[idx: idx + run_len]
        idx += run_len

    return True


def replace_text(doc: Document, placeholder: str, value: str):
    """Replace text trong toàn bộ DOCX (paragraph + bảng)."""
    replaced = False

    # paragraph
    for p in doc.paragraphs:
        if _replace_in_paragraph(p, placeholder, value):
            replaced = True

    # trong table
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _replace_in_paragraph(p, placeholder, value):
                        replaced = True

    return replaced


def _insert_img_to_paragraph(paragraph, placeholder, img_bytes, width_cm):
    full = ''.join(run.text for run in paragraph.runs)
    if placeholder not in full:
        return False

    # xoá run
    for r in list(paragraph.runs):
        r.clear()

    run = paragraph.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    return True


def insert_image(doc: Document, placeholder: str, img_bytes: bytes, width_cm=12):
    """Chèn ảnh vào vị trí placeholder."""
    inserted = False

    # paragraph
    for p in doc.paragraphs:
        if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
            inserted = True

    # trong table
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _insert_img_to_paragraph(p, placeholder, img_bytes, width_cm):
                        inserted = True

    return inserted
