import io
from docx import Document
from docx.shared import Cm


# ============================================================
#   LOAD & SAVE DOCX
# ============================================================

def load_docx_bytes(docx_bytes: bytes):
    bio = io.BytesIO(docx_bytes)
    return Document(bio)


def save_docx(doc: Document) -> bytes:
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ============================================================
#   MERGE RUNS (GIỮ NGUYÊN FORMAT)
# ============================================================

def _runs_have_same_format(r1, r2):
    """So sánh định dạng 2 run (rPr)."""
    p1 = r1._element.rPr
    p2 = r2._element.rPr
    return (p1.xml if p1 is not None else "") == (p2.xml if p2 is not None else "")


def _merge_runs(paragraph):
    """
    Gộp các run liên tiếp có chung định dạng.
    Sau bước này, placeholder luôn nằm trong MỘT run duy nhất.
    """
    runs = paragraph.runs
    if not runs:
        return []

    merged = []
    buffer_text = runs[0].text
    buffer_run = runs[0]

    for r in runs[1:]:
        if _runs_have_same_format(buffer_run, r):
            buffer_text += r.text
        else:
            nr = paragraph.add_run(buffer_text)
            # copy format
            if buffer_run._element.rPr is not None:
                nr._element.rPr = buffer_run._element.rPr
            merged.append(nr)

            buffer_run = r
            buffer_text = r.text

    # flush cuối
    nr = paragraph.add_run(buffer_text)
    if buffer_run._element.rPr is not None:
        nr._element.rPr = buffer_run._element.rPr
    merged.append(nr)

    # xóa run cũ
    for old in list(paragraph.runs):
        old._element.getparent().remove(old._element)

    return merged


# ============================================================
#   TEXT REPLACEMENT HOÀN HẢO GIỮ FORMAT
# ============================================================

def _replace_in_paragraph(paragraph, placeholder, value):
    """
    Replace nhưng giữ nguyên toàn bộ định dạng của đoạn chứa placeholder.
    """
    runs = _merge_runs(paragraph)

    replaced = False
    for r in runs:
        if placeholder in r.text:
            r.text = r.text.replace(placeholder, value)
            replaced = True

    return replaced


def replace_text(doc: Document, placeholder: str, value: str):
    """
    Replace trong toàn bộ tài liệu (paragraph + table), GIỮ ĐỊNH DẠNG.
    """
    replaced = False

    # paragraph
    for p in doc.paragraphs:
        if _replace_in_paragraph(p, placeholder, value):
            replaced = True

    # table
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _replace_in_paragraph(p, placeholder, value):
                        replaced = True

    return replaced


# ============================================================
#   IMAGE INSERTION (KHÔNG MẤT FORMAT)
# ============================================================

def _insert_img_to_paragraph(paragraph, placeholder, img_bytes, width_cm):
    runs = _merge_runs(paragraph)
    found = False

    for r in runs:
        if placeholder in r.text:
            found = True
            r.text = r.text.replace(placeholder, "")
            run_img = paragraph.add_run()
            run_img.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))

    return found


def insert_image(doc: Document, placeholder: str, img_bytes: bytes, width_cm=12):
    inserted = False

    # paragraph
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
