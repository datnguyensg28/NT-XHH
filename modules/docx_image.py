# modules/docx_image.py
from docx import Document
from docx.shared import Inches
import io
from PIL import Image
import tempfile

def insert_image_into_docx_bytes(docx_bytes: bytes, placeholder: str, image_bytes: bytes, width_in_inches=3):
    """
    Open docx from bytes, replace placeholder with image, return new docx bytes.
    """
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.write(docx_bytes)
    tmp.flush()
    doc = Document(tmp.name)

    # Save image bytes to temp file
    imgtmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
    imgtmp.write(image_bytes)
    imgtmp.flush()
    inserted = False

    # paragraphs
    for p in doc.paragraphs:
        if placeholder in p.text:
            p.text = p.text.replace(placeholder, "")
            run = p.add_run()
            run.add_picture(imgtmp.name, width=Inches(width_in_inches))
            inserted = True
            break

    if not inserted:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, "")
                        run = cell.paragraphs[0].add_run()
                        run.add_picture(imgtmp.name, width=Inches(width_in_inches))
                        inserted = True
                        break
                if inserted:
                    break
            if inserted:
                break

    outtmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(outtmp.name)
    with open(outtmp.name, "rb") as f:
        result = f.read()
    return result
