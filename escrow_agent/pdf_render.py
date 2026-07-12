"""Render PDF pages to JPEG images using PyMuPDF (pip-installable, no system
poppler/tesseract dependency — works on Render's plain Python runtime and
handles scanned + native-text PDFs identically, since Claude reads the images
directly rather than needing an OCR text layer)."""
import base64
import fitz  # PyMuPDF


def pdf_to_images(path_or_bytes, dpi=110, max_pages=None, jpeg_quality=70):
    """Returns a list of dicts: {"page": n, "media_type": "image/jpeg", "base64": ...}"""
    if isinstance(path_or_bytes, (bytes, bytearray)):
        doc = fitz.open(stream=path_or_bytes, filetype="pdf")
    else:
        doc = fitz.open(path_or_bytes)
    n_pages = doc.page_count
    pages = range(min(n_pages, max_pages)) if max_pages else range(n_pages)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    images = []
    for i in pages:
        pix = doc[i].get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        data = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
        images.append({"page": i + 1, "media_type": "image/jpeg", "base64": base64.b64encode(data).decode()})
    doc.close()
    return images, n_pages
