"""
screenshot.py
-------------
Renders a page of a PDF report to a PNG, so it can be embedded inline in an
email body. Requires: pip install pymupdf
"""
import fitz  # PyMuPDF


def render_pdf_page(pdf_path, out_png_path, page_number=0, dpi=150):
    """Renders `page_number` (0-indexed) of pdf_path to out_png_path.

    Uses pix.tobytes()+open() rather than pix.save() -- MuPDF's own file
    writer (fz_save_pixmap_as_png) can fail with a permission error on some
    network-share paths, while plain Python file I/O (used everywhere else
    in this pipeline) writes to the same paths without issue.
    """
    doc = fitz.open(pdf_path)
    try:
        if page_number >= doc.page_count:
            raise ValueError(f"{pdf_path} only has {doc.page_count} page(s).")
        pix = doc[page_number].get_pixmap(dpi=dpi)
        png_bytes = pix.tobytes("png")
    finally:
        doc.close()

    with open(out_png_path, "wb") as f:
        f.write(png_bytes)
    return out_png_path