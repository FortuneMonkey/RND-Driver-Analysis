import fitz  # PyMuPDF


def render_pdf_page(pdf_path, out_png_path, page_number=0, dpi=150):

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