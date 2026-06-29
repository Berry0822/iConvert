"""
iConvert conversion engine (no GUI dependencies).
Each converter takes absolute source + destination paths.
Platform/heavy libraries are imported lazily so this module always imports.
"""

import os
import tempfile
import shutil


# Target label -> (output extension, set of valid source extensions)
TARGETS = {
    "PDF":                ("pdf",  {"doc", "docx", "ppt", "pptx", "jpg", "jpeg", "png"}),
    "Word (.docx)":       ("docx", {"pdf"}),
    "PowerPoint (.pptx)": ("pptx", {"pdf"}),
    "PNG":                ("png",  {"jpg", "jpeg"}),
    "JPG":                ("jpg",  {"png"}),
}
TARGET_ORDER = ["PDF", "Word (.docx)", "PowerPoint (.pptx)", "PNG", "JPG"]
SUPPORTED_INPUTS = {"pdf", "doc", "docx", "ppt", "pptx", "jpg", "jpeg", "png"}


def unique_path(path):
    """If 'path' exists, append ' (1)', ' (2)', ... so we never overwrite."""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 1
    while True:
        candidate = "{} ({}){}".format(root, i, ext)
        if not os.path.exists(candidate):
            return candidate
        i += 1


# --- image conversions (Pillow) --------------------------------------------
def convert_image(src, dst, target_ext):
    from PIL import Image
    img = Image.open(src)
    target_ext = target_ext.lower()
    if target_ext in ("jpg", "jpeg"):
        if img.mode in ("RGBA", "LA", "P"):           # JPEG has no transparency
            img = img.convert("RGBA")
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        else:
            img = img.convert("RGB")
        img.save(dst, "JPEG", quality=95)
    elif target_ext == "png":
        img.save(dst, "PNG")
    else:
        raise ValueError("Unsupported image target: " + target_ext)


def image_to_pdf(src, dst):
    from PIL import Image
    Image.open(src).convert("RGB").save(dst, "PDF", resolution=150.0)


# --- PDF -> Word (pdf2docx) -------------------------------------------------
def pdf_to_word(src, dst):
    from pdf2docx import Converter
    cv = Converter(src)
    try:
        cv.convert(dst, start=0, end=None)
    finally:
        cv.close()


# --- PDF -> PowerPoint (PyMuPDF render + python-pptx) -----------------------
def pdf_to_ppt(src, dst):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        import pymupdf as fitz
    from pptx import Presentation
    from pptx.util import Emu

    doc = fitz.open(src)
    tmpdir = tempfile.mkdtemp(prefix="iconvert_")
    try:
        prs = Presentation()
        first = doc[0].rect                            # points -> EMU
        prs.slide_width = Emu(int(first.width * 12700))
        prs.slide_height = Emu(int(first.height * 12700))
        blank = prs.slide_layouts[6]
        zoom = 2.0

        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            png = os.path.join(tmpdir, "page_{}.png".format(i))
            pix.save(png)
            slide = prs.slides.add_slide(blank)
            sw, sh = int(prs.slide_width), int(prs.slide_height)
            iw, ih = pix.width, pix.height
            if iw / ih > sw / sh:                       # aspect-fit, centered
                w = sw; h = int(sw * ih / iw); left = 0; top = int((sh - h) / 2)
            else:
                h = sh; w = int(sh * iw / ih); top = 0; left = int((sw - w) / 2)
            slide.shapes.add_picture(png, left, top, width=w, height=h)
        prs.save(dst)
    finally:
        doc.close()
        shutil.rmtree(tmpdir, ignore_errors=True)


# --- Word -> PDF (Microsoft Word via COM) ----------------------------------
def word_to_pdf(src, dst):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(src, ReadOnly=True)
        doc.SaveAs(dst, FileFormat=17)                 # 17 = wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()


# --- PowerPoint -> PDF (Microsoft PowerPoint via COM) ----------------------
def ppt_to_pdf(src, dst):
    import win32com.client
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        deck = ppt.Presentations.Open(src, WithWindow=False)
        deck.SaveAs(dst, 32)                           # 32 = ppSaveAsPDF
        deck.Close()
    finally:
        ppt.Quit()


def route(src_ext, out_ext):
    """Return the conversion function for a (source, target) pair, or None."""
    src_ext = src_ext.lower()
    out_ext = out_ext.lower()
    if out_ext == "pdf":
        if src_ext in ("doc", "docx"):
            return word_to_pdf
        if src_ext in ("ppt", "pptx"):
            return ppt_to_pdf
        if src_ext in ("jpg", "jpeg", "png"):
            return image_to_pdf
    elif out_ext == "docx" and src_ext == "pdf":
        return pdf_to_word
    elif out_ext == "pptx" and src_ext == "pdf":
        return pdf_to_ppt
    elif out_ext in ("png", "jpg", "jpeg"):
        if src_ext in ("jpg", "jpeg", "png"):
            return lambda s, d: convert_image(s, d, out_ext)
    return None


# ---------------------------------------------------------------------------
# Version helpers (used by the in-app updater)
# ---------------------------------------------------------------------------
def parse_version(s):
    parts = []
    for p in str(s).strip().split("."):
        p = p.strip()
        try:
            parts.append(int(p))
        except ValueError:
            digits = "".join(ch for ch in p if ch.isdigit())
            parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def is_newer(remote, local):
    """True if remote version string is strictly newer than local."""
    return parse_version(remote) > parse_version(local)
