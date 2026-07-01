"""
iConvert - Local File Converter (v3, modern UI)
Tile-based offline converter. Conversion logic lives in converters.py;
this file is the window/UI + the in-app updater.
"""

import os
import sys
import json
import threading
import queue
import multiprocessing as mp
import traceback
import time
import random
import urllib.request
import webbrowser

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converters import (
    convert_image, image_to_pdf, pdf_to_word, pdf_to_ppt, word_to_pdf, ppt_to_pdf,
    pdf_merge, pdf_split, pdf_compress, pdf_to_images, excel_to_pdf, excel_to_csv,
    pdf_encrypt, pdf_decrypt, pdf_ocr, pdf_ocr_to_word, tesseract_ok,
    gather_files, unique_path, is_newer,
    pdf_rotate, pdf_delete_pages, pdf_add_page_numbers, pdf_watermark,
    pdf_extract_images, images_to_pdf, image_resize, process_batch,
)

APP_NAME = "iConvert"
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = APP_DIR
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
LOG_PATH = os.path.join(APP_DIR, "iConvert.log")


def log_line(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("%s  %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def _read_version():
    for base in (BUNDLE_DIR, APP_DIR):
        try:
            with open(os.path.join(base, "version.txt"), encoding="utf-8") as f:
                v = f.read().strip()
            if v:
                return v
        except Exception:
            pass
    return "3.4.3"


APP_VERSION = _read_version()

GITHUB_USER = "Berry0822"
GITHUB_REPO = "iConvert"
GITHUB_BRANCHES = ["main", "master"]
UPDATE_FILES = ("converters.py", "file_converter.py", "version.txt")
GUIDE_URL = "https://github.com/%s/%s#readme" % (GITHUB_USER, GITHUB_REPO)

# --- palette (light, dark) tuples so dark mode just works ---
PAGE = ("#F4F6FB", "#0F172A")
CARD = ("#FFFFFF", "#1E293B")
TEXT = ("#111827", "#F1F5F9")
MUTED = ("#6B7280", "#94A3B8")
TRACK = ("#E5E7EB", "#334155")
SOFT = ("#F3F4F6", "#334155")
SOFT_HOVER = ("#E5E7EB", "#475569")
SOFT_TEXT = ("#374151", "#E2E8F0")
HEADER_BG = ("#FFFFFF", "#1E293B")
DROP_BORDER = ("#C7D2FE", "#3B4A6B")
# solid status colors (same in both modes)
GREEN = "#16A34A"
RED = "#DC2626"
AMBER = "#D97706"
BLUE = "#2563EB"
ACCENT_RED = "#E5322D"


def _raw_url(branch, fname):
    return "https://raw.githubusercontent.com/{}/{}/{}/{}".format(
        GITHUB_USER, GITHUB_REPO, branch, fname)


def _gh_get(branch, path, timeout):
    stamp = "%d_%d" % (int(time.time()), random.randint(0, 999999))
    api = ("https://api.github.com/repos/{}/{}/contents/{}?ref={}&cb={}"
           .format(GITHUB_USER, GITHUB_REPO, path, branch, stamp))
    try:
        req = urllib.request.Request(api, headers={
            "Accept": "application/vnd.github.raw",
            "Cache-Control": "no-cache", "User-Agent": "iConvert-Updater"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read()
    except Exception:
        url = _raw_url(branch, path) + "?cb=" + stamp
        req = urllib.request.Request(url, headers={
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache", "User-Agent": "iConvert-Updater"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read()


# --- tools ------------------------------------------------------------------
def T(**k):
    return k

TOOLS = [
    T(id="pdf2word", badge="PDF", color="#E5322D", title="PDF to Word",
      sub="PDF into editable Word", accept={"pdf"}, kind="each", op="pdf2word", out="docx"),
    T(id="word2pdf", badge="DOC", color="#2563EB", title="Word to PDF",
      sub="Word files into PDF", accept={"doc", "docx"}, kind="each", op="word2pdf", out="pdf", office=True),
    T(id="ppt2pdf", badge="PPT", color="#EA580C", title="PowerPoint to PDF",
      sub="Slides into a PDF", accept={"ppt", "pptx"}, kind="each", op="ppt2pdf", out="pdf", office=True),
    T(id="pdf2ppt", badge="PDF", color="#DC2626", title="PDF to PowerPoint",
      sub="PDF pages into slides", accept={"pdf"}, kind="each", op="pdf2ppt", out="pptx"),
    T(id="excel2pdf", badge="XLS", color="#16A34A", title="Excel to PDF",
      sub="Spreadsheets into PDF", accept={"xls", "xlsx"}, kind="each", op="excel2pdf", out="pdf", office=True),
    T(id="excel2csv", badge="XLS", color="#0D9488", title="Excel to CSV",
      sub="Each sheet into a CSV", accept={"xls", "xlsx"}, kind="each", op="excel2csv", out=None),
    T(id="jpg2png", badge="JPG", color="#0EA5E9", title="JPG to PNG",
      sub="Images into PNG", accept={"jpg", "jpeg"}, kind="each", op="img", out="png"),
    T(id="png2jpg", badge="PNG", color="#7C3AED", title="PNG to JPG",
      sub="Images into JPG", accept={"png"}, kind="each", op="img", out="jpg"),
    T(id="img2pdf", badge="IMG", color="#D97706", title="Image to PDF",
      sub="JPG/PNG into a PDF", accept={"jpg", "jpeg", "png"}, kind="each", op="img2pdf", out="pdf"),
    T(id="pdf2jpg", badge="PDF", color="#DB2777", title="PDF to JPG",
      sub="Each page as a JPG", accept={"pdf"}, kind="each", op="pdf2img", fmt="jpg", out=None),
    T(id="pdf2png", badge="PDF", color="#9333EA", title="PDF to PNG",
      sub="Each page as a PNG", accept={"pdf"}, kind="each", op="pdf2img", fmt="png", out=None),
    T(id="pdf_merge", badge="PDF", color="#4F46E5", title="Merge PDFs",
      sub="Combine into one PDF", accept={"pdf"}, kind="combine", op="merge", out="pdf"),
    T(id="pdf_split", badge="PDF", color="#0891B2", title="Split PDF",
      sub="Pull out pages", accept={"pdf"}, kind="each", op="split", needs="pages", out=None),
    T(id="pdf_compress", badge="PDF", color="#64748B", title="Compress PDF",
      sub="Shrink the file size", accept={"pdf"}, kind="each", op="compress", out="pdf"),
    T(id="pdf_protect", badge="LCK", color="#334155", title="Protect PDF",
      sub="Add a password", accept={"pdf"}, kind="each", op="encrypt", needs="password", out="pdf"),
    T(id="pdf_unlock", badge="KEY", color="#6B7280", title="Unlock PDF",
      sub="Remove a password", accept={"pdf"}, kind="each", op="decrypt", needs="password", out="pdf"),
    T(id="pdf_ocr", badge="OCR", color="#15803D", title="OCR PDF (searchable)",
      sub="Make scans searchable", accept={"pdf"}, kind="each", op="ocr_pdf", out="pdf", ocr=True),
    T(id="pdf_ocr_word", badge="OCR", color="#1D4ED8", title="OCR PDF to Word",
      sub="Scanned PDF into Word", accept={"pdf"}, kind="each", op="ocr_word", out="docx", ocr=True),
    T(id="pdf_rotate", badge="PDF", color="#0E7490", title="Rotate PDF",
      sub="Turn pages 90/180/270", accept={"pdf"}, kind="each", op="rotate", needs="rotate", out="pdf"),
    T(id="pdf_delpages", badge="PDF", color="#B45309", title="Delete PDF pages",
      sub="Remove pages you list", accept={"pdf"}, kind="each", op="delpages", needs="pages", out=None),
    T(id="pdf_pagenum", badge="PDF", color="#4338CA", title="Add page numbers",
      sub="Stamp page numbers", accept={"pdf"}, kind="each", op="pagenum", out="pdf"),
    T(id="pdf_watermark", badge="PDF", color="#9D174D", title="Watermark PDF",
      sub="Stamp text on pages", accept={"pdf"}, kind="each", op="watermark", needs="text", out="pdf"),
    T(id="pdf_extract_img", badge="PDF", color="#0F766E", title="Extract images",
      sub="Pull images out of a PDF", accept={"pdf"}, kind="each", op="extractimg", out=None),
    T(id="img_combine", badge="IMG", color="#C2410C", title="Images to one PDF",
      sub="Combine images into a PDF", accept={"jpg", "jpeg", "png"}, kind="combine", op="images_pdf", out="pdf"),
    T(id="img_resize", badge="IMG", color="#7E22CE", title="Resize image",
      sub="Shrink big images", accept={"jpg", "jpeg", "png"}, kind="each", op="resize", needs="number", out=None),
]
COLS = 3
EXT_TO_TOOL = {"doc": "word2pdf", "docx": "word2pdf", "ppt": "ppt2pdf",
               "pptx": "ppt2pdf", "xls": "excel2pdf", "xlsx": "excel2pdf",
               "jpg": "jpg2png", "jpeg": "jpg2png", "png": "png2jpg"}


def run_one(tool, src, out_dir, opts, progress_cb=None):
    op = tool["op"]
    base = os.path.splitext(os.path.basename(src))[0]

    def d(ext, suffix=""):
        return unique_path(os.path.join(out_dir, base + suffix + "." + ext))

    if op == "img":
        o = d(tool["out"]); convert_image(src, o, tool["out"]); return [o]
    if op == "img2pdf":
        o = d("pdf"); image_to_pdf(src, o); return [o]
    if op == "pdf2word":
        o = d("docx"); pdf_to_word(src, o); return [o]
    if op == "pdf2ppt":
        o = d("pptx"); pdf_to_ppt(src, o); return [o]
    if op == "word2pdf":
        o = d("pdf"); word_to_pdf(src, o); return [o]
    if op == "ppt2pdf":
        o = d("pdf"); ppt_to_pdf(src, o); return [o]
    if op == "excel2pdf":
        o = d("pdf"); excel_to_pdf(src, o); return [o]
    if op == "excel2csv":
        return excel_to_csv(src, out_dir, base)
    if op == "pdf2img":
        return pdf_to_images(src, out_dir, base, tool["fmt"], progress_cb=progress_cb)
    if op == "compress":
        o = d("pdf", "_compressed"); pdf_compress(src, o); return [o]
    if op == "split":
        return pdf_split(src, out_dir, base, opts.get("pages", ""))
    if op == "encrypt":
        o = d("pdf", "_protected"); pdf_encrypt(src, o, opts.get("password", "")); return [o]
    if op == "decrypt":
        o = d("pdf", "_unlocked"); pdf_decrypt(src, o, opts.get("password", "")); return [o]
    if op == "ocr_pdf":
        o = d("pdf", "_ocr"); pdf_ocr(src, o, progress_cb=progress_cb); return [o]
    if op == "ocr_word":
        o = d("docx", "_ocr"); pdf_ocr_to_word(src, o, progress_cb=progress_cb); return [o]
    if op == "rotate":
        o = d("pdf", "_rotated"); pdf_rotate(src, o, opts.get("degrees", 90)); return [o]
    if op == "delpages":
        return pdf_delete_pages(src, out_dir, base, opts.get("pages", ""))
    if op == "pagenum":
        o = d("pdf", "_numbered"); pdf_add_page_numbers(src, o); return [o]
    if op == "watermark":
        o = d("pdf", "_watermark"); pdf_watermark(src, o, opts.get("text", "")); return [o]
    if op == "extractimg":
        return pdf_extract_images(src, out_dir, base)
    if op == "resize":
        ext = os.path.splitext(src)[1].lstrip(".").lower() or "png"
        o = d(ext, "_small"); image_resize(src, o, opts.get("number", 1600)); return [o]
    raise ValueError("unknown op " + op)


def load_settings():
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(s):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=2)
    except Exception:
        pass


try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

if DND_AVAILABLE:
    class BaseTk(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:
                pass
else:
    BaseTk = ctk.CTk


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", APP_DIR)
    return os.path.join(base, rel)


def _darken(hexc, f=0.85):
    hexc = hexc.lstrip("#")
    r, g, b = (int(hexc[i:i+2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % (int(r*f), int(g*f), int(b*f))


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        ctk.set_appearance_mode(self.settings.get("appearance", "light"))
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME + " - Local File Converter")
        self.geometry(self.settings.get("geometry", "1000x780"))
        self.minsize(900, 700)
        try:
            self.protocol("WM_DELETE_WINDOW", self._on_close)
        except Exception:
            pass
        self.configure(fg_color=PAGE)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.files = []
        self.output_dir = self.settings.get("output_dir") or None
        self.last_out_dir = None
        self.current = None
        self.working = False
        self.q = queue.Queue()
        self._toast_wins = []
        self._badge_cache = {}
        self._tiles = []
        self.proc = None
        self.mpq = None
        self._use_proc = False
        self._cancelled = False

        # progress animation state
        self._disp = 0.0
        self._target = 0.0
        self._dot = 0
        self._spin_angle = 0
        self._cur_i = 1
        self._cur_total = 1
        self._status_base = ""
        self._has_progress = False

        self._build_header()
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True)
        self.home = None
        self.conv_view = ctk.CTkFrame(self.body, fg_color="transparent")
        self.recent_view = ctk.CTkFrame(self.body, fg_color="transparent")
        self._build_home()
        self.show_home()

        self.after(120, self._poll)
        self.after(33, self._tick)
        self.after(900, lambda: self.check_for_updates(silent=True))
        self.after(700, self._whats_new)

    # ---------- header ----------
    def _build_header(self):
        h = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=HEADER_BG)
        h.pack(fill="x")
        h.pack_propagate(False)
        left = ctk.CTkFrame(h, fg_color="transparent")
        left.pack(side="left", padx=18)
        lg = ctk.CTkLabel(left, text="i", width=34, height=34, corner_radius=9,
                          fg_color=ACCENT_RED, text_color="white",
                          font=ctk.CTkFont(size=20, weight="bold"))
        lg.pack(side="left", pady=14)
        lg.bind("<Button-1>", lambda _e: self.show_home())
        ctk.CTkLabel(left, text="Convert", text_color=TEXT,
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=(6, 0))

        right = ctk.CTkFrame(h, fg_color="transparent")
        right.pack(side="right", padx=16)
        ctk.CTkLabel(right, text="v" + APP_VERSION, text_color=MUTED,
                     font=ctk.CTkFont(size=12)).pack(side="right", padx=(10, 2))
        self.dark_switch = ctk.CTkSwitch(right, text="Dark", command=self._toggle_dark,
                                         onvalue="dark", offvalue="light")
        if self.settings.get("appearance", "light") == "dark":
            self.dark_switch.select()
        self.dark_switch.pack(side="right", padx=10)
        ctk.CTkButton(right, text="Recent", width=78, height=30, corner_radius=8,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      command=self.show_recent).pack(side="right", padx=(0, 6))
        ctk.CTkButton(right, text="Check for updates", width=140, height=30,
                      corner_radius=8, fg_color="#EEF2FF", text_color="#3730A3",
                      hover_color="#E0E7FF",
                      command=lambda: self.check_for_updates(False)).pack(side="right", padx=(0, 6))

    def _toggle_dark(self):
        mode = self.dark_switch.get()
        ctk.set_appearance_mode(mode)
        self.settings["appearance"] = mode
        save_settings(self.settings)

    # ---------- home ----------
    def _build_home(self):
        self.home = ctk.CTkFrame(self.body, fg_color="transparent")
        head = ctk.CTkFrame(self.home, fg_color="transparent")
        head.pack(fill="x", padx=26, pady=(16, 4))
        ctk.CTkLabel(head, text="Every tool you need for your files", text_color=TEXT,
                     font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(head, text="100% offline on your laptop. Pick a tool to get started.",
                     text_color=MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", pady=(2, 6))
        self.search = ctk.CTkEntry(head, placeholder_text="Search tools...", height=34)
        self.search.pack(fill="x")
        self.search.bind("<KeyRelease>", lambda _e: self._filter_tiles())

        self.grid = ctk.CTkScrollableFrame(self.home, fg_color="transparent")
        self.grid.pack(fill="both", expand=True, padx=18, pady=(8, 16))
        for c in range(COLS):
            self.grid.grid_columnconfigure(c, weight=1, uniform="tiles")
        self._tiles = []
        for tool in TOOLS:
            self._tiles.append((self._make_tile(self.grid, tool), tool))
        self._filter_tiles()

    def _make_tile(self, parent, tool):
        return ctk.CTkButton(
            parent, image=self._badge_image(tool["badge"], tool["color"]),
            text="  " + tool["title"], compound="left", anchor="w",
            height=58, corner_radius=14, fg_color=CARD, hover_color=SOFT,
            text_color=TEXT, border_width=1, border_color=TRACK,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=lambda t=tool: self.show_tool(t))

    def _badge_image(self, text, color):
        key = (text, color)
        if key in self._badge_cache:
            return self._badge_cache[key]
        from PIL import Image, ImageDraw, ImageFont
        S = 40
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        dr = ImageDraw.Draw(img)
        dr.rounded_rectangle([0, 0, S - 1, S - 1], radius=11, fill=color)
        font = None
        for name in ("segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"):
            try:
                font = ImageFont.truetype(name, 15)
                break
            except Exception:
                pass
        if font is None:
            font = ImageFont.load_default()
        bb = dr.textbbox((0, 0), text, font=font)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
        dr.text(((S - tw) / 2 - bb[0], (S - th) / 2 - bb[1]), text, fill="white", font=font)
        cimg = ctk.CTkImage(light_image=img, dark_image=img, size=(S, S))
        self._badge_cache[key] = cimg
        return cimg

    def _filter_tiles(self):
        try:
            q = self.search.get().strip().lower()
        except Exception:
            q = ""
        idx = 0
        for btn, tool in self._tiles:
            hay = (tool["title"] + " " + tool["sub"] + " " + tool["badge"]).lower()
            if not q or q in hay:
                r, c = divmod(idx, COLS)
                btn.grid(row=r, column=c, padx=10, pady=8, sticky="ew")
                idx += 1
            else:
                btn.grid_remove()

    # ---------- view switching ----------
    def _hide_all(self):
        for v in (self.home, self.conv_view, self.recent_view):
            if v is not None:
                v.pack_forget()

    def show_home(self):
        self._has_progress = False
        self._hide_all()
        self.home.pack(fill="both", expand=True)

    def _clear(self, frame):
        for ch in frame.winfo_children():
            ch.destroy()

    def show_recent(self):
        self._has_progress = False
        self._hide_all()
        self._clear(self.recent_view)
        self.recent_view.pack(fill="both", expand=True)
        card = ctk.CTkFrame(self.recent_view, width=660, height=600, corner_radius=16,
                            fg_color=CARD, border_width=1, border_color=TRACK)
        card.pack(pady=16)
        card.pack_propagate(False)
        topr = ctk.CTkFrame(card, fg_color="transparent")
        topr.pack(fill="x", padx=18, pady=(16, 6))
        ctk.CTkButton(topr, text="< Back", width=70, height=30, corner_radius=8,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      command=self.show_home).pack(side="left")
        ctk.CTkLabel(topr, text="  Recent conversions", text_color=TEXT,
                     font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        lst = ctk.CTkScrollableFrame(card, fg_color="transparent")
        lst.pack(fill="both", expand=True, padx=14, pady=(4, 14))
        recent = self.settings.get("recent", [])
        if not recent:
            ctk.CTkLabel(lst, text="No conversions yet.", text_color=MUTED).pack(pady=20)
        for item in recent:
            row = ctk.CTkFrame(lst, fg_color=PAGE, corner_radius=10)
            row.pack(fill="x", padx=4, pady=4)
            txt = "%s  -  %s file(s)" % (item.get("title", "?"), item.get("count", 0))
            ctk.CTkLabel(row, text=txt, text_color=TEXT, anchor="w",
                         font=ctk.CTkFont(size=13)).pack(side="left", padx=12, pady=8)
            folder = item.get("folder")
            if folder:
                ctk.CTkButton(row, text="Open folder", width=110, height=28,
                              corner_radius=8, fg_color=SOFT, text_color=SOFT_TEXT,
                              hover_color=SOFT_HOVER,
                              command=lambda f=folder: self._open_dir(f)).pack(side="right", padx=10)

    def show_tool(self, tool):
        self.current = tool
        self.settings["last_tool"] = tool["id"]
        save_settings(self.settings)
        self.files = []
        accent = tool["color"]
        self._hide_all()
        self._clear(self.conv_view)
        self.conv_view.pack(fill="both", expand=True)

        card = ctk.CTkFrame(self.conv_view, width=680, height=628, corner_radius=16,
                            fg_color=CARD, border_width=1, border_color=TRACK)
        card.pack(pady=14)
        card.pack_propagate(False)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(16, 4))
        ctk.CTkButton(top, text="< Back", width=70, height=30, corner_radius=8,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      command=self.show_home).pack(side="left")
        ctk.CTkLabel(top, text=tool["badge"], width=44, height=30, corner_radius=8,
                     fg_color=accent, text_color="white",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(10, 8))
        ctk.CTkLabel(top, text=tool["title"], text_color=TEXT,
                     font=ctk.CTkFont(size=19, weight="bold")).pack(side="left")

        drop = ctk.CTkFrame(card, height=92, corner_radius=14, fg_color=PAGE,
                            border_width=2, border_color=DROP_BORDER)
        drop.pack(fill="x", padx=20, pady=8)
        drop.pack_propagate(False)
        msg = ("Drag files here, or use the buttons" if DND_AVAILABLE
               else "Add files or a whole folder")
        ctk.CTkLabel(drop, text=msg, text_color=MUTED,
                     font=ctk.CTkFont(size=13)).pack(pady=(14, 6))
        btns = ctk.CTkFrame(drop, fg_color="transparent")
        btns.pack()
        ctk.CTkButton(btns, text="Select files", width=130, height=34, corner_radius=9,
                      fg_color=accent, hover_color=_darken(accent),
                      font=ctk.CTkFont(size=13, weight="bold"),
                      command=self._add_files).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Select folder", width=130, height=34, corner_radius=9,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      font=ctk.CTkFont(size=13), command=self._add_folder).pack(side="left", padx=4)
        if DND_AVAILABLE:
            try:
                drop.drop_target_register(DND_FILES)
                drop.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        frow = ctk.CTkFrame(card, fg_color="transparent")
        frow.pack(fill="x", padx=20, pady=(2, 0))
        self.count_lbl = ctk.CTkLabel(frow, text="0 file(s)", text_color=SOFT_TEXT,
                                      font=ctk.CTkFont(size=13, weight="bold"))
        self.count_lbl.pack(side="left")
        ctk.CTkButton(frow, text="Clear", width=64, height=26, corner_radius=7,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      command=self._clear_files).pack(side="right")
        self.files_box = ctk.CTkTextbox(card, height=54, corner_radius=10, fg_color=PAGE,
                                        text_color=TEXT, border_width=1, border_color=TRACK)
        self.files_box.pack(fill="x", padx=20, pady=(4, 4))

        # conditional input
        self.pages_entry = None
        self.pw_entry = None
        self.text_entry = None
        self.num_entry = None
        self.rotate_menu = None
        need = tool.get("needs")
        if need:
            irow = ctk.CTkFrame(card, fg_color="transparent")
            irow.pack(fill="x", padx=20, pady=(2, 2))
            if need == "pages":
                ctk.CTkLabel(irow, text="Pages", text_color=SOFT_TEXT,
                             font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
                self.pages_entry = ctk.CTkEntry(
                    irow, height=32,
                    placeholder_text="e.g. 1-3,5  (blank = every page as its own file)")
                self.pages_entry.pack(side="left", fill="x", expand=True)
            elif need == "password":
                ctk.CTkLabel(irow, text="Password", text_color=SOFT_TEXT,
                             font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
                self.pw_entry = ctk.CTkEntry(irow, show="*", height=32,
                                             placeholder_text="enter password")
                self.pw_entry.pack(side="left", fill="x", expand=True)
            elif need == "text":
                ctk.CTkLabel(irow, text="Text", text_color=SOFT_TEXT,
                             font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
                self.text_entry = ctk.CTkEntry(
                    irow, height=32, placeholder_text="watermark text, e.g. CONFIDENTIAL")
                self.text_entry.pack(side="left", fill="x", expand=True)
            elif need == "number":
                ctk.CTkLabel(irow, text="Max size (px)", text_color=SOFT_TEXT,
                             font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
                self.num_entry = ctk.CTkEntry(irow, height=32, placeholder_text="e.g. 1600")
                self.num_entry.pack(side="left", fill="x", expand=True)
            elif need == "rotate":
                ctk.CTkLabel(irow, text="Rotate", text_color=SOFT_TEXT,
                             font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
                self.rotate_menu = ctk.CTkOptionMenu(
                    irow, values=["90 right", "180", "270 left"], width=150,
                    fg_color=accent, button_color=_darken(accent),
                    button_hover_color=_darken(accent, 0.75))
                self.rotate_menu.set("90 right")
                self.rotate_menu.pack(side="left")

        orow = ctk.CTkFrame(card, fg_color="transparent")
        orow.pack(fill="x", padx=20, pady=(2, 2))
        self.out_lbl = ctk.CTkLabel(orow, text=self._out_text(), text_color=MUTED,
                                    font=ctk.CTkFont(size=12))
        self.out_lbl.pack(side="left")
        ctk.CTkButton(orow, text="Output folder...", width=130, height=28, corner_radius=8,
                      fg_color=SOFT, text_color=SOFT_TEXT, hover_color=SOFT_HOVER,
                      command=self._choose_output).pack(side="right")

        self.convert_btn = ctk.CTkButton(
            card, text="Convert", height=42, corner_radius=10, fg_color=accent,
            hover_color=_darken(accent), font=ctk.CTkFont(size=16, weight="bold"),
            command=self._start)
        self.convert_btn.pack(fill="x", padx=20, pady=(6, 6))

        prow = ctk.CTkFrame(card, fg_color="transparent")
        prow.pack(fill="x", padx=20)
        _dark = ctk.get_appearance_mode() == "Dark"
        self.spinner = tk.Canvas(prow, width=22, height=22, highlightthickness=0, bd=0,
                                 bg="#1E293B" if _dark else "#FFFFFF")
        self.spinner.pack(side="left", padx=(0, 8))
        self._spin_angle = 0
        self.status_lbl = ctk.CTkLabel(prow, text="Ready to convert", text_color=MUTED,
                                       font=ctk.CTkFont(size=13))
        self.status_lbl.pack(side="left")
        self.pct_lbl = ctk.CTkLabel(prow, text="0%", text_color=accent,
                                    font=ctk.CTkFont(size=20, weight="bold"))
        self.pct_lbl.pack(side="right")
        self.progress = ctk.CTkProgressBar(card, height=14, corner_radius=7,
                                           fg_color=TRACK, progress_color=accent)
        self.progress.set(0)
        self.progress.pack(fill="x", padx=20, pady=(4, 8))
        self._disp = 0.0
        self._target = 0.0
        self._has_progress = True

        self.results = ctk.CTkScrollableFrame(card, fg_color=PAGE, corner_radius=10,
                                              label_text="Results", label_text_color=MUTED)
        self.results.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        self.cancel_btn = ctk.CTkButton(card, text="Cancel", height=34, corner_radius=9,
                                        fg_color=SOFT, text_color=RED,
                                        hover_color=SOFT_HOVER, command=self._cancel)
        self.cancel_btn.pack_forget()
        self.open_btn = ctk.CTkButton(card, text="Open output folder", height=34,
                                      corner_radius=9, fg_color=GREEN,
                                      hover_color=_darken(GREEN), command=self._open_output)
        self.open_btn.pack(fill="x", padx=20, pady=(0, 14))
        self.open_btn.pack_forget()
        self._refresh_files()

    def _out_text(self):
        if self.output_dir:
            return "Saving to: " + os.path.basename(self.output_dir)
        return "Saved beside each original"

    # ---------- files ----------
    def _add_paths(self, paths):
        files = gather_files(paths, self.current["accept"])
        added = 0
        for p in files:
            p = os.path.abspath(p)
            if p not in self.files:
                self.files.append(p)
                added += 1
        if added:
            self.show_toast("Added %d file(s)" % added, "info")
        elif paths:
            self.show_toast("No matching files found", "warn")
        self._refresh_files()

    def _on_drop(self, event):
        self._add_paths(self.tk.splitlist(event.data))

    def _add_files(self):
        exts = " ".join("*." + e for e in sorted(self.current["accept"]))
        paths = filedialog.askopenfilenames(
            title="Choose files",
            filetypes=[(self.current["title"] + " input", exts), ("All files", "*.*")])
        if paths:
            self._add_paths(paths)

    def _add_folder(self):
        d = filedialog.askdirectory(title="Choose a folder (all matching files inside)")
        if d:
            self._add_paths([d])

    def _clear_files(self):
        self.files = []
        self._refresh_files()

    def _refresh_files(self):
        self.count_lbl.configure(text="%d file(s)" % len(self.files))
        self.files_box.configure(state="normal")
        self.files_box.delete("1.0", "end")
        if self.files:
            for f in self.files:
                self.files_box.insert("end", "  " + os.path.basename(f) + "\n")
        else:
            self.files_box.insert("end", "  No files added yet.\n")
        self.files_box.configure(state="disabled")

    def _choose_output(self):
        d = filedialog.askdirectory(title="Choose output folder")
        self.output_dir = d or None
        self.settings["output_dir"] = self.output_dir
        save_settings(self.settings)
        self.out_lbl.configure(text=self._out_text())

    # ---------- convert ----------
    def _start(self):
        if self.working:
            return
        tool = self.current
        if tool.get("ocr") and not tesseract_ok():
            self.show_toast("Tesseract not found. Install it (guide below), then reopen iConvert.",
                            "error")
            return
        if not self.files:
            self.show_toast("Add some files first", "warn")
            return
        opts = {}
        need = tool.get("needs")
        if need == "pages":
            opts["pages"] = self.pages_entry.get().strip()
        elif need == "password":
            opts["password"] = self.pw_entry.get()
            if not opts["password"]:
                self.show_toast("Enter a password first", "warn")
                return
        elif need == "text":
            opts["text"] = self.text_entry.get().strip()
            if not opts["text"]:
                self.show_toast("Enter the watermark text", "warn")
                return
        elif need == "number":
            v = self.num_entry.get().strip()
            if not v.isdigit() or int(v) <= 0:
                self.show_toast("Enter a size in pixels, e.g. 1600", "warn")
                return
            opts["number"] = int(v)
        elif need == "rotate":
            opts["degrees"] = {"90 right": 90, "180": 180, "270 left": 270}.get(
                self.rotate_menu.get(), 90)
        if tool["kind"] == "combine" and len(self.files) < 2:
            self.show_toast("Add at least 2 files", "warn")
            return

        self._clear(self.results)
        self.open_btn.pack_forget()
        self._disp = 0.0
        self._target = 0.0
        self.progress.set(0)
        self.pct_lbl.configure(text="0%")
        self.status_lbl.configure(text="Starting...", text_color=MUTED)
        self.working = True
        self._cancelled = False
        self.convert_btn.configure(state="disabled", text="Converting...")
        self.cancel_btn.pack(fill="x", padx=20, pady=(0, 8))
        files = list(self.files)
        self._end_proc()  # clear any leftover worker before starting a new one
        started = False
        if not tool.get("office"):
            # Office (Word/Excel/PowerPoint) automation must run in the main
            # process; everything else runs in a background process for smoothness.
            try:
                self.mpq = mp.Queue()
                self.proc = mp.Process(target=process_batch,
                                       args=(tool, files, self.output_dir, opts, self.mpq),
                                       daemon=True)
                self.proc.start()
                self._use_proc = True
                started = True
            except Exception:
                self._use_proc = False
                self.proc = None
                self.mpq = None
        if not started:
            self._use_proc = False
            threading.Thread(target=self._worker,
                             args=(tool, files, self.output_dir, opts),
                             daemon=True).start()

    def _worker(self, tool, files, out_dir, opts):
        com = False
        try:
            import pythoncom
            pythoncom.CoInitialize()
            com = True
        except Exception:
            pass
        ok = 0
        if tool["kind"] == "combine":
            total = 1
            self.q.put(("start", (1, 1, "Combining %d files" % len(files))))
            try:
                folder = out_dir or os.path.dirname(files[0])
                if tool["op"] == "images_pdf":
                    dst = unique_path(os.path.join(folder, "combined.pdf"))
                    images_to_pdf(files, dst)
                else:
                    dst = unique_path(os.path.join(folder, "merged.pdf"))
                    pdf_merge(files, dst)
                ok = 1
                self.last_out_dir = folder
                self.q.put(("result", (True, "%d files" % len(files), os.path.basename(dst))))
            except Exception as e:
                self.q.put(("result", (False, tool["title"], str(e)[:80])))
                sys.stderr.write(traceback.format_exc() + "\n")
            self.q.put(("doneone", (1, 1)))
        else:
            total = len(files)
            for i, src in enumerate(files, start=1):
                if self._cancelled:
                    break
                self.q.put(("start", (i, total, os.path.basename(src))))

                def cb(done, pages):
                    self.q.put(("sub", (done, pages)))

                try:
                    folder = out_dir or os.path.dirname(src)
                    outs = run_one(tool, src, folder, opts, progress_cb=cb)
                    ok += 1
                    self.last_out_dir = folder
                    label = (os.path.basename(outs[0]) if len(outs) == 1
                             else "%d files" % len(outs))
                    self.q.put(("result", (True, os.path.basename(src), label)))
                except Exception as e:
                    self.q.put(("result", (False, os.path.basename(src), str(e)[:80])))
                    sys.stderr.write(traceback.format_exc() + "\n")
                self.q.put(("doneone", (i, total)))
        if com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        self.q.put(("alldone", (ok, total)))

    def _add_result(self, ok, left, right):
        rowf = ctk.CTkFrame(self.results, fg_color="transparent")
        rowf.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(rowf, text="●", width=14, text_color=GREEN if ok else RED,
                     font=ctk.CTkFont(size=14)).pack(side="left")
        if ok:
            txt = left + "   ->   " + right
            col = SOFT_TEXT
        else:
            txt = left + "  -  failed: " + right
            col = RED
        ctk.CTkLabel(rowf, text=txt, text_color=col, anchor="w",
                     font=ctk.CTkFont(size=12)).pack(side="left", padx=4)

    # ---------- toasts (top-right mini-windows with real fade in/out) ----------
    def show_toast(self, text, kind="info", link=None):
        cmap = {"success": GREEN, "error": RED, "warn": AMBER, "info": BLUE}
        c = cmap.get(kind, BLUE)
        if link is None and kind == "error":
            link = GUIDE_URL
        dark = ctk.get_appearance_mode() == "Dark"
        bg = "#1E293B" if dark else "#FFFFFF"
        fg = "#F1F5F9" if dark else "#111827"
        linkfg = "#93C5FD" if dark else "#1D4ED8"
        try:
            win = tk.Toplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            try:
                win.attributes("-alpha", 0.0)
            except Exception:
                pass
            win.configure(bg=c)
        except Exception:
            return
        inner = tk.Frame(win, bg=bg)
        inner.pack(fill="both", expand=True, padx=2, pady=2)
        tk.Frame(inner, bg=c, width=6).pack(side="left", fill="y")
        bodyf = tk.Frame(inner, bg=bg)
        bodyf.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        tk.Label(bodyf, text=text, bg=bg, fg=fg, justify="left", wraplength=260,
                 font=("Segoe UI", 10)).pack(anchor="w")
        if link:
            tk.Label(bodyf, text="Click here to open the guide", bg=bg, fg=linkfg,
                     cursor="hand2", font=("Segoe UI", 9, "underline")).pack(anchor="w", pady=(4, 0))

        def dismiss(_e=None):
            self._fade_out(win)

        def openlink(_e=None):
            try:
                webbrowser.open(link)
            except Exception:
                pass
            self._fade_out(win)

        handler = openlink if link else dismiss
        for w in (win, inner, bodyf):
            w.bind("<Button-1>", handler)
        for ch in bodyf.winfo_children():
            ch.bind("<Button-1>", handler)
        win._alpha = 0.0
        win._goal = 0.96
        self._toast_wins.append(win)
        self._place_toasts()
        self._fade_in(win)
        win.after(6000 if link else 3400, lambda: self._fade_out(win))

    def _place_toasts(self):
        try:
            self.update_idletasks()
            base_x = self.winfo_rootx() + self.winfo_width() - 20
            y = self.winfo_rooty() + 72
        except Exception:
            return
        for win in list(self._toast_wins):
            try:
                win.update_idletasks()
                w = max(win.winfo_reqwidth(), 300)
                h = win.winfo_reqheight()
                win.geometry("%dx%d+%d+%d" % (w, h, base_x - w, y))
                y += h + 10
            except Exception:
                pass

    def _fade_in(self, win):
        try:
            a = min(getattr(win, "_alpha", 0.0) + 0.12, win._goal)
            win._alpha = a
            win.attributes("-alpha", a)
            if a < win._goal:
                win.after(16, lambda: self._fade_in(win))
        except Exception:
            pass

    def _fade_out(self, win):
        try:
            a = getattr(win, "_alpha", 0.96) - 0.12
            win._alpha = a
            if a <= 0:
                self._destroy_toast(win)
                return
            win.attributes("-alpha", a)
            win.after(16, lambda: self._fade_out(win))
        except Exception:
            self._destroy_toast(win)

    def _destroy_toast(self, win):
        if win in self._toast_wins:
            self._toast_wins.remove(win)
        try:
            win.destroy()
        except Exception:
            pass
        self._place_toasts()

    # ---------- progress animation ----------
    def _tick(self):
        active = False
        if self._has_progress and self.current:
            try:
                if self.working or abs(self._target - self._disp) > 0.001:
                    active = True
                    self._disp += (self._target - self._disp) * 0.18
                    if abs(self._target - self._disp) < 0.002:
                        self._disp = self._target
                    self.progress.set(self._disp)
                    self.pct_lbl.configure(text="%d%%" % int(round(self._disp * 100)))
                if self.working:
                    active = True
                    self._spin_angle = (self._spin_angle + 40) % 360
                    self._draw_spinner(True)
                else:
                    self._draw_spinner(False)
            except Exception:
                pass
        # fast frames only while animating; idle slowly so scrolling stays smooth
        self.after(40 if active else 300, self._tick)

    def _draw_spinner(self, on):
        cv = getattr(self, "spinner", None)
        if cv is None:
            return
        try:
            cv.delete("all")
            if on:
                col = self.current["color"] if self.current else BLUE
                cv.create_arc(3, 3, 19, 19, start=self._spin_angle, extent=270,
                              style="arc", outline=col, width=3)
        except Exception:
            pass

    # ---------- updates ----------
    def check_for_updates(self, silent=False):
        threading.Thread(target=self._update_worker, args=(silent,), daemon=True).start()

    def _update_worker(self, silent):
        if not GITHUB_USER or "YOUR_GITHUB" in GITHUB_USER:
            if not silent:
                self.q.put(("toast", ("Updates not set up - see GITHUB_SETUP.md", "warn")))
            return
        remote = None
        branch_found = None
        for branch in GITHUB_BRANCHES:
            try:
                remote = _gh_get(branch, "version.txt", 10).decode("utf-8").strip()
                branch_found = branch
                break
            except Exception:
                pass
        if not remote:
            if not silent:
                self.q.put(("toast", ("Couldn't reach your repo - is it public?", "error")))
            return
        if is_newer(remote, APP_VERSION):
            self.q.put(("update", (remote, branch_found)))
        elif not silent:
            self.q.put(("toast", ("You're up to date (v%s)" % APP_VERSION, "success")))

    def _do_update(self, remote, branch):
        try:
            for fname in UPDATE_FILES:
                data = _gh_get(branch, fname, 25)
                with open(os.path.join(APP_DIR, fname), "wb") as fh:
                    fh.write(data)
            messagebox.showinfo(APP_NAME, "Updated to version %s.\nPlease close and "
                                          "reopen iConvert to finish." % remote)
            self.destroy()
        except Exception as e:
            self.show_toast("Update failed: %s" % e, "error")

    # ---------- misc ----------
    def _open_dir(self, d):
        if d and os.path.isdir(d):
            try:
                os.startfile(d)
            except Exception:
                self.show_toast("Saved in: " + d, "info")

    def _open_output(self):
        self._open_dir(self.last_out_dir or self.output_dir)

    def handle_launch_file(self, path):
        ext = os.path.splitext(path)[1].lstrip(".").lower()
        tid = EXT_TO_TOOL.get(ext)
        if tid:
            tool = next(t for t in TOOLS if t["id"] == tid)
            self.show_tool(tool)
            self._add_paths([path])
        else:
            self.show_toast("Pick a tool for " + os.path.basename(path), "info")

    def _remember_recent(self, tool, folder, count):
        rec = self.settings.get("recent", [])
        rec.insert(0, {"title": tool["title"], "folder": folder,
                       "count": count, "ts": int(time.time())})
        self.settings["recent"] = rec[:20]
        save_settings(self.settings)

    def _on_close(self):
        try:
            self.settings["geometry"] = self.geometry()
            save_settings(self.settings)
        except Exception:
            pass
        self.destroy()

    def _whats_new(self):
        seen = self.settings.get("seen_version")
        if seen and seen != APP_VERSION:
            self.show_toast("Updated to iConvert v%s" % APP_VERSION, "success")
        self.settings["seen_version"] = APP_VERSION
        save_settings(self.settings)

    # ---------- queue pump ----------
    def _end_proc(self):
        p, q = self.proc, self.mpq
        self.proc = None
        self.mpq = None
        self._use_proc = False
        if p is not None:
            try:
                if p.is_alive():
                    p.join(timeout=1.0)
                if p.is_alive():
                    p.terminate()
                    p.join(timeout=1.0)
            except Exception:
                pass
        if q is not None:
            # release the OS handles this queue used, so they can't pile up
            try:
                q.close()
            except Exception:
                pass
            try:
                q.join_thread()
            except Exception:
                pass

    def _cancel(self):
        if not self.working:
            return
        self._cancelled = True
        if self._use_proc and self.proc is not None:
            try:
                self.proc.terminate()
            except Exception:
                pass
            self._end_proc()
            self.working = False
            self.convert_btn.configure(state="normal", text="Convert")
            if hasattr(self, "cancel_btn"):
                self.cancel_btn.pack_forget()
            self._status_base = "Cancelled"
            self.status_lbl.configure(text="Cancelled", text_color=AMBER)
            self.show_toast("Cancelled", "warn")

    def _handle_msg(self, kind, payload):
        if kind == "start":
            i, total, name = payload
            self._cur_i = i
            self._cur_total = total
            if self.current.get("kind") == "combine":
                base = name
            else:
                verb = "Scanning" if self.current.get("ocr") else "Converting"
                base = "%s %s  (%d of %d)" % (verb, name, i, total)
            self._status_base = base
            self.status_lbl.configure(text=base, text_color=MUTED)
            self._target = (i - 1 + 0.6) / total
        elif kind == "sub":
            done, pages = payload
            if pages:
                self._target = (self._cur_i - 1 + done / float(pages)) / self._cur_total
        elif kind == "doneone":
            i, total = payload
            self._target = i / total
        elif kind == "lastdir":
            self.last_out_dir = payload
        elif kind == "result":
            ok, left, right = payload
            self._add_result(ok, left, right)
            log_line(("OK   " if ok else "FAIL ") + left + " -> " + right)
        elif kind == "alldone":
            ok, total = payload
            log_line("done %d/%d  [%s]" % (ok, total, self.current.get("title", "")))
            self.working = False
            self._target = 1.0
            self.convert_btn.configure(state="normal", text="Convert")
            self._end_proc()
            if hasattr(self, "cancel_btn"):
                self.cancel_btn.pack_forget()
            if ok == total and ok > 0:
                self._status_base = "Done - %d of %d converted" % (ok, total)
                self.status_lbl.configure(text=self._status_base, text_color=GREEN)
                self.show_toast("Done! %d file(s) converted" % ok, "success")
            elif ok > 0:
                self._status_base = "Finished - %d of %d converted" % (ok, total)
                self.status_lbl.configure(text=self._status_base, text_color=AMBER)
                self.show_toast("%d of %d converted" % (ok, total), "warn", link=GUIDE_URL)
            else:
                self._status_base = "Nothing converted"
                self.status_lbl.configure(text=self._status_base, text_color=RED)
                self.show_toast("No files were converted", "error")
            if ok:
                self.open_btn.pack(fill="x", padx=20, pady=(0, 14))
                self._remember_recent(self.current,
                                      self.last_out_dir or self.output_dir, ok)
        elif kind == "toast":
            self.show_toast(payload[0], payload[1])
        elif kind == "update":
            remote, branch = payload
            if messagebox.askyesno(
                    APP_NAME, "Version %s is available (you have %s).\nUpdate now?"
                              % (remote, APP_VERSION)):
                self._do_update(remote, branch)

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                self._handle_msg(kind, payload)
        except queue.Empty:
            pass
        if self._use_proc and self.mpq is not None:
            try:
                while True:
                    kind, payload = self.mpq.get_nowait()
                    self._handle_msg(kind, payload)
            except queue.Empty:
                pass
        if (self._use_proc and self.working and self.proc is not None
                and not self.proc.is_alive()):
            self._handle_msg("alldone", (0, 1))
        self.after(120, self._poll)


def main():
    mp.freeze_support()
    app = App()
    # Always open on the home page. (A right-click "Convert with iConvert"
    # launch is the only time we jump straight to a tool.)
    files = [a for a in sys.argv[1:] if os.path.isfile(a)]
    if files:
        try:
            app.handle_launch_file(files[0])
        except Exception:
            pass
    app.mainloop()


if __name__ == "__main__":
    main()
