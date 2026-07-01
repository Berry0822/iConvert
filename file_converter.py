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
)

APP_NAME = "iConvert"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")


def _read_version():
    try:
        with open(os.path.join(APP_DIR, "version.txt"), encoding="utf-8") as f:
            v = f.read().strip()
        if v:
            return v
    except Exception:
        pass
    return "3.1.1"


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
      sub="Word files into PDF", accept={"doc", "docx"}, kind="each", op="word2pdf", out="pdf"),
    T(id="ppt2pdf", badge="PPT", color="#EA580C", title="PowerPoint to PDF",
      sub="Slides into a PDF", accept={"ppt", "pptx"}, kind="each", op="ppt2pdf", out="pdf"),
    T(id="pdf2ppt", badge="PDF", color="#DC2626", title="PDF to PowerPoint",
      sub="PDF pages into slides", accept={"pdf"}, kind="each", op="pdf2ppt", out="pptx"),
    T(id="excel2pdf", badge="XLS", color="#16A34A", title="Excel to PDF",
      sub="Spreadsheets into PDF", accept={"xls", "xlsx"}, kind="each", op="excel2pdf", out="pdf"),
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
]
COLS = 3
EXT_TO_TOOL = {"doc": "word2pdf", "docx": "word2pdf", "ppt": "ppt2pdf",
               "pptx": "ppt2pdf", "xls": "excel2pdf", "xlsx": "excel2pdf",
               "jpg": "jpg2png", "jpeg": "jpg2png", "png": "png2jpg"}


def run_one(tool, src, out_dir, opts):
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
        return pdf_to_images(src, out_dir, base, tool["fmt"])
    if op == "compress":
        o = d("pdf", "_compressed"); pdf_compress(src, o); return [o]
    if op == "split":
        return pdf_split(src, out_dir, base, opts.get("pages", ""))
    if op == "encrypt":
        o = d("pdf", "_protected"); pdf_encrypt(src, o, opts.get("password", "")); return [o]
    if op == "decrypt":
        o = d("pdf", "_unlocked"); pdf_decrypt(src, o, opts.get("password", "")); return [o]
    if op == "ocr_pdf":
        o = d("pdf", "_ocr"); pdf_ocr(src, o); return [o]
    if op == "ocr_word":
        o = d("docx", "_ocr"); pdf_ocr_to_word(src, o); return [o]
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
        self.geometry("1000x780")
        self.minsize(900, 700)
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

        # progress animation state
        self._disp = 0.0
        self._target = 0.0
        self._dot = 0
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
        ctk.CTkLabel(self.home, text="Every tool you need for your files",
                     text_color=TEXT, font=ctk.CTkFont(size=22, weight="bold")
                     ).pack(anchor="w", padx=26, pady=(18, 0))
        ctk.CTkLabel(self.home,
                     text="100% offline on your laptop. Pick a tool to get started.",
                     text_color=MUTED, font=ctk.CTkFont(size=13)
                     ).pack(anchor="w", padx=26, pady=(2, 8))
        grid = ctk.CTkScrollableFrame(self.home, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        for c in range(COLS):
            grid.grid_columnconfigure(c, weight=1, uniform="tiles")
        for i, tool in enumerate(TOOLS):
            r, c = divmod(i, COLS)
            self._make_tile(grid, tool).grid(row=r, column=c, padx=12, pady=12)

    def _make_tile(self, parent, tool):
        card = ctk.CTkFrame(parent, width=270, height=110, corner_radius=14,
                            fg_color=CARD, border_width=1, border_color=TRACK)
        card.grid_propagate(False)
        card.pack_propagate(False)
        badge = ctk.CTkLabel(card, text=tool["badge"], width=48, height=48,
                             corner_radius=12, fg_color=tool["color"], text_color="white",
                             font=ctk.CTkFont(size=14, weight="bold"))
        badge.place(x=16, y=16)
        title = ctk.CTkLabel(card, text=tool["title"], text_color=TEXT, anchor="w",
                             font=ctk.CTkFont(size=15, weight="bold"))
        title.place(x=76, y=22)
        sub = ctk.CTkLabel(card, text=tool["sub"], text_color=MUTED, anchor="w",
                           justify="left", wraplength=165, font=ctk.CTkFont(size=11))
        sub.place(x=76, y=48)

        def on_enter(_e): card.configure(border_color=tool["color"], border_width=2)
        def on_leave(_e): card.configure(border_color=TRACK, border_width=1)
        def on_click(_e): self.show_tool(tool)
        for w in (card, badge, title, sub):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass
        return card

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
        if tool.get("needs") == "pages":
            irow = ctk.CTkFrame(card, fg_color="transparent")
            irow.pack(fill="x", padx=20, pady=(2, 2))
            ctk.CTkLabel(irow, text="Pages", text_color=SOFT_TEXT,
                         font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
            self.pages_entry = ctk.CTkEntry(
                irow, placeholder_text="e.g. 1-3,5  (blank = every page as its own file)",
                height=32)
            self.pages_entry.pack(side="left", fill="x", expand=True)
        elif tool.get("needs") == "password":
            irow = ctk.CTkFrame(card, fg_color="transparent")
            irow.pack(fill="x", padx=20, pady=(2, 2))
            ctk.CTkLabel(irow, text="Password", text_color=SOFT_TEXT,
                         font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))
            self.pw_entry = ctk.CTkEntry(irow, show="*", height=32,
                                         placeholder_text="enter password")
            self.pw_entry.pack(side="left", fill="x", expand=True)

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
        if tool.get("needs") == "pages":
            opts["pages"] = self.pages_entry.get().strip()
        if tool.get("needs") == "password":
            opts["password"] = self.pw_entry.get()
            if not opts["password"]:
                self.show_toast("Enter a password first", "warn")
                return
        if tool["kind"] == "combine" and len(self.files) < 2:
            self.show_toast("Add at least 2 PDFs to merge", "warn")
            return

        self._clear(self.results)
        self.open_btn.pack_forget()
        self._disp = 0.0
        self._target = 0.0
        self.progress.set(0)
        self.pct_lbl.configure(text="0%")
        self.status_lbl.configure(text="Starting...", text_color=MUTED)
        self.working = True
        self.convert_btn.configure(state="disabled", text="Converting...")
        threading.Thread(target=self._worker,
                         args=(tool, list(self.files), self.output_dir, opts),
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
            self.q.put(("start", (1, 1, "Merging %d files" % len(files))))
            try:
                folder = out_dir or os.path.dirname(files[0])
                dst = unique_path(os.path.join(folder, "merged.pdf"))
                pdf_merge(files, dst)
                ok = 1
                self.last_out_dir = folder
                self.q.put(("result", (True, "%d files" % len(files), os.path.basename(dst))))
            except Exception as e:
                self.q.put(("result", (False, "merge", str(e)[:80])))
                sys.stderr.write(traceback.format_exc() + "\n")
            self.q.put(("doneone", (1, 1)))
        else:
            total = len(files)
            for i, src in enumerate(files, start=1):
                self.q.put(("start", (i, total, os.path.basename(src))))
                try:
                    folder = out_dir or os.path.dirname(src)
                    outs = run_one(tool, src, folder, opts)
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
                        self._dot = (self._dot + 1) % 48
                        self.status_lbl.configure(
                            text=self._status_base + "." * (1 + self._dot // 12))
            except Exception:
                pass
        # fast frames only while animating; idle slowly so scrolling stays smooth
        self.after(40 if active else 300, self._tick)

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

    # ---------- queue pump ----------
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "start":
                    i, total, name = payload
                    self._status_base = "Converting %s" % name
                    self._target = (i - 1 + 0.85) / total
                elif kind == "doneone":
                    i, total = payload
                    self._target = i / total
                elif kind == "result":
                    ok, left, right = payload
                    self._add_result(ok, left, right)
                elif kind == "alldone":
                    ok, total = payload
                    self.working = False
                    self._target = 1.0
                    self.convert_btn.configure(state="normal", text="Convert")
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
        except queue.Empty:
            pass
        self.after(120, self._poll)


def main():
    app = App()
    files = [a for a in sys.argv[1:] if os.path.isfile(a)]
    if files:
        try:
            app.handle_launch_file(files[0])
        except Exception:
            pass
    app.mainloop()


if __name__ == "__main__":
    main()
