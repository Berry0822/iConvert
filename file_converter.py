"""
iConvert - Local File Converter (modern UI)
===========================================
A tile-based, offline desktop converter inspired by iLovePDF.
Conversion logic lives in converters.py; this file is the window/UI + updater.

Supported: PDF<->Word, PPT->PDF, PDF->PowerPoint, JPG<->PNG, image->PDF.
Everything runs locally - no files leave your laptop (except the optional
'Check for updates', which only contacts your own GitHub repo).
"""

import os
import sys
import threading
import queue
import traceback
import urllib.request
import time
import random

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converters import route, unique_path, is_newer

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
APP_NAME = "iConvert"
def _read_version():
    # Single source of truth: the version lives in version.txt next to this file.
    # (Keeping it out of the code means an update can never drag the number
    # backwards or disagree with itself.)
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, "version.txt"), encoding="utf-8") as _f:
            _v = _f.read().strip()
        if _v:
            return _v
    except Exception:
        pass
    return "2.1.2"


APP_VERSION = _read_version()

# Set these after you create your GitHub repo (see GITHUB_SETUP.md).
# Example: GITHUB_USER = "looxu"  ->  the in-app "Check for updates" will work.
GITHUB_USER = "Berry0822"
GITHUB_REPO = "iConvert"
GITHUB_BRANCHES = ["main", "master"]   # tries each until one responds
UPDATE_FILES = ("converters.py", "file_converter.py", "version.txt")


def _raw_url(branch, fname):
    return "https://raw.githubusercontent.com/{}/{}/{}/{}".format(
        GITHUB_USER, GITHUB_REPO, branch, fname)


def _gh_get(branch, path, timeout):
    # Read a file from the repo, always fresh. The GitHub API is NOT served by
    # the raw CDN (which caches for ~5 min), so a just-pushed change shows up
    # immediately. Falls back to the raw URL (with a cache-buster) if the API
    # is unavailable, e.g. rate-limited.
    stamp = "%d_%d" % (int(time.time()), random.randint(0, 999999))
    api = ("https://api.github.com/repos/{}/{}/contents/{}?ref={}&cb={}"
           .format(GITHUB_USER, GITHUB_REPO, path, branch, stamp))
    try:
        req = urllib.request.Request(api, headers={
            "Accept": "application/vnd.github.raw",
            "Cache-Control": "no-cache",
            "User-Agent": "iConvert-Updater",
        })
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read()
    except Exception:
        url = _raw_url(branch, path) + "?cb=" + stamp
        req = urllib.request.Request(url, headers={
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
            "User-Agent": "iConvert-Updater",
        })
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read()

# Optional drag-and-drop (works without it via the Select files button)
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

# ----------------------------------------------------------------------------
# Conversions shown as tiles
# ----------------------------------------------------------------------------
CONVERSIONS = [
    dict(id="pdf2word", badge="PDF", color="#E5322D", title="PDF to Word",
         sub="Make PDFs into editable Word docs", src={"pdf"}, out="docx"),
    dict(id="word2pdf", badge="DOC", color="#2563EB", title="Word to PDF",
         sub="Turn Word files into PDF", src={"doc", "docx"}, out="pdf"),
    dict(id="ppt2pdf", badge="PPT", color="#EA580C", title="PowerPoint to PDF",
         sub="Turn slides into a PDF", src={"ppt", "pptx"}, out="pdf"),
    dict(id="pdf2ppt", badge="PDF", color="#DC2626", title="PDF to PowerPoint",
         sub="Turn PDF pages into slides", src={"pdf"}, out="pptx"),
    dict(id="jpg2png", badge="JPG", color="#0D9488", title="JPG to PNG",
         sub="Convert images to PNG", src={"jpg", "jpeg"}, out="png"),
    dict(id="png2jpg", badge="PNG", color="#7C3AED", title="PNG to JPG",
         sub="Convert images to JPG", src={"png"}, out="jpg"),
    dict(id="img2pdf", badge="IMG", color="#D97706", title="Image to PDF",
         sub="Combine JPG/PNG into a PDF", src={"jpg", "jpeg", "png"}, out="pdf"),
]
COLS = 3


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ----------------------------------------------------------------------------
# App
# ----------------------------------------------------------------------------
class App(BaseTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME + " - Local File Converter")
        self.geometry("880x660")
        self.minsize(780, 580)
        self.configure(fg_color="#F4F6FB")
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.files = []
        self.output_dir = None
        self.current = None
        self.working = False
        self.q = queue.Queue()

        self._build_header()
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True)
        self._build_home()
        self._build_conv_view()
        self.show_home()

        self.after(120, self._poll)
        # quiet auto-check on launch (only prompts if a newer version exists)
        self.after(800, lambda: self.check_for_updates(silent=True))

    # ---------- header ----------
    def _build_header(self):
        h = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color="#FFFFFF")
        h.pack(fill="x")
        h.pack_propagate(False)
        left = ctk.CTkFrame(h, fg_color="transparent")
        left.pack(side="left", padx=18)
        ctk.CTkLabel(left, text="i", width=34, height=34, corner_radius=9,
                     fg_color="#E5322D", text_color="white",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", pady=14)
        ctk.CTkLabel(left, text="Convert", text_color="#111827",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=(6, 0))

        right = ctk.CTkFrame(h, fg_color="transparent")
        right.pack(side="right", padx=16)
        ctk.CTkLabel(right, text="v" + APP_VERSION, text_color="#9CA3AF",
                     font=ctk.CTkFont(size=12)).pack(side="right", padx=(8, 4))
        self.update_btn = ctk.CTkButton(right, text="Check for updates", width=140,
                                        height=30, corner_radius=8,
                                        fg_color="#EEF2FF", text_color="#3730A3",
                                        hover_color="#E0E7FF",
                                        command=lambda: self.check_for_updates(False))
        self.update_btn.pack(side="right")

    # ---------- home (tiles) ----------
    def _build_home(self):
        self.home = ctk.CTkFrame(self.body, fg_color="transparent")
        title = ctk.CTkLabel(self.home, text="Every tool you need for your files",
                             text_color="#111827",
                             font=ctk.CTkFont(size=22, weight="bold"))
        title.pack(anchor="w", padx=26, pady=(18, 0))
        ctk.CTkLabel(self.home,
                     text="100% offline on your laptop. Pick a tool to get started.",
                     text_color="#6B7280", font=ctk.CTkFont(size=13)).pack(
            anchor="w", padx=26, pady=(2, 8))

        grid = ctk.CTkScrollableFrame(self.home, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        for c in range(COLS):
            grid.grid_columnconfigure(c, weight=1, uniform="tiles")
        for i, conv in enumerate(CONVERSIONS):
            r, c = divmod(i, COLS)
            self._make_tile(grid, conv).grid(row=r, column=c, padx=12, pady=12)

    def _make_tile(self, parent, conv):
        card = ctk.CTkFrame(parent, width=250, height=118, corner_radius=14,
                            fg_color="#FFFFFF", border_width=1,
                            border_color="#E5E7EB")
        card.grid_propagate(False)
        card.pack_propagate(False)

        badge = ctk.CTkLabel(card, text=conv["badge"], width=50, height=50,
                             corner_radius=12, fg_color=conv["color"],
                             text_color="white",
                             font=ctk.CTkFont(size=14, weight="bold"))
        badge.place(x=16, y=16)
        title = ctk.CTkLabel(card, text=conv["title"], text_color="#111827",
                             anchor="w", font=ctk.CTkFont(size=15, weight="bold"))
        title.place(x=78, y=22)
        sub = ctk.CTkLabel(card, text=conv["sub"], text_color="#6B7280",
                           anchor="w", justify="left", wraplength=150,
                           font=ctk.CTkFont(size=11))
        sub.place(x=78, y=48)

        def on_enter(_e):
            card.configure(border_color=conv["color"], border_width=2)
        def on_leave(_e):
            card.configure(border_color="#E5E7EB", border_width=1)
        def on_click(_e):
            self.show_conversion(conv)

        for w in (card, badge, title, sub):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass
        return card

    # ---------- conversion view ----------
    def _build_conv_view(self):
        self.conv_view = ctk.CTkFrame(self.body, fg_color="transparent")

    def _clear(self, frame):
        for ch in frame.winfo_children():
            ch.destroy()

    def show_home(self):
        self.conv_view.pack_forget()
        self.home.pack(fill="both", expand=True)

    def show_conversion(self, conv):
        self.current = conv
        self.files = []
        self.output_dir = None
        self.home.pack_forget()
        self._clear(self.conv_view)
        self.conv_view.pack(fill="both", expand=True)

        top = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        top.pack(fill="x", padx=22, pady=(16, 6))
        ctk.CTkButton(top, text="< Back", width=70, height=30, corner_radius=8,
                      fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self.show_home).pack(side="left")
        ctk.CTkLabel(top, text="   " + conv["title"], text_color="#111827",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        # drop / select area
        drop = ctk.CTkFrame(self.conv_view, height=120, corner_radius=14,
                            fg_color="#FFFFFF", border_width=2,
                            border_color="#DBE3FF")
        drop.pack(fill="x", padx=22, pady=8)
        drop.pack_propagate(False)
        msg = ("Drag files here, or" if DND_AVAILABLE else "Add the files you want to convert")
        ctk.CTkLabel(drop, text=msg, text_color="#6B7280",
                     font=ctk.CTkFont(size=14)).pack(pady=(24, 6))
        ctk.CTkButton(drop, text="Select files", width=150, height=38,
                      corner_radius=9, fg_color=conv["color"],
                      hover_color=self._darken(conv["color"]),
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._add_files).pack()
        if DND_AVAILABLE:
            try:
                drop.drop_target_register(DND_FILES)
                drop.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # file list
        row = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        row.pack(fill="x", padx=22, pady=(4, 0))
        ctk.CTkLabel(row, text="Files", text_color="#374151",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(row, text="Clear", width=64, height=26, corner_radius=7,
                      fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self._clear_files).pack(side="right")
        self.files_box = ctk.CTkTextbox(self.conv_view, height=96, corner_radius=10,
                                        fg_color="#FFFFFF", text_color="#111827",
                                        border_width=1, border_color="#E5E7EB")
        self.files_box.pack(fill="x", padx=22, pady=(4, 6))
        self.files_box.configure(state="disabled")

        # output + convert
        out = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        out.pack(fill="x", padx=22, pady=(2, 4))
        self.out_lbl = ctk.CTkLabel(out, text="Saved beside each original",
                                    text_color="#6B7280", font=ctk.CTkFont(size=12))
        self.out_lbl.pack(side="left")
        ctk.CTkButton(out, text="Output folder...", width=130, height=28,
                      corner_radius=8, fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self._choose_output).pack(side="right")

        self.convert_btn = ctk.CTkButton(
            self.conv_view, text="Convert", height=44, corner_radius=10,
            fg_color=conv["color"], hover_color=self._darken(conv["color"]),
            font=ctk.CTkFont(size=16, weight="bold"), command=self._start)
        self.convert_btn.pack(fill="x", padx=22, pady=(6, 6))

        self.progress = ctk.CTkProgressBar(self.conv_view, height=10)
        self.progress.set(0)
        self.progress.pack(fill="x", padx=22, pady=(0, 6))
        self.log = ctk.CTkTextbox(self.conv_view, height=120, corner_radius=10,
                                  fg_color="#0F172A", text_color="#E2E8F0",
                                  font=ctk.CTkFont(family="Consolas", size=12))
        self.log.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        self.log.configure(state="disabled")
        self._log("Ready. Add %s file(s) to convert to %s." % (
            "/".join(sorted(conv["src"])).upper(), conv["out"].upper()))

    # ---------- helpers ----------
    @staticmethod
    def _darken(hexc, f=0.85):
        hexc = hexc.lstrip("#")
        r, g, b = (int(hexc[i:i+2], 16) for i in (0, 2, 4))
        return "#%02x%02x%02x" % (int(r*f), int(g*f), int(b*f))

    def _add_paths(self, paths):
        conv = self.current
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if not os.path.isfile(p):
                continue
            ext = os.path.splitext(p)[1].lstrip(".").lower()
            if ext not in conv["src"]:
                self._log("Skipped (not %s): %s" % (
                    "/".join(sorted(conv["src"])).upper(), os.path.basename(p)))
                continue
            if p not in self.files:
                self.files.append(p)
                added += 1
        if added:
            self._log("Added %d file(s)." % added)
        self._refresh_files()

    def _on_drop(self, event):
        self._add_paths(self.tk.splitlist(event.data))

    def _add_files(self):
        conv = self.current
        exts = " ".join("*." + e for e in sorted(conv["src"]))
        paths = filedialog.askopenfilenames(
            title="Choose files", filetypes=[(conv["title"] + " input", exts),
                                             ("All files", "*.*")])
        if paths:
            self._add_paths(paths)

    def _clear_files(self):
        self.files = []
        self._refresh_files()
        self._log("Cleared file list.")

    def _refresh_files(self):
        self.files_box.configure(state="normal")
        self.files_box.delete("1.0", "end")
        for f in self.files:
            self.files_box.insert("end", "  " + os.path.basename(f) + "\n")
        self.files_box.configure(state="disabled")

    def _choose_output(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir = d
            self.out_lbl.configure(text="Saving to: " + os.path.basename(d))
        else:
            self.output_dir = None
            self.out_lbl.configure(text="Saved beside each original")

    # ---------- conversion ----------
    def _start(self):
        if self.working:
            return
        if not self.files:
            messagebox.showinfo(APP_NAME, "Add some files first.")
            return
        out_ext = self.current["out"]
        jobs = []
        for f in self.files:
            src_ext = os.path.splitext(f)[1].lstrip(".").lower()
            fn = route(src_ext, out_ext)
            if fn:
                jobs.append((f, fn, out_ext))
        if not jobs:
            messagebox.showinfo(APP_NAME, "Nothing to convert.")
            return
        self.working = True
        self.convert_btn.configure(state="disabled", text="Converting...")
        self.progress.set(0)
        threading.Thread(target=self._worker, args=(jobs,), daemon=True).start()

    def _worker(self, jobs):
        com = False
        try:
            import pythoncom
            pythoncom.CoInitialize()
            com = True
        except Exception:
            pass
        ok = 0
        total = len(jobs)
        for i, (src, fn, out_ext) in enumerate(jobs, start=1):
            self.q.put(("log", "[%d/%d] %s ..." % (i, total, os.path.basename(src))))
            try:
                base = os.path.splitext(os.path.basename(src))[0]
                folder = self.output_dir or os.path.dirname(src)
                dst = unique_path(os.path.join(folder, base + "." + out_ext))
                fn(src, dst)
                ok += 1
                self.q.put(("log", "    -> " + os.path.basename(dst)))
            except Exception as e:
                self.q.put(("log", "    ERROR: " + str(e)))
                sys.stderr.write(traceback.format_exc() + "\n")
            self.q.put(("progress", i / total))
        if com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        self.q.put(("done", (ok, total)))

    # ---------- updates ----------
    def check_for_updates(self, silent=False):
        threading.Thread(target=self._update_worker, args=(silent,), daemon=True).start()

    def _update_worker(self, silent):
        if not GITHUB_USER or "YOUR_GITHUB" in GITHUB_USER:
            if not silent:
                self.q.put(("info", ("Updates not set up yet",
                            "Upload the files to your GitHub repo first "
                            "(see GITHUB_SETUP.md).")))
            return
        remote = None
        branch_found = None
        last_err = ""
        for branch in GITHUB_BRANCHES:
            try:
                remote = _gh_get(branch, "version.txt", 10).decode("utf-8").strip()
                branch_found = branch
                break
            except Exception as e:
                last_err = str(e)
        if not remote:
            if not silent:
                self.q.put(("info", ("Couldn't check for updates",
                    "Could not read the version from your repo. Make sure the "
                    "files (including version.txt) are uploaded to "
                    "github.com/%s/%s and the repo is Public.\n\n(%s)"
                    % (GITHUB_USER, GITHUB_REPO, last_err))))
            return
        if is_newer(remote, APP_VERSION):
            self.q.put(("update", (remote, branch_found)))
        elif not silent:
            self.q.put(("info", ("You're up to date",
                                 "Installed version: %s\nLatest in your repo: %s"
                                 % (APP_VERSION, remote))))

    def _do_update(self, remote, branch):
        try:
            base = os.path.dirname(os.path.abspath(__file__))
            for fname in UPDATE_FILES:
                data = _gh_get(branch, fname, 25)
                with open(os.path.join(base, fname), "wb") as fh:
                    fh.write(data)
            messagebox.showinfo(APP_NAME,
                                "Updated to version %s.\nPlease close and reopen "
                                "iConvert to finish." % remote)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_NAME, "Update failed: %s" % e)

    # ---------- queue pump ----------
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    self.progress.set(payload)
                elif kind == "info":
                    messagebox.showinfo(payload[0], payload[1])
                elif kind == "update":
                    remote, branch = payload
                    if messagebox.askyesno(
                            APP_NAME,
                            "Version %s is available (you have %s).\nUpdate now?"
                            % (remote, APP_VERSION)):
                        self._do_update(remote, branch)
                elif kind == "done":
                    ok, total = payload
                    self.working = False
                    self.convert_btn.configure(state="normal", text="Convert")
                    self._log("Finished: %d/%d converted." % (ok, total))
                    if ok:
                        where = self.output_dir or "the same folder as each original"
                        messagebox.showinfo(APP_NAME,
                                            "Done! %d of %d file(s) converted.\n\n"
                                            "Saved to: %s" % (ok, total, where))
                    else:
                        messagebox.showwarning(APP_NAME,
                                               "No files were converted. See the log.")
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
