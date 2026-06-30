"""
iConvert - Local File Converter (modern UI)
A tile-based, offline converter. Conversion logic lives in converters.py;
this file is the window/UI + the in-app updater.
"""

import os
import sys
import threading
import queue
import traceback
import time
import random
import urllib.request

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converters import route, unique_path, is_newer


def _read_version():
    # Single source of truth: the version lives in version.txt next to this file.
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, "version.txt"), encoding="utf-8") as _f:
            _v = _f.read().strip()
        if _v:
            return _v
    except Exception:
        pass
    return "2.2.0"


APP_NAME = "iConvert"
APP_VERSION = _read_version()

GITHUB_USER = "Berry0822"
GITHUB_REPO = "iConvert"
GITHUB_BRANCHES = ["main", "master"]
UPDATE_FILES = ("converters.py", "file_converter.py", "version.txt")

# palette
ACCENT_RED = "#E5322D"
GREEN = "#16A34A"
RED = "#DC2626"
AMBER = "#D97706"
TEXT = "#111827"
MUTED = "#6B7280"
TRACK = "#E5E7EB"
PAGE = "#F4F6FB"
CARD = "#FFFFFF"


def _raw_url(branch, fname):
    return "https://raw.githubusercontent.com/{}/{}/{}/{}".format(
        GITHUB_USER, GITHUB_REPO, branch, fname)


def _gh_get(branch, path, timeout):
    # Always-fresh read: GitHub API is not served by the raw CDN, so a just-
    # pushed change shows up immediately. Falls back to raw with a cache-buster.
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
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def _darken(hexc, f=0.85):
    hexc = hexc.lstrip("#")
    r, g, b = (int(hexc[i:i+2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % (int(r*f), int(g*f), int(b*f))


class App(BaseTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME + " - Local File Converter")
        self.geometry("900x720")
        self.minsize(820, 640)
        self.configure(fg_color=PAGE)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.files = []
        self.output_dir = None
        self.last_out_dir = None
        self.current = None
        self.working = False
        self.q = queue.Queue()

        # progress animation state
        self._disp = 0.0
        self._target = 0.0
        self._dot = 0
        self._status_base = ""
        self._has_progress = False

        self._build_header()
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True)
        self._build_home()
        self.conv_view = ctk.CTkFrame(self.body, fg_color="transparent")
        self.show_home()

        self.after(120, self._poll)
        self.after(33, self._tick)
        self.after(900, lambda: self.check_for_updates(silent=True))

    # ---------- header ----------
    def _build_header(self):
        h = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=CARD)
        h.pack(fill="x")
        h.pack_propagate(False)
        left = ctk.CTkFrame(h, fg_color="transparent")
        left.pack(side="left", padx=18)
        ctk.CTkLabel(left, text="i", width=34, height=34, corner_radius=9,
                     fg_color=ACCENT_RED, text_color="white",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", pady=14)
        ctk.CTkLabel(left, text="Convert", text_color=TEXT,
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=(6, 0))
        right = ctk.CTkFrame(h, fg_color="transparent")
        right.pack(side="right", padx=16)
        ctk.CTkLabel(right, text="v" + APP_VERSION, text_color="#9CA3AF",
                     font=ctk.CTkFont(size=12)).pack(side="right", padx=(8, 4))
        ctk.CTkButton(right, text="Check for updates", width=140, height=30,
                      corner_radius=8, fg_color="#EEF2FF", text_color="#3730A3",
                      hover_color="#E0E7FF",
                      command=lambda: self.check_for_updates(False)).pack(side="right")

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
        for i, conv in enumerate(CONVERSIONS):
            r, c = divmod(i, COLS)
            self._make_tile(grid, conv).grid(row=r, column=c, padx=12, pady=12)

    def _make_tile(self, parent, conv):
        card = ctk.CTkFrame(parent, width=250, height=118, corner_radius=14,
                            fg_color=CARD, border_width=1, border_color=TRACK)
        card.grid_propagate(False)
        card.pack_propagate(False)
        badge = ctk.CTkLabel(card, text=conv["badge"], width=50, height=50,
                             corner_radius=12, fg_color=conv["color"],
                             text_color="white",
                             font=ctk.CTkFont(size=14, weight="bold"))
        badge.place(x=16, y=16)
        title = ctk.CTkLabel(card, text=conv["title"], text_color=TEXT, anchor="w",
                             font=ctk.CTkFont(size=15, weight="bold"))
        title.place(x=78, y=22)
        sub = ctk.CTkLabel(card, text=conv["sub"], text_color=MUTED, anchor="w",
                           justify="left", wraplength=150,
                           font=ctk.CTkFont(size=11))
        sub.place(x=78, y=48)

        def on_enter(_e): card.configure(border_color=conv["color"], border_width=2)
        def on_leave(_e): card.configure(border_color=TRACK, border_width=1)
        def on_click(_e): self.show_conversion(conv)
        for w in (card, badge, title, sub):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass
        return card

    # ---------- view switch ----------
    def show_home(self):
        self._has_progress = False
        self.conv_view.pack_forget()
        self.home.pack(fill="both", expand=True)

    def _clear(self, frame):
        for ch in frame.winfo_children():
            ch.destroy()

    def show_conversion(self, conv):
        self.current = conv
        self.files = []
        self.output_dir = None
        self.last_out_dir = None
        accent = conv["color"]
        self.home.pack_forget()
        self._clear(self.conv_view)
        self.conv_view.pack(fill="both", expand=True)

        top = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        top.pack(fill="x", padx=22, pady=(14, 4))
        ctk.CTkButton(top, text="< Back", width=70, height=30, corner_radius=8,
                      fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self.show_home).pack(side="left")
        ctk.CTkLabel(top, text="  " + conv["badge"], width=44, height=30,
                     corner_radius=8, fg_color=accent, text_color="white",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(10, 8))
        ctk.CTkLabel(top, text=conv["title"], text_color=TEXT,
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        drop = ctk.CTkFrame(self.conv_view, height=104, corner_radius=14,
                            fg_color=CARD, border_width=2, border_color="#DBE3FF")
        drop.pack(fill="x", padx=22, pady=8)
        drop.pack_propagate(False)
        msg = ("Drag files here, or" if DND_AVAILABLE else "Add the files to convert")
        ctk.CTkLabel(drop, text=msg, text_color=MUTED,
                     font=ctk.CTkFont(size=14)).pack(pady=(20, 6))
        ctk.CTkButton(drop, text="Select files", width=150, height=38, corner_radius=9,
                      fg_color=accent, hover_color=_darken(accent),
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._add_files).pack()
        if DND_AVAILABLE:
            try:
                drop.drop_target_register(DND_FILES)
                drop.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        row = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        row.pack(fill="x", padx=22, pady=(2, 0))
        ctk.CTkLabel(row, text="Files", text_color="#374151",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(row, text="Clear", width=64, height=26, corner_radius=7,
                      fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self._clear_files).pack(side="right")
        self.files_box = ctk.CTkTextbox(self.conv_view, height=64, corner_radius=10,
                                        fg_color=CARD, text_color=TEXT,
                                        border_width=1, border_color=TRACK)
        self.files_box.pack(fill="x", padx=22, pady=(4, 4))
        self.files_box.configure(state="disabled")

        out = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        out.pack(fill="x", padx=22, pady=(2, 2))
        self.out_lbl = ctk.CTkLabel(out, text="Saved beside each original",
                                    text_color=MUTED, font=ctk.CTkFont(size=12))
        self.out_lbl.pack(side="left")
        ctk.CTkButton(out, text="Output folder...", width=130, height=28,
                      corner_radius=8, fg_color="#F3F4F6", text_color="#374151",
                      hover_color="#E5E7EB", command=self._choose_output).pack(side="right")

        self.convert_btn = ctk.CTkButton(
            self.conv_view, text="Convert", height=44, corner_radius=10,
            fg_color=accent, hover_color=_darken(accent),
            font=ctk.CTkFont(size=16, weight="bold"), command=self._start)
        self.convert_btn.pack(fill="x", padx=22, pady=(6, 8))

        # progress block
        prow = ctk.CTkFrame(self.conv_view, fg_color="transparent")
        prow.pack(fill="x", padx=22)
        self.status_lbl = ctk.CTkLabel(prow, text="Ready to convert",
                                       text_color=MUTED, font=ctk.CTkFont(size=13))
        self.status_lbl.pack(side="left")
        self.pct_lbl = ctk.CTkLabel(prow, text="0%", text_color=accent,
                                    font=ctk.CTkFont(size=22, weight="bold"))
        self.pct_lbl.pack(side="right")
        self.progress = ctk.CTkProgressBar(self.conv_view, height=14, corner_radius=7,
                                           fg_color=TRACK, progress_color=accent)
        self.progress.set(0)
        self.progress.pack(fill="x", padx=22, pady=(4, 8))
        self._disp = 0.0
        self._target = 0.0
        self._has_progress = True

        self.results = ctk.CTkScrollableFrame(self.conv_view, fg_color=CARD,
                                              corner_radius=10, label_text="Results",
                                              label_text_color=MUTED)
        self.results.pack(fill="both", expand=True, padx=22, pady=(0, 6))

        self.open_btn = ctk.CTkButton(self.conv_view, text="Open output folder",
                                      height=34, corner_radius=9, fg_color=GREEN,
                                      hover_color=_darken(GREEN),
                                      command=self._open_output)
        # shown only after a successful run
        self.open_btn.pack(fill="x", padx=22, pady=(0, 14))
        self.open_btn.pack_forget()

    # ---------- files ----------
    def _add_paths(self, paths):
        conv = self.current
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if not os.path.isfile(p):
                continue
            ext = os.path.splitext(p)[1].lstrip(".").lower()
            if ext not in conv["src"]:
                continue
            if p not in self.files:
                self.files.append(p)
                added += 1
        self._refresh_files()

    def _on_drop(self, event):
        self._add_paths(self.tk.splitlist(event.data))

    def _add_files(self):
        conv = self.current
        exts = " ".join("*." + e for e in sorted(conv["src"]))
        paths = filedialog.askopenfilenames(
            title="Choose files",
            filetypes=[(conv["title"] + " input", exts), ("All files", "*.*")])
        if paths:
            self._add_paths(paths)

    def _clear_files(self):
        self.files = []
        self._refresh_files()

    def _refresh_files(self):
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
        self._clear(self.results)
        self.open_btn.pack_forget()
        self._disp = 0.0
        self._target = 0.0
        self.progress.set(0)
        self.pct_lbl.configure(text="0%")
        self.working = True
        self.convert_btn.configure(state="disabled", text="Converting...")
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
            self.q.put(("start", (i, total, os.path.basename(src))))
            try:
                base = os.path.splitext(os.path.basename(src))[0]
                folder = self.output_dir or os.path.dirname(src)
                dst = unique_path(os.path.join(folder, base + "." + out_ext))
                fn(src, dst)
                ok += 1
                self.last_out_dir = folder
                self.q.put(("result", (True, os.path.basename(src), os.path.basename(dst))))
            except Exception as e:
                self.q.put(("result", (False, os.path.basename(src), str(e)[:60])))
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
        ctk.CTkLabel(rowf, text="●", width=14,
                     text_color=GREEN if ok else RED,
                     font=ctk.CTkFont(size=14)).pack(side="left")
        if ok:
            txt = left + "   →   " + right
            col = "#374151"
        else:
            txt = left + "  -  failed: " + right
            col = RED
        ctk.CTkLabel(rowf, text=txt, text_color=col, anchor="w",
                     font=ctk.CTkFont(size=12)).pack(side="left", padx=4)

    # ---------- progress animation ----------
    def _tick(self):
        if self._has_progress and self.current:
            try:
                self._disp += (self._target - self._disp) * 0.18
                if self._target - self._disp < 0.002:
                    self._disp = self._target
                self.progress.set(self._disp)
                self.pct_lbl.configure(text="%d%%" % int(round(self._disp * 100)))
                if self.working:
                    self._dot = (self._dot + 1) % 48
                    self.status_lbl.configure(
                        text=self._status_base + "." * (1 + (self._dot // 12)))
            except Exception:
                pass
        self.after(33, self._tick)

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

    def _open_output(self):
        d = self.last_out_dir or self.output_dir
        if d and os.path.isdir(d):
            try:
                os.startfile(d)
            except Exception:
                messagebox.showinfo(APP_NAME, "Files saved in:\n" + d)

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
                    if ok == total:
                        self._status_base = "Done - %d of %d converted" % (ok, total)
                        self.status_lbl.configure(text=self._status_base, text_color=GREEN)
                    else:
                        self._status_base = "Finished - %d of %d converted" % (ok, total)
                        self.status_lbl.configure(text=self._status_base, text_color=AMBER)
                    if ok:
                        self.open_btn.pack(fill="x", padx=22, pady=(0, 14))
                elif kind == "info":
                    messagebox.showinfo(payload[0], payload[1])
                elif kind == "update":
                    remote, branch = payload
                    if messagebox.askyesno(
                            APP_NAME,
                            "Version %s is available (you have %s).\nUpdate now?"
                            % (remote, APP_VERSION)):
                        self._do_update(remote, branch)
        except queue.Empty:
            pass
        self.after(120, self._poll)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
