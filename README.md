# iConvert — Local File Converter (v2.0)

A modern, **tile-based** offline converter (inspired by iLovePDF). Pick a tool,
drop your files, convert. Everything runs **on your laptop** — nothing is
uploaded.

## Tools

| Tile | Does |
|------|------|
| PDF to Word | PDF → editable .docx |
| Word to PDF | .doc/.docx → PDF |
| PowerPoint to PDF | .ppt/.pptx → PDF |
| PDF to PowerPoint | PDF pages → .pptx slides |
| JPG to PNG | .jpg → .png |
| PNG to JPG | .png → .jpg |
| Image to PDF | JPG/PNG → PDF |

Word→PDF and PowerPoint→PDF use the **Microsoft Office** already on your PC for
faithful layout. The rest use built-in libraries.

---

## Step 1 — Install Python (one time, ~5 min)

1. Open **https://www.python.org/downloads/** → **Download Python 3**.
   *(Use python.org — not the Microsoft Store version.)*
2. Run the installer. On the **first screen, tick “Add python.exe to PATH.”**
3. Click **Install Now**.

## Step 2 — Install iConvert

1. Unzip this folder to a **short path** like `C:\iConvert` or your Downloads
   folder *(deep folders can trigger a “path too long” error when unzipping)*.
2. Double-click **`INSTALL.bat`**. Wait for **“DONE!”**
3. Launch it: press the **Windows key**, type **iConvert** — or use the Desktop icon.

---

## Updating later — two easy ways (no deleting needed)

- **In the app:** click **“Check for updates.”** Works once your GitHub repo is
  set up — see **GITHUB_SETUP.md**.
- **Or `UPDATE.bat`:** download the new files and double-click it; it copies them
  into the installed app instantly. No GitHub required.

You never have to uninstall first — both methods overwrite in place.

---

## Using it

1. Open iConvert and click a **tool tile** (e.g. *PDF to Word*).
2. **Drag files in**, or click **Select files**.
3. *(Optional)* **Output folder…** — otherwise files save next to the originals.
4. Click **Convert**. A popup confirms when it’s done.

---

## Troubleshooting

- **“Python was not found” / Store opens** — real Python isn’t installed. Do
  Step 1 (python.org, tick *Add to PATH*) and re-run `INSTALL.bat`.
- **Word/PowerPoint conversions fail** — these need Microsoft Office installed
  and working. Open Word/PowerPoint once to confirm.
- **PDF→Word off on complex layouts** — text and simple tables convert well;
  heavily designed PDFs may need light cleanup.
- **“Check for updates” can’t reach the repo** — upload the files to your repo
  (see **GITHUB_SETUP.md**), or just use `UPDATE.bat`.

---

## What’s in this folder

- `file_converter.py` — the app window (modern UI + updater)
- `converters.py` — the conversion engine
- `INSTALL.bat` — installs the app + Start-menu/Desktop shortcuts
- `UPDATE.bat` — one-click update (no GitHub)
- `make_shortcuts.ps1` — helper used by INSTALL.bat
- `UNINSTALL.bat` — removes the app + shortcuts
- `requirements.txt` — Python packages it needs
- `version.txt` — current version (used by the updater)
- `GITHUB_SETUP.md` — how to set up auto-updates via GitHub
- `icon.ico`, `make_icon.py` — app icon (+ the script that made it)
- `.gitignore` — for the GitHub repo

Everything runs locally. Internet is only used for the one-time package install
and the optional update check.
