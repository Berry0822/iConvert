# iConvert — Local File Converter

iConvert is a simple desktop app for converting your everyday work files —
PDFs, Word documents, PowerPoint slides, and images — right on your own
computer. It has a clean, tile-based screen: pick a tool, add your files, and
convert. Your files stay on your laptop.

## What it can convert

| Tool | Converts |
|------|----------|
| PDF to Word | PDF → editable Word (.docx) |
| Word to PDF | Word (.doc/.docx) → PDF |
| PowerPoint to PDF | PowerPoint (.ppt/.pptx) → PDF |
| PDF to PowerPoint | PDF pages → PowerPoint slides |
| JPG to PNG | .jpg → .png |
| PNG to JPG | .png → .jpg |
| Image to PDF | JPG/PNG → PDF |

## What you'll need

- A Windows PC
- Microsoft Office (for the Word and PowerPoint conversions)
- Python — a quick one-time install in Step 1

---

## Step 1 — Install Python (one time, ~5 minutes)

1. Go to **https://www.python.org/downloads/** and click **Download Python 3**. (Use python.org, not the Microsoft Store version.)
2. Run the installer. On the **first screen, tick “Add python.exe to PATH.”**
3. Click **Install Now** and let it finish.

You only do this once.

## Step 2 — Install iConvert

1. Unzip the iConvert folder to a short location such as **`C:\iConvert`** or your **Downloads** folder.
2. Double-click **`INSTALL.bat`** and wait until it says **“DONE!”**
3. Open it any time: press the **Windows key** and type **iConvert**, or use the **iConvert** icon on your Desktop.

---

## How to use it

1. Open iConvert and click a **tool** (for example, *PDF to Word*).
2. **Drag your files in**, or click **Select files**.
3. *(Optional)* Click **Output folder…** to choose where the results go — otherwise they’re saved next to your original files.
4. Click **Convert**. A message confirms when it’s done.

---

## Keeping it up to date

When a new version is available you have two easy options, and you never need to uninstall first:

- Click **Check for updates** inside the app *(one-time setup in `GITHUB_SETUP.md`)*, or
- Double-click **UPDATE.bat**.

---

## If something doesn’t work

- **“Python was not found” appears, or the Microsoft Store opens** — Python isn’t installed yet. Do Step 1, then run **INSTALL.bat** again.
- **Word or PowerPoint conversions don’t work** — make sure Microsoft Office is installed and opens normally.
- **PDF to Word looks a little off** — plain text and simple tables convert cleanly; very fancy page layouts may need a small tidy-up afterward.

---

## The files in this folder

- **INSTALL.bat** — installs iConvert (adds the Start-menu and Desktop shortcuts)
- **UPDATE.bat** — updates iConvert to the newest version
- **UNINSTALL.bat** — removes iConvert
- **GITHUB_SETUP.md** — optional: set up the in-app “Check for updates” button

Everything runs on your computer, and your files never leave your laptop.
