# iConvert — Local File Converter

iConvert is a simple desktop app for converting and tidying your everyday work
files — PDFs, Word, PowerPoint, Excel, and images — right on your own computer.
Pick a tool, add your files (or a whole folder), and go. Your files stay on your
laptop.

## Tools

**Convert**

- PDF to Word, Word to PDF
- PowerPoint to PDF, PDF to PowerPoint
- Excel to PDF, Excel to CSV
- JPG to PNG, PNG to JPG, Image to PDF
- PDF to JPG, PDF to PNG (one image per page)
- Combine images into one PDF, resize/shrink images

**PDF tools**

- Merge PDFs, Split PDF (pull out pages)
- Rotate pages, delete pages, add page numbers, add a watermark, extract images
- Compress PDF (shrink the file size)
- Protect PDF (add a password), Unlock PDF (remove a password)
- OCR PDF — make a scanned PDF searchable, or turn it into Word

**Handy extras**

- Add a **whole folder** at once (it finds every matching file inside)
- **Dark mode** switch (top-right)
- **Recent** list and an **Open output folder** button
- Optional **right-click menu** in File Explorer

## What you'll need

- A Windows PC
- Microsoft Office (for the Word / PowerPoint / Excel conversions)
- Python — a quick one-time install in Step 1
- For OCR only: the free **Tesseract** program (Step 3, optional)

---

## Step 1 — Install Python (one time, ~5 minutes)

1. Go to **https://www.python.org/downloads/** and click **Download Python 3**. (Use python.org, not the Microsoft Store version.)
2. Run the installer. On the **first screen, tick “Add python.exe to PATH.”**
3. Click **Install Now**.

## Step 2 — Install iConvert

1. Unzip the iConvert folder to a short location such as **`C:\iConvert`**.
2. Double-click **`INSTALL.bat`** and wait until it says **“DONE!”**
3. Open it: press the **Windows key** and type **iConvert**, or use the Desktop icon.

## Step 3 — (Optional) Enable OCR

OCR turns scanned/image PDFs into searchable PDFs or Word. It needs the free
Tesseract program:

1. Download from **https://github.com/UB-Mannheim/tesseract/wiki** and install it (accept the defaults).
2. That's it — iConvert will detect it. If it's missing, the OCR tools will tell you.

---

## How to use it

1. Open iConvert and click a **tool tile**.
2. **Drag files in**, click **Select files**, or **Select folder** to add everything inside.
3. Some tools ask for a little extra: *Split* wants the pages (e.g. `1-3,5`), *Protect/Unlock* want a password.
4. *(Optional)* **Output folder…** — otherwise results save next to your originals.
5. Click **Convert**. A progress bar and pop-up notifications show how it's going.

## Right-click menu (optional)

Want “Convert with iConvert” when you right-click a file? After installing,
double-click **`Add right-click menu.bat`** once. To remove it, use
**`Remove right-click menu.bat`**.

## Updating later

- Click **Check for updates** in the app *(setup in `GITHUB_SETUP.md`)*, or
- Double-click **UPDATE.bat**. You never need to uninstall first.

---

## If something doesn’t work

- **“Python was not found” / the Microsoft Store opens** — Python isn’t installed. Do Step 1, then run **INSTALL.bat** again.
- **OCR says Tesseract is missing** — do Step 3.
- **Word / PowerPoint / Excel conversions fail** — make sure Microsoft Office is installed and opens normally.
- **PDF to Word looks a little off** — plain text and simple tables convert cleanly; very fancy layouts may need a small tidy-up.

## Make a shareable .exe (optional)

Want to give iConvert to someone who doesn't have Python? Double-click **`BUILD_EXE.bat`** once — it produces `dist\\iConvert.exe`, a single file that runs on any Windows PC with no Python needed. (OCR still needs Tesseract.) That .exe is a snapshot; the auto-update feature is for the normal Python install.

Everything runs on your computer, and your files never leave your laptop.
