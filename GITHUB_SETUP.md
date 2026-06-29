# Auto-updates via your GitHub repo

Your repo is ready: **https://github.com/Berry0822/iConvert** (Public, branch `main`),
and the app is already pointed at it. You just need to put the files **into** the repo.

## One-time: upload the files (drag & drop, ~2 minutes)

1. Open **https://github.com/Berry0822/iConvert** in your browser and sign in.
2. Click **Add file → Upload files**.
3. Open your unzipped **iConvert** folder, select **all** the files, and drag them
   onto the page (file_converter.py, converters.py, version.txt, INSTALL.bat, etc.).
4. Scroll down and click **Commit changes**.

Done. Now open iConvert and click **Check for updates** — it should say
**“You’re up to date.”** That confirms the app is talking to your repo. ✔

## See an update actually happen (optional 30-second test)

1. In the repo, click **version.txt → the pencil (Edit)**.
2. Change `2.0.0` to `2.0.1` and click **Commit changes**.
3. In iConvert, click **Check for updates** → it offers the update, downloads it,
   and asks you to reopen. That’s the full auto-update working end to end.

## Releasing a new version later

When I send you updated files:
1. Repo → **Add file → Upload files** → drag the changed files in
   (same filenames overwrite the old ones).
2. Make sure **version.txt** has a higher number than before (e.g. `2.1.0`).
3. **Commit changes.**

Any iConvert on any PC will then offer the update on **Check for updates**.

## Notes
- Keep the repo **Public** so the app can read updates without a login (yours is).
- The in-app updater refreshes the app code (file_converter.py, converters.py,
  version.txt). If a future version needs new Python packages, I’ll tell you to
  run **INSTALL.bat** once.

## No-GitHub option (still there)
Don’t want to touch GitHub for a given update? Download the new files and
double-click **UPDATE.bat** — it copies them straight into the installed app.
