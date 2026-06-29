# Auto-updates via your GitHub repo

Your repo: **https://github.com/Berry0822/iConvert** (Public, branch `main`).
The app reads new versions through the GitHub API, so changes show up immediately.
The version number lives in **version.txt** (one source of truth).

## IMPORTANT — what to upload for a new version

When I send you a new build, upload **all** the files and let them **overwrite**
the ones in the repo — especially **file_converter.py** and **converters.py**,
not just version.txt. (If the repo keeps an old file_converter.py, every update
would pull that old program back. That was the earlier problem.)

For a *quick test* you can edit just version.txt — that's fine **as long as the
repo already holds the latest program files.**

## Upload steps (~2 min)

1. Open **https://github.com/Berry0822/iConvert** and sign in.
2. **Add file -> Upload files**, drag in **every** file from the new folder,
   then **Commit changes**. Confirm `file_converter.py` is in the list.

## Update your installed copy

- Double-click **UPDATE.bat** from the new folder, **or**
- In the app, click **Check for updates** (it downloads + applies when the repo
  is newer).

When app and repo match, **Check for updates** shows both numbers:
> Installed version: 2.1.2
> Latest in your repo: 2.1.2

## See an update happen live (30-second test)

1. In the repo: **version.txt -> pencil (Edit) -> change to `2.1.3` -> Commit**.
2. In iConvert: **Check for updates** -> it offers 2.1.3, downloads, asks to reopen.
3. Reopen -> it now shows **2.1.3** and stays there. (No more sliding back.)

## Notes
- Keep the repo **Public**. Keep `version.txt` to just the number.
- If a future build needs new Python packages, run **INSTALL.bat** once.
