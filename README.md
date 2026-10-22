# Claudius PDF Editor

**Free for personal and commercial use, source on GitHub** (licence: GNU AGPL v3, see `LICENSE`). PDF editor for Windows 10/11 (x64) - Tkinter + PyMuPDF. Light/dark look, six languages (pl, de, es, it, fr, en),
printing, tabs, text selection, continuous scrolling, redaction, comments, compare, Bates numbers and more.
See `README.txt` for the full feature list and `SECURITY_NOTES.txt` for privacy / security notes.

## Download the installer (built by GitHub)

1. Open the **Actions** tab -> latest green run of "Build Claudius PDF Editor for Windows" -> **Artifacts**
   -> download `Claudius-PDF-Editor-installer` (Claudius-PDF-Editor-Setup-3.7.exe) or `Claudius-PDF-Editor-windows-x64` (portable zip: unzip, start Claudius PDF Editor.exe).
2. Or, for a public download page: create a tag `v3.7` (see below); the installer then appears under **Releases**.

## How to build from scratch on GitHub

1. Create a new repository on github.com (e.g. `claudius`).
2. Upload **all** files of this folder, keeping the structure (including the hidden folder `.github/workflows/`).
   Easiest: `git clone`, copy the files in, then `git add . && git commit -m "Claudius PDF Editor 3.7" && git push`.
3. Push to `main` -> the build starts automatically (Actions tab, about 5 minutes). Also use *Run workflow* to start it by hand.
4. Release: `git tag v3.7 && git push origin v3.7` -> the workflow also creates a GitHub Release with
   `Claudius-PDF-Editor-Setup-3.7.exe`, `Claudius-PDF-Editor-portable-3.7.zip` and `SHA256SUMS.txt`.

## Build on your own PC

* Run from source: `python -m pip install -r requirements.txt` then `python claudius.py` (or `run.bat`).
* EXE: `build_exe.bat` -> `dist\Claudius PDF Editor\Claudius PDF Editor.exe` (a folder; keep it together).  Installer: `build_installer.bat` (needs free Inno Setup 6).

Optional tools used when present: Tesseract (OCR), LibreOffice (Office conversion), Ghostscript (PDF/A).

## Files

| File | Purpose |
|---|---|
| `claudius.py` + `claudius_*.py` (14 modules) | the program - all must stay together |
| `claudius.ico`, `splash.png`, `version_info.txt` | icon, EXE start-up splash picture, Windows version info |
| `FREEWARE.txt`, `LICENSE` | freeware notice shown by the installer; the licence (GNU AGPL v3) |
| `build_exe.bat`, `build_installer.bat`, `installer.iss` | local builds |
| `.github/workflows/build.yml` | automatic build on GitHub |
| `requirements.txt` | Python packages |
