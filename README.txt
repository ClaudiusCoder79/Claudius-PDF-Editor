Claudius PDF Editor  -  PDF editor for Windows 10 / 11 x64 (also runs on macOS / Linux)
Free for personal and commercial use, source on GitHub: https://github.com/ClaudiusCoder79/Claudius-PDF  (see FREEWARE.txt)
Created by Michal Czapla & Claude AI

RUN
1. Install Python 3.9+ (64-bit) from python.org - tick "Add python.exe to PATH".
2. Double-click run.bat  (installs the needed packages on first run).
   Keep all files of the folder together (claudius.py and the claudius_*.py modules, requirements.txt, claudius.ico, splash.png).

INSTALLER (Claudius-PDF-Editor-Setup-3.0.exe)
  Double-click build_installer.bat (needs Inno Setup, free: jrsoftware.org/isdl.php) -> installer\Claudius-PDF-Editor-Setup-3.0.exe
  or let GitHub build it: the workflow in claudius_github_repo.zip makes both the .exe and the installer.
  The installer asks for administrator rights and installs into C:\Program Files\Claudius PDF Editor.
  It adds a Start menu folder (program + Uninstall), an optional desktop icon, an optional "Open with" entry for PDFs,
  and an uninstaller. If an older CLAUDIUS version is installed, uninstall it first (Settings > Apps).

STANDALONE .EXE
  Double-click build_exe.bat -> dist\Claudius PDF Editor.exe
  (or upload claudius_github_repo.zip's contents to GitHub: the workflow builds the .exe for you)

WINDOWS 10 / 11
  - Sharp on high-DPI screens (125 %, 150 % ...); own icon and grouping in the taskbar.
  - Windows 11: the title bar takes the Claudius PDF Editor blue.
  - Touchpad scrolling is smooth (small movements add up instead of jumping).
  - Drag PDFs onto the window to open them (several files = one combined PDF); drag photos onto
    an open PDF to place them on the page. Needs the tkinterdnd2 package (run.bat installs it).
  - The window remembers its size, position and maximised state.
  - Saving writes a temporary file first, so a failed save (file open elsewhere, OneDrive hiccup)
    can never damage the original; a clear message explains what happened.
  - File menu: "Add Claudius PDF Editor to the 'Open with' list" (your account only) and
    "Choose default PDF app" (opens Windows Settings).
  - Fonts installed "for me only" (Windows 11 default) are found for signatures / watermarks.
  - The .exe carries version information (right-click > Properties > Details).

OPTIONAL PROGRAMS (the editor works without them)
  Tesseract OCR   -> Document > OCR, and "Extract text" on scanned pages
                     (Windows installer: github.com/UB-Mannheim/tesseract/wiki)
  LibreOffice or Microsoft Office -> opening / converting Word, Excel, PowerPoint files to PDF
  Word / Excel / PowerPoint export needs python-docx, openpyxl, python-pptx (run.bat installs them)

TABS
  Home      open / save / print, undo / redo, find, edit text, zoom, rotate
  Comment   highlight (4 colours), underline, strikeout, squiggly, pen, shapes, sticky notes,
            text box, stamps
  Shapes    lines, arrows, rectangle, circle, triangle, diamond, star, check / cross / plus
  Edit      EDIT EXISTING TEXT (click a line; Shift+click = whole paragraph), add text, text box,
            images, links, date, whiteout, REDACT (true removal), Redact text everywhere,
            Extract text (drag an area -> copy; OCR on scans)
  Sign      signature & initials (type / draw / picture, saved for next time), stamps,
            date, fillable form fields; fill any form by clicking its fields with the Hand tool
  Photos    place, move, resize, rotate, crop, adjust photos; Photos -> Pages
            (Adjust tab: "Keep aspect" on = proportional resize; off = stretch width and height
            separately. Holding Shift while dragging a corner temporarily inverts the option.)
  Adjust    photo effects
  Organize  rotate, delete, move, extract, split, insert PDF, merge PDFs, blank page,
            crop pages (drag an area or by margins), photos -> pages
  Document  watermark (text / picture), background colour, header & footer, page numbers,
            compress (batch), flatten, OCR, password protect / remove, properties, print
  Convert   PDF -> Word, Excel, PowerPoint, images (PNG/JPG/TIFF/BMP), text, HTML, XML
            (single or batch); File > Create PDF from files (images, text, Word, Excel, PowerPoint)
  View      zoom, fit, side panel (pages / bookmarks), find, show form fields

PAGES PANEL (left)
  Click = go to page, Ctrl+click / Shift+click = select several, drag to reorder,
  right-click for rotate / extract / delete / insert. Bookmarks tab: add, rename, delete, click to jump.
  Hide the side panel with the little "◀" at its top right, Ctrl+B, or View > Side panel > Hide / show; clicking
  Pages / Bookmarks / Comments brings it back.

SHORTCUTS
  Ctrl+O open, Ctrl+S save, Ctrl+Shift+S save as, Ctrl+P print, Ctrl+F find (Enter / Shift+Enter =
  next / previous), Ctrl+Z undo, Ctrl+Y redo, Ctrl+A select all pages, Ctrl+wheel zoom,
  Ctrl+V paste image, Delete removes the selected photo / pages, Esc closes find, Ctrl+C copy selected text,
  Ctrl+W close tab, Ctrl+Tab next tab, F11 full screen, F5 presentation.

SCROLLING AND VIEW
  Mouse wheel scrolls the page area and the page list; at the end of a page it carries on to the next
  page. Every PDF opens in Fit page; Fit width / manual zoom are on the Home and View tabs.

PHOTOS
  Add Image / Paste: drag to move, corner squares resize, blue dot rotates (Shift = 15 degree steps),
  arrow keys nudge, Ctrl+D duplicate, Enter applies the photo to the page.
  Output quality: Original / High 300 dpi / Medium 200 dpi / Small 120 dpi.

GOOD TO KNOW
  - Edit text re-writes a line in a standard font (Arial-like / Times-like / Courier-like, bold and italic
    kept); the original embedded font is not reused. Text that is slanted cannot be edited.
  - Whiteout only covers content. Use Redact when information must really disappear.
  - A password you type to open a protected PDF is kept when you save (use Document > Remove password to drop it).
  - Print sends the PDF to the default printer through Windows.

LANGUAGES
---------
Claudius PDF Editor speaks English, Polski, Deutsch, Español, Italiano and Français.
The language follows your Windows language the first time you start it.
To change it: File > Language. The program restarts in the new language.
The translations were written by Claude and have not been checked by native speakers;
corrections are welcome - the whole text table is in claudius_lang.py.
Text that goes INTO the PDF (header/footer, page-number formats) stays as you type it.

APPEARANCE
----------
File > Appearance: Automatic (follows the Windows light / dark app mode), Light or Dark.
The program restarts to apply the choice. The pages of your PDF always stay white.
Windows' own message boxes and file pickers keep the system look.

PRINTING
--------
File > Print (Ctrl+P) or the printer button in the quick toolbar: choose the printer, number of copies,
pages and scale. On Windows the pages are sent straight to the printer - no other PDF program is needed.

NEW IN 3.0
----------
SETTINGS               File > Settings (Ctrl+,): language, appearance, what to remember, view of a new document, autosave
                       interval, how long safety copies are kept, default printer and number of copies, highlight colours.
REMEMBERED STATE       The window size and position, the side panel (shown / hidden, which tab) and the page + zoom where you
                       stopped in each file are remembered (switch each off in Settings).
UNDO HISTORY           File > Undo history lists every change by name - click one to go back to the state before it.
CROP AND RESIZE        Tools > Crop / resize: drag the frame (or type margins in mm) to crop; or give pages a new size (A4, A5,
                       Letter ... or your own mm) - the content is fitted, links move with it.
PRESETS                Header/Footer, Watermark and Page numbers dialogs have Load / Save / Delete for your own presets.
EDIT TEXT              Edit text now keeps the original font: the font inside the PDF when it has the letters, else an installed
                       font of the same family (Arial, Times, Calibri, Verdana ...), else the nearest standard font.
EDIT PICTURES          Tools > Edit pictures: click a picture to replace, delete or save it, drag it to move it.
REMOVE A PASSWORD     Open the PDF (enter the password), then Sign/Protect tab > Unprotect: you choose a name and a new file is
                       saved without any password. The original stays as it was. You need to know the password.
DELETE A PICTURE      Right-click it on the page > Delete this picture (a picture you have just added but not applied yet:
                       click it and press Delete). Undo (Ctrl+Z) brings it back.
ALIGN PICTURES        Photos > Align (or right-click a picture > Align to page): centre horizontally / vertically / both,
                       or align to the left, right, top or bottom of the page. Works on new and existing pictures.
SIGNATURE BLOCKS       Tools > Signature block: signature + line + printed name + position + today's date as one stamp.
                       Saved blocks are listed under Sign > Signature.
FORMS                  Tools > Drop-down and Radio button add new field types. Field editor: rename fields, change the choices of
                       a drop-down, delete fields, set the Tab-key order. Fill from table: one PDF per row of a CSV or Excel file.
BATCH                  Tools > Batch folder: compress, watermark, rotate, OCR, remove hidden data, number pages or convert every
                       PDF of a folder. Originals are never changed - results go into a separate folder.
SMART COMPRESS         Tools > Smart compress: shows the resulting size of every quality level and a before / after preview.
MERGE + PICK           Tools > Merge + pick: join several files and choose the pages of each from a picture grid.
SEARCH FOLDER          Tools > Search folder: find a word in all PDFs of a folder; double-click a hit to open it.
HIGHLIGHTS             Give each highlight colour a meaning (Settings), then Review > Export highlights writes them to Word,
                       text or Markdown, grouped by meaning.
LOOK UP / TRANSLATE    Select text, right-click > Look up / translate: opens your browser at Google Translate, DeepL, Wiktionary
                       or a web search. Only the text you look up is sent, and only when you press a button.
REDACTION CHECK        Review > Redaction check: finds unapplied redaction marks, text that is still there, hidden / white text,
                       and words in comments, form fields, bookmarks, attachments, document info and XMP data.
CLEAN DATA             Review > Clean data: remove document info, XMP, scripts, thumbnails, hidden text, attachments, comments.
SAFETY COPY            Autosave can be switched off or set to 1-30 minutes (Settings).

LICENCE
-------
Claudius PDF Editor is free for personal and commercial use and open source: GNU Affero General Public License v3
(file LICENSE), because it is built on PyMuPDF / MuPDF (AGPL). Source: https://github.com/ClaudiusCoder79/Claudius-PDF

NEW IN 2.1
----------
SELECT AND COPY TEXT   Hand tool: drag over text (double-click = word, triple-click = line), Ctrl+C copies.
                       Right-click the selection: Copy, Highlight, Underline, Strikethrough, Redact.
CONTINUOUS SCROLLING   All pages one below the other (View > Continuous switches back to one page at a time).
NIGHT VIEW             View > Night view inverts the page colours on screen only; the PDF is not changed.
FULL SCREEN            F11 hides all bars. Presentation (F5): black background, one page at a time,
                       click / Space / arrow keys = next, right-click / Backspace = previous, Esc leaves.
TABS                   Every document opens in its own tab (Ctrl+Tab next tab, Ctrl+W close tab). Each tab keeps its own
                       undo history, zoom and position. Results (comparison, scans) open in their own tab.
CRASH RECOVERY         Unsaved changes are written to ~/.claudius/recovery every minute. After a crash Claudius PDF Editor offers to
                       bring the documents back. Password-protected documents are never written there.
SIGNED PDFs            Claudius PDF Editor can no longer create or check certificate signatures. If you open a PDF that is already
                       digitally signed, it warns you that editing and saving will invalidate that signature.
COMMENTS               Side panel "Comments" lists every note / markup; click to jump, Delete removes. Export / Import
                       save and load the annotations (or take them from another PDF); Summary writes a readable list.
COMPARE TWO PDFS       Document > Compare: opens a report in a new tab - red = only in the open file, green = only in
                       the other one - plus a list of changed words per page.
PAGE LABELS            Document > Page labels: i, ii, iii / 1, 2, 3 / A-1 ... shown next to the page number.
ATTACHMENTS            Document > Attachments: attach files to the PDF, save or remove them.
BATES NUMBERING        Document > Bates: prefix + running number on every page, or across several PDFs at once.
ACCESSIBILITY CHECK    Document > Accessibility: checks title, language, text, bookmarks, tags and fixes title / language.
                       Claudius PDF Editor cannot create tagged PDFs; for those export from Word / Writer with "tagged PDF" ticked.
PRINT PREVIEW          "Preview..." in the Print dialog.
PDF/A                  Document > PDF/A (needs the free program Ghostscript; Claudius PDF Editor offers the download page).
                       Check important archive files with a validator such as veraPDF.
SCANNER                Convert > From scanner (Windows): uses the Windows scanner dialog, builds a PDF in a new tab.
READ ALOUD             View > Read aloud: the selection, this page or the rest of the document, using the Windows voices.
REDACT DATA            Edit > Redact data: e-mail addresses, phone numbers, IBAN, card, PESEL / NIP, SSN, dates or your
                       own pattern - found and listed first, removed only after you confirm. Numbers are checksum-tested.

NOT TESTED ON REAL HARDWARE
  The scanner, printing, read aloud, default-app setup and the Windows title bar colours can only run on Windows;
  they were tested here with stand-ins. PDF/A was tested with a stand-in for Ghostscript. If one of them does not
  work on your PC, tell me what the message says.

VERSION 3.0  (2.1 added tabs, selection, recovery ...; 2.0 brought the iOS look, dark mode, six languages, printing)
-----------
All the above, plus the changes listed in 2.0: new look (iOS style, light and dark), 3-D icon buttons,
quick toolbar, six languages, "Make Claudius PDF Editor the default PDF app", print dialog.
