# PDF Toolbox

A Windows tool for the PDF work that comes up in an audit: making scanned bank
statements searchable, merging and splitting files, pulling pages out, and
converting to and from PDF. Everything runs on your own machine. No file is
uploaded anywhere.

Built by Ismail.

## Download

Go to the [Releases page](../../releases/latest) and download
`PDF Toolbox v<version>.zip`.

## Install

1. Unzip the whole folder somewhere you can keep it, for example
   `Documents\PDF Toolbox`. Keep all the files together.
2. Windows marks files downloaded from the internet as blocked. Right-click
   the zip **before** extracting, choose **Properties**, tick **Unblock**, then
   extract. If you already extracted it, do the same on each `.ps1` file.
3. Double-click **Install PDFToolbox.bat**. It checks what is already on the
   machine and installs only what is missing, then offers to put a shortcut on
   your desktop.
4. Start it from the desktop shortcut, or **Launch PDF Toolbox Window.bat**.

The installer needs no admin rights for the Python parts. Windows may ask for
permission when installing Tesseract or Ghostscript.

## What it does

| Task | Notes |
| --- | --- |
| OCR | Makes scanned PDFs searchable. Runs several files at once and writes a log. |
| Merge | Combine PDFs, in an order you set by dragging. |
| Split, extract, delete, rotate, insert, rearrange | Page-level work, with thumbnails to pick from. |
| Compress | Three levels, with before and after sizes. |
| Convert | Office to PDF, and PDF to Word, Excel tables, text or images. |
| Search | Find text across a folder of PDFs. |
| Count pages | Totals across a folder, with an optional CSV. |
| Remove a password | Needs the password. Writes an unlocked copy. |

There are two versions in the zip. The window version
(`Launch PDF Toolbox Window.bat`) is the one to use. The console version
(`Launch PDFToolbox.bat`) does the same jobs from a text menu and is a fallback
if the window will not open.

## Before you trust the output

- A scanned PDF has no text until you OCR it. Text extraction, table
  extraction and search all return nothing until then.
- PDF to Excel guesses where the columns are. Tie the totals back to the PDF
  before using any figure.
- PDF to Word rebuilds the layout, it does not copy it.
- OCR is not perfect. Spot-check figures on anything that matters.
- Every task writes a new file. Your originals are left alone.

## What it runs on

OCRmyPDF, Tesseract and Ghostscript do the OCR and compression. pypdf,
pdfplumber, openpyxl and Pillow handle the page and table work. Microsoft
Word, Excel and PowerPoint do the Office conversions. Everything except Office
is open source.

## Updates

Open **About this tool** and press **Check for a new version**. It reads
`version.json` from this repository and tells you if there is a newer release.
It never installs anything by itself.

## Publishing a new version (for the maintainer)

1. Change `$script:Version` near the top of `PDFToolbox.ps1`.
2. Run `Build Release.ps1`. It writes `dist\PDF Toolbox v<version>.zip` and
   regenerates `version.json`.
3. Commit and push `version.json`.
4. Create a Release, tag it `v<version>`, and attach the zip.

Step 3 is what makes the update check notice. Step 4 is what people download.

## Problems

Open **About this tool**, run **Check setup** first: most faults are a missing
install. If that does not sort it, email the address on the About page and say
which version you are on and what you were doing.
