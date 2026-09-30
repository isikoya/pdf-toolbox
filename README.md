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
| Compress | Six levels by image resolution. Shows what your scans are, estimates the result, reports before and after. |
| Convert | Office to PDF, and PDF to Word, Excel tables, text or images. |
| Bank statement to Excel | Downloaded or scanned statements (FNB, Standard Bank, Absa, Nedbank, Capitec and similar). Every row is checked against the running balance with live formulas; rows that do not add up are red. |
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
- Bank statement workbooks: a row that adds up has been confirmed by the balance, but dates are not covered by that check. Check red rows and amber cells against the PDF, and tie the closing balance on the Summary sheet.
- OCR is not perfect. Spot-check figures on anything that matters.
- Every task writes a new file. Your originals are left alone.

## What it runs on

OCRmyPDF, Tesseract and Ghostscript do the OCR and compression. pypdf,
pdfplumber, openpyxl and Pillow handle the page and table work. Microsoft
Word, Excel and PowerPoint do the Office conversions. Everything except Office
is open source.

## Updates

Open **About this tool** and press **Check for a new version**. If there is a
newer release, press **Download and install it**: the toolbox downloads the
release, checks it against the published checksum, saves your current version
alongside the folder, then closes and reopens on the new one.

It never updates on its own. If the download fails any check, nothing on the
machine is touched and you are pointed at the download page instead.

## Publishing a new version (for the maintainer)

1. Change `$script:Version` near the top of `PDFToolbox.ps1`.
2. Run `Build Release.ps1`. It writes `dist\PDF Toolbox v<version>.zip` and
   regenerates `version.json`.
3. Commit and push `version.json`.
4. Create a Release, tag it exactly `v<version>`, and attach the zip under the
   name the build script prints.

Step 3 is what makes the update check notice. Step 4 is what people download.
The tag and file name must match what `version.json` records, or in-app
updating will 404 and fall back to opening the page.

## Problems

Open **About this tool**, run **Check setup** first: most faults are a missing
install. If that does not sort it, press **Copy details for an email** on the
About page, then email the address there and paste it in. That tells the
maintainer your version, what is installed and what was on screen.
