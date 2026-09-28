#Requires -Version 5.1
<#
    PDF TOOLBOX
    Built by Ismail (toolkitpdf@gmail.com).

    OCR, merge, split, extract, delete, rotate, compress, search,
    page counts, password removal, and conversion both to and from PDF.

    HOW TO START
      Run "Install PDFToolbox.bat" once, then double-click
      "Launch PDFToolbox.bat". Keep all four files in the same folder.

    NAVIGATION
      At any question:  P = go back one step,  M = main menu,  Q = cancel.
      When asked for a file or folder you can paste the path, drag the
      file into the window, or type B to browse.
#>

$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'PDF Toolbox' } catch { }

$script:Version     = '1.0'
$script:Author      = 'Ismail'
$script:AuthorEmail = 'toolkitpdf@gmail.com'
$script:ReleaseDate = '23/09/2026'
$script:Stamp       = "PDF Toolbox v$script:Version ($script:Author)"

# ----------------------------------------------------------------------
# WHERE UPDATES COME FROM
# Put version.json in the repository root and publish each release on the
# Releases page. Change GITHUBUSER and REPONAME to your own, in both lines.
# ----------------------------------------------------------------------
$script:UpdateManifestUrl = 'https://raw.githubusercontent.com/isikoya/pdf-toolbox/main/version.json'
$script:UpdateDownloadUrl = 'https://github.com/isikoya/pdf-toolbox/releases/latest'

$script:ConfigDir  = Join-Path $env:APPDATA 'PDFToolbox'
$script:ConfigFile = Join-Path $script:ConfigDir 'settings.json'
$script:WorkDir    = Join-Path $env:TEMP 'PDFToolbox'
$script:Helper     = Join-Path $script:WorkDir 'pdftoolboxhelper.py'
$script:LastFolder = $null
$script:PyExe      = $null
$script:PickerType = $null

$script:ImageExtensions = @('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif')
$script:OfficeExtensions = @('.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx')

# ======================================================================
# PYTHON HELPER
# Written to %TEMP% at start-up. Paths are passed as arguments, never
# pasted into the code, so spaces and symbols in paths are safe.
# ======================================================================
$script:HelperCode = @'
import os
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: The pypdf package is not installed. Run Install PDFToolbox.bat.")
    sys.exit(10)


def fail(msg, code=4):
    print("ERROR: " + msg)
    sys.exit(code)


def password():
    return os.environ.get("PDFTB_PASSWORD", "")


def open_reader(path):
    path = Path(path)
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        fail(f"Cannot open {path.name}: {exc}")
    if reader.is_encrypted:
        try:
            result = reader.decrypt(password())
        except Exception as exc:
            fail(f"Cannot decrypt {path.name}: {exc}", 3)
        if not result:
            fail(f"{path.name} is password protected and the password is wrong or missing.", 3)
    return reader


def read_list(list_file):
    with open(list_file, "r", encoding="utf-8-sig") as fh:
        return [line.strip() for line in fh if line.strip()]


def parse_spec(spec, page_count):
    groups = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            bits = part.split("-")
            if len(bits) != 2 or not bits[0].isdigit() or not bits[1].isdigit():
                fail(f"'{part}' is not a valid range. Use a format like 5-10.", 2)
            start, end = int(bits[0]), int(bits[1])
        else:
            if not part.isdigit():
                fail(f"'{part}' is not a valid page number.", 2)
            start = end = int(part)
        if start < 1 or end > page_count or start > end:
            fail(f"'{part}' is outside the document (pages 1 to {page_count}).", 2)
        groups.append((start, end))
    if not groups:
        fail("No pages were entered.", 2)
    return groups


def pages_from_spec(spec, page_count):
    if spec.strip().lower() == "all":
        return set(range(1, page_count + 1))
    wanted = set()
    for start, end in parse_spec(spec, page_count):
        wanted.update(range(start, end + 1))
    return wanted


def write_pdf(writer, out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "wb") as fh:
        writer.write(fh)
    print(f"Created: {out_path.name}")


def label(start, end):
    return f"p{start:03}" if start == end else f"p{start:03}-{end:03}"


def cmd_info(path):
    path = Path(path)
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        fail(f"Cannot open {path.name}: {exc}")
    encrypted = reader.is_encrypted
    pages = -1
    if encrypted:
        try:
            if reader.decrypt(password()):
                pages = len(reader.pages)
        except Exception:
            pages = -1
    else:
        pages = len(reader.pages)
    print(f"PAGES={pages}")
    print(f"ENCRYPTED={1 if encrypted else 0}")


def cmd_merge(out_path, list_file):
    writer = PdfWriter()
    total = 0
    for f in read_list(list_file):
        reader = open_reader(f)
        for page in reader.pages:
            writer.add_page(page)
        total += len(reader.pages)
        print(f"Added: {Path(f).name} ({len(reader.pages)} pages)")
    write_pdf(writer, out_path)
    print(f"Total pages: {total}")


def cmd_split(in_path, out_dir):
    reader = open_reader(in_path)
    stem = Path(in_path).stem
    for i, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        write_pdf(writer, Path(out_dir) / f"{stem} {label(i, i)}.pdf")
    print(f"Split into {len(reader.pages)} files.")


def cmd_extract(in_path, out_path, spec):
    reader = open_reader(in_path)
    groups = parse_spec(spec, len(reader.pages))
    writer = PdfWriter()
    count = 0
    for start, end in groups:
        for n in range(start, end + 1):
            writer.add_page(reader.pages[n - 1])
            count += 1
    write_pdf(writer, out_path)
    print(f"Extracted {count} pages.")


def cmd_ranges(in_path, out_dir, spec):
    reader = open_reader(in_path)
    groups = parse_spec(spec, len(reader.pages))
    stem = Path(in_path).stem
    for start, end in groups:
        writer = PdfWriter()
        for n in range(start, end + 1):
            writer.add_page(reader.pages[n - 1])
        write_pdf(writer, Path(out_dir) / f"{stem} {label(start, end)}.pdf")
    print(f"Created {len(groups)} files.")


def cmd_delete(in_path, out_path, spec):
    reader = open_reader(in_path)
    total = len(reader.pages)
    drop = pages_from_spec(spec, total)
    keep = [i for i in range(1, total + 1) if i not in drop]
    if not keep:
        fail("That would delete every page.", 2)
    writer = PdfWriter()
    for i in keep:
        writer.add_page(reader.pages[i - 1])
    write_pdf(writer, out_path)
    print(f"Removed {len(drop)} page(s). Kept {len(keep)} of {total}.")


def cmd_rotate(in_path, out_path, spec, degrees):
    try:
        deg = int(degrees)
    except ValueError:
        fail("Rotation must be 90, 180 or 270.", 2)
    reader = open_reader(in_path)
    total = len(reader.pages)
    targets = pages_from_spec(spec, total)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        if i in targets:
            page.rotate(deg)
        writer.add_page(page)
    write_pdf(writer, out_path)
    print(f"Rotated {len(targets)} of {total} page(s) by {deg} degrees.")


def image_as_page(image_path, page_size, dpi=300):
    """Turns an image into one PDF page the same size as the target document,
    so the inserted page prints like every other page."""
    try:
        from PIL import Image
    except ImportError:
        fail("Pillow is not installed. Run Install PDFToolbox.bat.", 10)
    import io
    img = Image.open(image_path)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    width_pt, height_pt = page_size
    canvas_w = max(1, int(round(width_pt / 72.0 * dpi)))
    canvas_h = max(1, int(round(height_pt / 72.0 * dpi)))
    scale = min(canvas_w / img.width, canvas_h / img.height)
    new_size = (max(1, int(img.width * scale)), max(1, int(img.height * scale)))
    img = img.resize(new_size, Image.LANCZOS)
    canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
    canvas.paste(img, ((canvas_w - new_size[0]) // 2, (canvas_h - new_size[1]) // 2))
    buffer = io.BytesIO()
    canvas.save(buffer, "PDF", resolution=float(dpi))
    buffer.seek(0)
    return list(PdfReader(buffer).pages)


def load_insert_pages(insert_path, spec, page_size):
    """Returns the pages to insert. Images become one page first."""
    suffix = Path(insert_path).suffix.lower()
    if suffix in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"):
        return image_as_page(insert_path, page_size), 1
    reader = open_reader(insert_path)
    total = len(reader.pages)
    if spec.strip().lower() == "all":
        return list(reader.pages), total
    wanted = sorted(pages_from_spec(spec, total))
    return [reader.pages[i - 1] for i in wanted], total


def cmd_insert(base_path, out_path, insert_path, spec, position):
    base = open_reader(base_path)
    total = len(base.pages)
    first = base.pages[0].mediabox
    page_size = (float(first.width), float(first.height))
    incoming, source_total = load_insert_pages(insert_path, spec, page_size)
    if not incoming:
        fail("Nothing to insert.", 2)
    place = position.strip().lower()
    if place == "start":
        at = 0
    elif place == "end":
        at = total
    elif place.isdigit():
        at = int(place) - 1
        if at < 0 or at > total:
            fail(f"Position must be between 1 and {total + 1}.", 2)
    else:
        fail("Position must be a page number, or start, or end.", 2)
    writer = PdfWriter()
    for page in base.pages[:at]:
        writer.add_page(page)
    for page in incoming:
        writer.add_page(page)
    for page in base.pages[at:]:
        writer.add_page(page)
    write_pdf(writer, out_path)
    where = "at the start" if at == 0 else ("at the end" if at == total else f"before page {at + 1}")
    print(f"Inserted {len(incoming)} page(s) {where}.")
    print(f"Pages: {total} -> {total + len(incoming)}")


def cmd_reorder(in_path, out_path, order):
    reader = open_reader(in_path)
    total = len(reader.pages)
    numbers = []
    for part in order.replace(" ", "").split(","):
        if not part.isdigit():
            fail(f"'{part}' is not a page number.", 2)
        value = int(part)
        if value < 1 or value > total:
            fail(f"Page {value} is outside the document (1 to {total}).", 2)
        numbers.append(value)
    if sorted(numbers) != list(range(1, total + 1)):
        fail("The new order must list every page exactly once.", 2)
    writer = PdfWriter()
    for value in numbers:
        writer.add_page(reader.pages[value - 1])
    write_pdf(writer, out_path)
    print(f"Reordered {total} pages.")


def cmd_count(list_file):
    for f in read_list(list_file):
        name = Path(f).name
        try:
            reader = PdfReader(f)
            if reader.is_encrypted and not reader.decrypt(password()):
                print(f"{name}\tLOCKED")
                continue
            print(f"{name}\t{len(reader.pages)}")
        except Exception:
            print(f"{name}\tERROR")


def cmd_unlock(in_path, out_path):
    reader = open_reader(in_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    write_pdf(writer, out_path)


def cmd_text(in_path, out_path):
    reader = open_reader(in_path)
    parts = []
    found = 0
    for i, page in enumerate(reader.pages, start=1):
        body = page.extract_text() or ""
        found += len(body.strip())
        parts.append(f"--- Page {i} ---\n{body}\n")
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(parts))
    print(f"Created: {out.name}")
    print(f"Characters of text: {found}")
    if found < 50:
        print("WARNING: almost no text was found. This PDF is probably scanned. OCR it first.")


def cmd_search(list_file, term):
    needle = term.lower()
    hits = 0
    scanned_only = 0
    for f in read_list(list_file):
        name = Path(f).name
        try:
            reader = PdfReader(f)
            if reader.is_encrypted and not reader.decrypt(password()):
                print(f"{name}\tLOCKED\t")
                continue
            chars = 0
            for i, page in enumerate(reader.pages, start=1):
                body = (page.extract_text() or "")
                chars += len(body.strip())
                if needle in body.lower():
                    hits += 1
                    print(f"{name}\t{i}\tmatch")
            if chars < 50:
                scanned_only += 1
                print(f"{name}\tNOTEXT\t")
        except Exception:
            print(f"{name}\tERROR\t")
    print(f"MATCHES={hits}")
    print(f"NOTEXTFILES={scanned_only}")


def rows_from_words(page, gap=6.0, line_tolerance=3.0):
    """Build rows by grouping words into lines, then splitting a line into
    cells wherever the horizontal gap between words is wide. Whole words are
    kept, so amounts are never clipped part-way through."""
    words = page.extract_words()
    if not words:
        return []
    lines = {}
    for word in words:
        key = round(word["top"] / line_tolerance)
        lines.setdefault(key, []).append(word)
    rows = []
    for key in sorted(lines):
        line = sorted(lines[key], key=lambda w: w["x0"])
        cells = []
        current = [line[0]["text"]]
        for previous, word in zip(line, line[1:]):
            if word["x0"] - previous["x1"] > gap:
                cells.append(" ".join(current))
                current = [word["text"]]
            else:
                current.append(word["text"])
        cells.append(" ".join(current))
        rows.append(cells)
    return rows


def cmd_tables(in_path, out_base):
    try:
        import pdfplumber
    except ImportError:
        fail("pdfplumber is not installed. Run Install PDFToolbox.bat.", 10)
    tables = []
    method = "ruled lines"
    with pdfplumber.open(in_path, password=password()) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            for table in page.extract_tables():
                if table:
                    tables.append((i, table))
        # Many statements have no ruling lines. Fall back to column spacing.
        if not tables:
            method = "column spacing"
            for i, page in enumerate(pdf.pages, start=1):
                rows = [r for r in rows_from_words(page) if len(r) > 1]
                if len(rows) > 1:
                    tables.append((i, rows))
    if not tables:
        fail("No tables were found. If this PDF is scanned, OCR it first.", 2)
    print(f"Detected by: {method}")
    try:
        from openpyxl import Workbook
    except ImportError:
        import csv
        out_dir = Path(out_base)
        out_dir.mkdir(parents=True, exist_ok=True)
        for idx, (pageno, table) in enumerate(tables, start=1):
            out = out_dir / f"page {pageno:03} table {idx}.csv"
            with open(out, "w", newline="", encoding="utf-8-sig") as fh:
                csv.writer(fh).writerows([["" if c is None else c for c in row] for row in table])
            print(f"Created: {out.name}")
        print(f"Tables: {len(tables)}")
        return
    book = Workbook()
    book.remove(book.active)
    for idx, (pageno, table) in enumerate(tables, start=1):
        sheet = book.create_sheet(title=f"p{pageno} t{idx}"[:31])
        for row in table:
            sheet.append(["" if c is None else str(c) for c in row])
    out = str(out_base) + ".xlsx"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    book.save(out)
    print(f"Created: {Path(out).name}")
    print(f"Tables: {len(tables)}")


def cmd_imgtopdf(list_file, out_path):
    try:
        from PIL import Image
    except ImportError:
        fail("Pillow is not installed. Run Install PDFToolbox.bat.", 10)
    images = []
    for f in read_list(list_file):
        try:
            img = Image.open(f)
        except Exception as exc:
            fail(f"Cannot read {Path(f).name}: {exc}")
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGB")
        images.append(img)
        print(f"Added: {Path(f).name}")
    if not images:
        fail("No images were found.", 2)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    images[0].save(str(out), save_all=True, append_images=images[1:])
    print(f"Created: {out.name}")
    print(f"Pages: {len(images)}")


COMMANDS = {
    "info": (cmd_info, 1),
    "merge": (cmd_merge, 2),
    "split": (cmd_split, 2),
    "extract": (cmd_extract, 3),
    "ranges": (cmd_ranges, 3),
    "delete": (cmd_delete, 3),
    "rotate": (cmd_rotate, 4),
    "count": (cmd_count, 1),
    "unlock": (cmd_unlock, 2),
    "insert": (cmd_insert, 5),
    "reorder": (cmd_reorder, 3),
    "text": (cmd_text, 2),
    "search": (cmd_search, 2),
    "tables": (cmd_tables, 2),
    "imgtopdf": (cmd_imgtopdf, 2),
}

if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in COMMANDS:
        fail("Unknown helper command.", 2)
    func, nargs = COMMANDS[sys.argv[1]]
    args = sys.argv[2:]
    if len(args) != nargs:
        fail(f"{sys.argv[1]} expects {nargs} arguments.", 2)
    try:
        func(*args)
    except SystemExit:
        raise
    except Exception as exc:
        fail(str(exc))
'@

# ======================================================================
# MODERN FOLDER PICKER
# FolderBrowserDialog is the small old tree view. This uses the same
# large, resizable Explorer window that Chrome and Outlook open.
# ======================================================================
$script:FolderPickerCode = @'
using System;
using System.Runtime.InteropServices;

namespace PdfToolbox
{
    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
    }

    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    internal class FileOpenDialogRcw { }

    public static class FolderPicker
    {
        private const uint FOS_PICKFOLDERS = 0x20;
        private const uint FOS_FORCEFILESYSTEM = 0x40;
        private const uint FOS_PATHMUSTEXIST = 0x800;
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppv);

        public static string Show(string title, string startFolder, IntPtr owner)
        {
            IFileDialog dialog = (IFileDialog)(new FileOpenDialogRcw());
            try
            {
                uint options;
                dialog.GetOptions(out options);
                dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
                if (!string.IsNullOrEmpty(title)) { dialog.SetTitle(title); }
                if (!string.IsNullOrEmpty(startFolder) && System.IO.Directory.Exists(startFolder))
                {
                    Guid shellItemId = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
                    object item;
                    SHCreateItemFromParsingName(startFolder, IntPtr.Zero, ref shellItemId, out item);
                    dialog.SetFolder((IShellItem)item);
                }
                int hr = dialog.Show(owner);
                if (hr != 0) { return null; }
                IShellItem result;
                dialog.GetResult(out result);
                string path;
                result.GetDisplayName(SIGDN_FILESYSPATH, out path);
                return path;
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }
    }
}
'@

# ======================================================================
# NAVIGATION (P = previous step, M = main menu)
# ======================================================================
function Stop-Task { throw 'PDFTB_MENU' }
function Step-Back { throw 'PDFTB_BACK' }

function Test-NavigationWord {
    param([string]$Answer)
    if ($Answer -ieq 'p') { Step-Back }
    if ($Answer -ieq 'm') { Stop-Task }
}

function Invoke-Steps {
    param([string]$Title, [hashtable]$State, [array]$Steps)
    if (-not $State.ContainsKey('Recap')) { $State['Recap'] = @{} }
    $i = 0
    while ($i -lt $Steps.Count) {
        Write-Header $Title
        Write-Host '  P = previous step    M = main menu' -ForegroundColor DarkGray
        Write-Host ''
        for ($k = 0; $k -lt $i; $k++) {
            if ($State['Recap'][$k]) { Write-Host "  $($State['Recap'][$k])" -ForegroundColor DarkGray }
        }
        if ($i -gt 0) { Write-Host '' }
        $State['StepIndex'] = $i
        try {
            & $Steps[$i] $State
            $i++
        } catch {
            $message = $_.Exception.Message
            if ($message -eq 'PDFTB_BACK') {
                $State['Recap'].Remove($i)
                if ($i -eq 0) { return }
                $i--
                $State['Recap'].Remove($i)
                continue
            }
            if ($message -eq 'PDFTB_MENU') { return }
            throw
        }
    }
}

function Set-Recap {
    param([hashtable]$State, [string]$Text)
    $State['Recap'][[int]$State['StepIndex']] = $Text
}

# ======================================================================
# GENERAL HELPERS
# ======================================================================
function Write-Header {
    param([string]$Title, [string]$Subtitle)
    $width = 58
    $rule = '=' * $width
    Clear-Host
    Write-Host ''
    Write-Host "  $rule" -ForegroundColor DarkCyan
    $left = '  ' + $Title.ToUpper()
    $right = "v$script:Version"
    $pad = $width - $left.Length - $right.Length
    if ($pad -lt 2) { $pad = 2 }
    Write-Host "  $left" -NoNewline -ForegroundColor Cyan
    Write-Host ((' ' * $pad) + $right) -ForegroundColor DarkGray
    if ($Subtitle) { Write-Host "    $Subtitle" -ForegroundColor DarkGray }
    Write-Host "  $rule" -ForegroundColor DarkCyan
    Write-Host ''
}

function Wait-ForEnter {
    Write-Host ''
    [void](Read-Host 'Press Enter to return to the menu')
}

function Get-ToolboxSettings {
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            $s = Get-Content -LiteralPath $script:ConfigFile -Raw | ConvertFrom-Json
            if ($s.LastFolder -and (Test-Path -LiteralPath $s.LastFolder)) { $script:LastFolder = $s.LastFolder }
        } catch { }
    }
}

function Save-ToolboxSettings {
    try {
        New-Item -ItemType Directory -Force -Path $script:ConfigDir | Out-Null
        @{ LastFolder = $script:LastFolder } | ConvertTo-Json | Set-Content -LiteralPath $script:ConfigFile -Encoding UTF8
    } catch { }
}

function Read-Text {
    param([string]$Prompt, [switch]$NoNavigation)
    $answer = ([string](Read-Host $Prompt)).Trim()
    if (-not $NoNavigation) { Test-NavigationWord $answer }
    return $answer
}

function Read-Choice {
    param([string]$Prompt, [string[]]$Valid, [string]$Default)
    while ($true) {
        $suffix = if ($Default) { " [$Default]" } else { '' }
        $answer = ([string](Read-Host "$Prompt$suffix")).Trim()
        Test-NavigationWord $answer
        if (-not $answer -and $Default) { return $Default }
        if ($Valid -contains $answer) { return $answer.ToUpper() }
        Write-Host "  Please enter one of: $($Valid -join ', ')   (or P / M)" -ForegroundColor Yellow
    }
}

function Read-YesNo {
    param([string]$Prompt, [string]$Default = 'N')
    return ((Read-Choice "$Prompt (Y/N)" @('Y', 'N') $Default) -eq 'Y')
}

function Get-ConsoleHandle {
    try { return [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle } catch { return [IntPtr]::Zero }
}

function Test-CanBrowse {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        Write-Host '  Browsing needs an STA window. Paste or drag the path instead.' -ForegroundColor Yellow
        return $false
    }
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop; return $true }
    catch {
        Write-Host '  Browsing is not available on this machine. Paste or drag the path instead.' -ForegroundColor Yellow
        return $false
    }
}

function Show-FilePicker {
    param([string]$Title, [string[]]$Extensions = @(), [switch]$Multiple)
    if (-not (Test-CanBrowse)) { return $null }
    $filter = 'All files (*.*)|*.*'
    if ($Extensions.Count -gt 0) {
        $pattern = ($Extensions | ForEach-Object { "*$_" }) -join ';'
        $filter = "Supported files ($pattern)|$pattern|All files (*.*)|*.*"
    }
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $filter
    $dialog.AutoUpgradeEnabled = $true
    $dialog.Multiselect = [bool]$Multiple
    if ($script:LastFolder -and (Test-Path -LiteralPath $script:LastFolder)) { $dialog.InitialDirectory = $script:LastFolder }
    $owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }
    try {
        if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($Multiple) { return @($dialog.FileNames) }
            return $dialog.FileName
        }
        return $null
    } finally { $owner.Dispose() }
}

function Show-FolderPicker {
    param([string]$Title)
    if (-not (Test-CanBrowse)) { return $null }
    if ($script:PickerType -ne 'Old') {
        try {
            if (-not ('PdfToolbox.FolderPicker' -as [type])) {
                Add-Type -TypeDefinition $script:FolderPickerCode -ErrorAction Stop
            }
            $start = if ($script:LastFolder -and (Test-Path -LiteralPath $script:LastFolder)) { $script:LastFolder } else { '' }
            $script:PickerType = 'Modern'
            return [PdfToolbox.FolderPicker]::Show($Title, $start, (Get-ConsoleHandle))
        } catch {
            $script:PickerType = 'Old'
        }
    }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title
    $dialog.ShowNewFolderButton = $true
    if ($script:LastFolder -and (Test-Path -LiteralPath $script:LastFolder)) { $dialog.SelectedPath = $script:LastFolder }
    $owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }
    try {
        if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
        return $null
    } finally { $owner.Dispose() }
}

function Read-UserPath {
    param(
        [string]$Prompt,
        [ValidateSet('File', 'Folder', 'Any')][string]$Type = 'Any',
        [string[]]$Extensions = @()
    )
    while ($true) {
        Write-Host $Prompt
        Write-Host '  Paste a path (Shift + right-click a file in Explorer, then "Copy as path"),' -ForegroundColor DarkGray
        Write-Host '  or drag it into this window, or type B to browse.' -ForegroundColor DarkGray
        if ($script:LastFolder) {
            if ($Type -eq 'File') {
                Write-Host "  A file name on its own is looked up in: $($script:LastFolder)" -ForegroundColor DarkGray
            } else {
                Write-Host "  Press Enter to use: $($script:LastFolder)" -ForegroundColor DarkGray
            }
        }
        $raw = [string](Read-Host '>')
        $path = $raw.Trim().Trim('"').Trim("'").Trim()
        Test-NavigationWord $path
        if ($path -ieq 'q') { Stop-Task }
        if ($path -ieq 'b') {
            $browseType = $Type
            if ($Type -eq 'Any') {
                Write-Host '  Browse for:  1  a file    2  a folder'
                $browseType = if ((Read-Choice '  Choice' @('1', '2') '1') -eq '1') { 'File' } else { 'Folder' }
            }
            $picked = if ($browseType -eq 'Folder') { Show-FolderPicker -Title $Prompt } else { Show-FilePicker -Title $Prompt -Extensions $Extensions }
            if (-not $picked) { Write-Host '  Nothing selected.' -ForegroundColor Yellow; continue }
            $path = $picked
        }
        if (-not $path) {
            if ($script:LastFolder -and $Type -ne 'File') { $path = $script:LastFolder }
            else { Write-Host '  Nothing entered.' -ForegroundColor Yellow; continue }
        }
        if (-not [System.IO.Path]::IsPathRooted($path) -and $script:LastFolder) {
            $path = Join-Path $script:LastFolder $path
        }
        if (-not (Test-Path -LiteralPath $path)) {
            Write-Host "  Not found: $path" -ForegroundColor Yellow
            continue
        }
        $item = Get-Item -LiteralPath $path
        if ($Type -eq 'File' -and $item.PSIsContainer) { Write-Host '  That is a folder. Please enter a file.' -ForegroundColor Yellow; continue }
        if ($Type -eq 'Folder' -and -not $item.PSIsContainer) { Write-Host '  That is a file. Please enter a folder.' -ForegroundColor Yellow; continue }
        if (-not $item.PSIsContainer -and $Extensions.Count -gt 0 -and ($Extensions -notcontains $item.Extension.ToLower())) {
            Write-Host "  Expected a file of type: $($Extensions -join ', ')" -ForegroundColor Yellow
            continue
        }
        $script:LastFolder = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
        Save-ToolboxSettings
        Write-Host ''
        return $item.FullName
    }
}

function Read-OutputName {
    param([string]$Default, [string]$Extension = '.pdf')
    $name = (Read-Text "Output file name [$Default]").Trim('"')
    if (-not $name) { $name = $Default }
    if ($name -notmatch ([regex]::Escape($Extension) + '$')) { $name += $Extension }
    if ($name.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        Write-Host '  That name has characters Windows does not allow. Using the default.' -ForegroundColor Yellow
        $name = $Default
    }
    return $name
}

function Confirm-Overwrite {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    return (Read-YesNo "$(Split-Path $Path -Leaf) already exists. Overwrite?" 'N')
}

function Open-OutputFolder {
    param([string]$Folder)
    Write-Host ''
    if ((Test-Path -LiteralPath $Folder) -and (Read-YesNo 'Open the output folder?' 'Y')) {
        Invoke-Item -LiteralPath $Folder
    }
}

function Select-OutputFolder {
    param([hashtable]$State, [string]$DefaultName)
    Write-Host 'Where should the results be saved?'
    Write-Host "  1  '$DefaultName' folder inside: $($State['SourceFolder'])"
    Write-Host '  2  Choose another folder'
    $choice = Read-Choice 'Choice' @('1', '2') '1'
    if ($choice -eq '1') { return (Join-Path $State['SourceFolder'] $DefaultName) }
    $picked = Show-FolderPicker -Title 'Choose where to save the results'
    if (-not $picked) { $picked = Read-UserPath -Prompt 'Enter the output folder:' -Type Folder }
    $script:LastFolder = $picked
    Save-ToolboxSettings
    return $picked
}

function Show-FileList {
    param($Files)
    for ($k = 0; $k -lt $Files.Count; $k++) {
        Write-Host ('  {0,3}. {1}' -f ($k + 1), $Files[$k].Name)
    }
    Write-Host ''
}

function New-ListFile {
    param([string[]]$Paths)
    $file = Join-Path $script:WorkDir ('list{0}.txt' -f [guid]::NewGuid().ToString('N'))
    [System.IO.File]::WriteAllLines($file, $Paths, (New-Object System.Text.UTF8Encoding($false)))
    return $file
}

# %TEMP% is often an 8.3 short path such as C:\Users\ISMAIL~1.KOY\...
# Some path handling chokes on those, so expand to the real long form.
function Expand-LongPath {
    param([string]$Path)
    if (-not $Path) { return $Path }
    try {
        if (-not ('PdfToolbox.PathHelper' -as [type])) {
            Add-Type -Namespace PdfToolbox -Name PathHelper -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
public static extern uint GetLongPathName(string shortPath, System.Text.StringBuilder longPath, uint buffer);
'@ -ErrorAction Stop
        }
        $builder = New-Object System.Text.StringBuilder 1024
        $length = [PdfToolbox.PathHelper]::GetLongPathName($Path, $builder, 1024)
        if ($length -gt 0 -and $length -lt 1024) { return $builder.ToString() }
    } catch { }
    return $Path
}

# Logging must never be the thing that kills a batch.
function Write-RunLog {
    param([string]$LogFile, [string]$Text)
    try { Add-Content -LiteralPath $LogFile -Value $Text -ErrorAction Stop }
    catch { }
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return "$Bytes bytes"
}

# ======================================================================
# PASSWORD HANDLING (held only in memory for the current task)
# ======================================================================
function Set-PdfPassword {
    param([string]$Prompt)
    $secure = Read-Host -Prompt $Prompt -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    $env:PDFTB_PASSWORD = $plain
}

function Clear-PdfPassword {
    Remove-Item Env:PDFTB_PASSWORD -ErrorAction SilentlyContinue
}

# ======================================================================
# CALLING PYTHON
# ======================================================================
function Invoke-Helper {
    param([string[]]$Arguments, [switch]$Quiet)
    if (-not $script:PyExe) { throw 'Python was not found. Run Install PDFToolbox.bat.' }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $lines = @(& $script:PyExe $script:Helper @Arguments 2>&1 | ForEach-Object { "$_" })
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
    if (-not $Quiet) {
        foreach ($l in $lines) {
            if ($l -like 'ERROR:*') { Write-Host "  $l" -ForegroundColor Red }
            elseif ($l -like 'WARNING:*') { Write-Host "  $l" -ForegroundColor Yellow }
            else { Write-Host "  $l" }
        }
    }
    return [pscustomobject]@{ Code = $code; Lines = $lines }
}

function Get-PdfInfo {
    param([string]$Path)
    $r = Invoke-Helper -Arguments @('info', $Path) -Quiet
    if ($r.Code -ne 0) {
        $r.Lines | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        return $null
    }
    $pages = [int](($r.Lines | Where-Object { $_ -like 'PAGES=*' } | Select-Object -First 1) -replace 'PAGES=', '')
    $enc = (($r.Lines | Where-Object { $_ -like 'ENCRYPTED=*' } | Select-Object -First 1) -eq 'ENCRYPTED=1')
    return [pscustomobject]@{ Pages = $pages; Encrypted = $enc }
}

function Get-PdfPages {
    param([string]$Path)
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        $info = Get-PdfInfo $Path
        if ($null -eq $info) { return $null }
        if ($info.Pages -ge 0) { return $info.Pages }
        if ($attempt -eq 3) { break }
        if ($attempt -gt 0) { Write-Host '  Wrong password.' -ForegroundColor Yellow }
        Set-PdfPassword -Prompt "$(Split-Path $Path -Leaf) is password protected. Enter the password (blank to cancel)"
        if (-not $env:PDFTB_PASSWORD) { return $null }
    }
    Write-Host '  Could not open the file with that password.' -ForegroundColor Red
    return $null
}

function Test-PdfsOpenable {
    param($Files)
    for ($attempt = 0; $attempt -lt 3; $attempt++) {
        $list = New-ListFile @($Files | ForEach-Object { $_.FullName })
        $r = Invoke-Helper -Arguments @('count', $list) -Quiet
        Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
        if ($r.Code -ne 0) {
            $r.Lines | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            return $false
        }
        $rows = @($r.Lines | Where-Object { $_ -match "`t" } | ForEach-Object {
                $p = $_ -split "`t"; [pscustomobject]@{ Name = $p[0]; Pages = $p[1] } })
        $broken = @($rows | Where-Object { $_.Pages -eq 'ERROR' })
        $locked = @($rows | Where-Object { $_.Pages -eq 'LOCKED' })
        if ($broken.Count -gt 0) {
            Write-Host 'These files could not be read (damaged or not real PDFs):' -ForegroundColor Red
            $broken | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Red }
            return $false
        }
        if ($locked.Count -eq 0) { return $true }
        Write-Host 'These files are password protected:' -ForegroundColor Yellow
        $locked | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Yellow }
        Set-PdfPassword -Prompt 'Enter the password (blank to cancel)'
        if (-not $env:PDFTB_PASSWORD) { return $false }
    }
    Write-Host 'Still locked after 3 attempts.' -ForegroundColor Red
    return $false
}

function Select-TargetFiles {
    param([hashtable]$State, [string]$Prompt, [string[]]$Extensions, [string]$Kind)
    $target = Read-UserPath -Prompt $Prompt -Type Any -Extensions $Extensions
    $item = Get-Item -LiteralPath $target
    if ($item.PSIsContainer) {
        $files = @(Get-ChildItem -LiteralPath $target -File | Where-Object {
                $Extensions -contains $_.Extension.ToLower() -and $_.Name -notlike '~$*' } | Sort-Object Name)
        $State['SourceFolder'] = $item.FullName
    } else {
        $files = @($item)
        $State['SourceFolder'] = $item.DirectoryName
    }
    if ($files.Count -eq 0) {
        Write-Host "No $Kind files found there." -ForegroundColor Yellow
        Stop-Task
    }
    $State['Files'] = $files
    Set-Recap $State ("Input: {0} ({1} file(s))" -f $target, $files.Count)
}

function Select-SinglePdf {
    param([hashtable]$State, [string]$Prompt = 'Enter the PDF:')
    $file = Read-UserPath -Prompt $Prompt -Type File -Extensions @('.pdf')
    $pages = Get-PdfPages $file
    if ($null -eq $pages) { Stop-Task }
    $State['File'] = $file
    $State['Item'] = Get-Item -LiteralPath $file
    $State['Pages'] = $pages
    Set-Recap $State "File: $(Split-Path $file -Leaf) ($pages pages)"
}

# ======================================================================
# GHOSTSCRIPT (compress, PDF to images)
# ======================================================================
function Get-GhostscriptExe {
    foreach ($name in @('gswin64c', 'gswin32c', 'gs')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Invoke-Ghostscript {
    param([string[]]$Arguments)
    $exe = Get-GhostscriptExe
    if (-not $exe) { throw 'Ghostscript was not found. Run Install PDFToolbox.bat.' }
    $all = @('-dNOPAUSE', '-dBATCH', '-dQUIET', '-dSAFER')
    if ($env:PDFTB_PASSWORD) { $all += "-sPDFPassword=$env:PDFTB_PASSWORD" }
    $all += $Arguments
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = @(& $exe @all 2>&1 | ForEach-Object { "$_" })
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previous }
    return [pscustomobject]@{ Code = $code; Lines = $output }
}

# ======================================================================
# OCR
# ======================================================================
function Get-OcrCommand {
    $cmd = Get-Command ocrmypdf -ErrorAction SilentlyContinue
    if ($cmd) { return @{ Exe = $cmd.Source; Pre = @() } }
    if ($script:PyExe) {
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { & $script:PyExe -m ocrmypdf --version *> $null; $ok = ($LASTEXITCODE -eq 0) }
        finally { $ErrorActionPreference = $previous }
        if ($ok) { return @{ Exe = $script:PyExe; Pre = @('-m', 'ocrmypdf') } }
    }
    return $null
}

function Get-ExtraTesseractLanguages {
    $cmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if (-not $cmd) { return @() }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $out = @(& tesseract --list-langs 2>&1 | ForEach-Object { "$_".Trim() }) }
    catch { return @() }
    finally { $ErrorActionPreference = $previous }
    return @($out | Where-Object { $_ -match '^[a-z_]{3,}$' -and $_ -notin @('eng', 'osd') })
}

function Get-OcrExitText {
    param($Code)
    switch ($Code) {
        0 { 'OK' }
        1 { 'Bad arguments' }
        2 { 'Input file is not a valid PDF' }
        3 { 'Missing dependency (Tesseract or Ghostscript).' }
        4 { 'Output PDF failed validation' }
        5 { 'File access error. Is the file open in another program?' }
        6 { 'Page already has text. Re-run using Mixed or Force mode.' }
        7 { 'Tesseract or Ghostscript crashed on this file' }
        8 { 'Password protected. Remove the password first.' }
        9 { 'Invalid configuration' }
        10 { 'PDF/A conversion failed' }
        130 { 'Cancelled' }
        default { "Failed (exit code $Code)" }
    }
}

# OCRmyPDF exit 4 means the output failed its final check. That is usually a
# damaged image inside the original scan, and the file is still complete and
# searchable. Confirm it opens and has every page before calling it a failure.
function Test-OcrOutput {
    param([string]$InPath, [string]$OutPath)
    if (-not (Test-Path -LiteralPath $OutPath)) {
        return [pscustomobject]@{ Usable = $false; Reason = 'no output file was written' }
    }
    $inInfo = Get-PdfInfo $InPath
    $outInfo = Get-PdfInfo $OutPath
    if ($null -eq $outInfo -or $outInfo.Pages -lt 1) {
        return [pscustomobject]@{ Usable = $false; Reason = 'the output file cannot be opened' }
    }
    if ($null -ne $inInfo -and $inInfo.Pages -gt 0 -and $inInfo.Pages -ne $outInfo.Pages) {
        return [pscustomobject]@{ Usable = $false; Reason = "page count changed: $($inInfo.Pages) in, $($outInfo.Pages) out" }
    }
    return [pscustomobject]@{ Usable = $true; Reason = "all $($outInfo.Pages) page(s) present and readable" }
}

function Update-OcrJobs {
    param($Jobs, [string]$LogFile, [hashtable]$Stats)
    foreach ($j in $Jobs) {
        if ($j.Done -or -not $j.Proc.HasExited) { continue }
        $j.Done = $true
        # Nothing in here may throw: one odd file must not stop the batch.
        try {
            $code = $j.Proc.ExitCode
            if ($code -eq 0) {
                $Stats.OK++
                Write-Host "  Done:    $($j.Name)" -ForegroundColor Green
                Write-RunLog $LogFile "OK      $($j.Name)"
            } elseif ($code -eq 4) {
                $check = Test-OcrOutput -InPath $j.In -OutPath $j.OutPath
                if ($check.Usable) {
                    $Stats.Warn++
                    Write-Host "  Done:    $($j.Name)  (check this one)" -ForegroundColor Yellow
                    Write-RunLog $LogFile "WARN    $($j.Name) (damaged image in the original scan; $($check.Reason))"
                } else {
                    $Stats.Failed++
                    Write-Host "  FAILED:  $($j.Name) ($($check.Reason))" -ForegroundColor Red
                    Write-RunLog $LogFile "FAILED  $($j.Name) ($($check.Reason))"
                }
                try {
                    if (Test-Path -LiteralPath $j.Err) {
                        Get-Content -LiteralPath $j.Err -Tail 6 -ErrorAction SilentlyContinue | ForEach-Object {
                            if ($_.Trim()) { Write-RunLog $LogFile "        $_" }
                        }
                    }
                } catch { }
            } else {
                $Stats.Failed++
                $reason = Get-OcrExitText $code
                Write-Host "  FAILED:  $($j.Name) ($reason)" -ForegroundColor Red
                Write-RunLog $LogFile "FAILED  $($j.Name) ($reason)"
                try {
                    if (Test-Path -LiteralPath $j.Err) {
                        Get-Content -LiteralPath $j.Err -Tail 5 -ErrorAction SilentlyContinue | ForEach-Object {
                            if ($_.Trim()) { Write-RunLog $LogFile "        $_" }
                        }
                    }
                } catch { }
            }
        } catch {
            $Stats.Failed++
            Write-Host "  PROBLEM: $($j.Name) ($($_.Exception.Message))" -ForegroundColor Red
            Write-RunLog $LogFile "FAILED  $($j.Name) (while finishing: $($_.Exception.Message))"
        }
        try { Remove-Item -LiteralPath $j.Err, $j.Out -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Invoke-OcrMenu {
    $ocr = Get-OcrCommand
    if (-not $ocr) {
        Write-Header 'OCR PDFs'
        Write-Host 'OCRmyPDF was not found. Run Install PDFToolbox.bat.' -ForegroundColor Red
        return
    }
    $extraLanguages = Get-ExtraTesseractLanguages
    $state = @{ Ocr = $ocr; ExtraLanguages = $extraLanguages }
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            Write-Host 'What kind of files are these?'
            Write-Host '  1  Fully scanned, no searchable text (fastest)'
            Write-Host '  2  Mixed: skip pages that already have text'
            Write-Host '  3  Force OCR: replace any existing text layer'
            $s['Mode'] = Read-Choice 'Choice' @('1', '2', '3') '1'
            $names = @{ '1' = 'Fully scanned'; '2' = 'Mixed'; '3' = 'Force OCR' }
            Set-Recap $s "Type: $($names[$s['Mode']])"
        },
        {
            param($s)
            $s['Deskew'] = Read-YesNo 'Straighten and auto-rotate pages? Only needed for crooked or sideways scans' 'N'
            Set-Recap $s ('Straighten / rotate: ' + $(if ($s['Deskew']) { 'yes' } else { 'no' }))
        },
        {
            param($s)
            Write-Host 'Compression:'
            Write-Host '  0  None (fastest)'
            Write-Host '  1  Lossless (safe default)'
            Write-Host '  2  Balanced (smaller files, slight image quality loss)'
            $s['Optimize'] = Read-Choice 'Choice' @('0', '1', '2') '1'
            Set-Recap $s "Compression: level $($s['Optimize'])"
        },
        {
            param($s)
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'OCR PDFs'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            $cores = [Environment]::ProcessorCount
            # OCRmyPDF already spreads one file's pages across cores, so the
            # cores are shared between the files running at the same time.
            $suggested = [Math]::Min($s['Files'].Count, $(if ($cores -le 4) { 2 } else { 3 }))
            Write-Host "This machine has $cores logical cores."
            $raw = Read-Text "Files to process at the same time [$suggested]"
            $s['MaxJobs'] = if ($raw -match '^\d+$' -and [int]$raw -ge 1) { [int]$raw } else { $suggested }
            $s['JobsPerFile'] = [Math]::Max(1, [Math]::Floor($cores / $s['MaxJobs']))
            Set-Recap $s "At a time: $($s['MaxJobs']) file(s), $($s['JobsPerFile']) core(s) each"
        },
        {
            param($s)
            if (-not (Read-YesNo "Start OCR on $($s['Files'].Count) file(s)?" 'Y')) { Step-Back }
            Start-OcrRun -State $s
        }
    )
    # The language question only appears when a pack other than English is
    # installed. A standard Tesseract install has only eng, so it is skipped.
    if ($extraLanguages.Count -gt 0) {
        $languageStep = {
            param($s)
            Write-Host "Installed language packs besides English: $($s['ExtraLanguages'] -join ', ')"
            Write-Host 'Leave blank for English. Join languages with +, e.g. eng+afr' -ForegroundColor DarkGray
            $lang = Read-Text 'OCR language(s) [eng]'
            if ($lang -and $lang -notmatch '^[A-Za-z_]+(\+[A-Za-z_]+)*$') {
                Write-Host '  Not a valid language code. Using eng.' -ForegroundColor Yellow
                $lang = ''
            }
            $s['Language'] = $lang
            Set-Recap $s ('Language: ' + $(if ($lang) { $lang } else { 'eng' }))
        }
        $steps = @($steps[0..3]) + @($languageStep) + @($steps[4..($steps.Count - 1)])
    }
    Invoke-Steps -Title 'OCR PDFs' -State $state -Steps $steps
}

function Start-OcrRun {
    param([hashtable]$State)
    $ocr = $State['Ocr']
    $files = $State['Files']
    $outDir = $State['OutDir']
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $sameFolder = ((Get-Item -LiteralPath $outDir).FullName.TrimEnd('\') -ieq ([string]$State['SourceFolder']).TrimEnd('\'))

    $flags = New-Object System.Collections.Generic.List[string]
    switch ($State['Mode']) { '2' { $flags.Add('--skip-text') } '3' { $flags.Add('--force-ocr') } }
    if ($State['Deskew']) { $flags.Add('--rotate-pages'); $flags.Add('--deskew') }
    $flags.Add('--optimize'); $flags.Add([string]$State['Optimize'])
    if ($State['Language']) { $flags.Add('-l'); $flags.Add([string]$State['Language']) }
    $flags.Add('--jobs'); $flags.Add([string]$State['JobsPerFile'])

    $logFile = Join-Path $outDir 'OCR log.txt'
    $tmpDir = Join-Path $script:WorkDir 'ocr'
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
    $tmpDir = Expand-LongPath $tmpDir
    Write-RunLog $logFile ''
    Write-RunLog $logFile ('=== OCR run started {0} | {1} file(s) | options: {2} ===' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $files.Count, ($flags -join ' '))
    Write-RunLog $logFile ("    $script:Stamp | run by $env:USERNAME on $env:COMPUTERNAME")

    Write-Host ''
    Write-Host "Running $($State['MaxJobs']) file(s) at a time." -ForegroundColor Cyan
    $jobs = New-Object System.Collections.ArrayList
    $stats = @{ OK = 0; Warn = 0; Failed = 0 }
    $started = Get-Date
    $i = 0
    foreach ($file in $files) {
        $i++
        while (@($jobs | Where-Object { -not $_.Done -and -not $_.Proc.HasExited }).Count -ge $State['MaxJobs']) {
            Start-Sleep -Milliseconds 700
            Update-OcrJobs -Jobs $jobs -LogFile $logFile -Stats $stats
        }
        Update-OcrJobs -Jobs $jobs -LogFile $logFile -Stats $stats

        $outName = if ($sameFolder) { "$($file.BaseName) OCR.pdf" } else { $file.Name }
        $outPath = Join-Path $outDir $outName
        $id = [guid]::NewGuid().ToString('N')
        $errFile = Join-Path $tmpDir "$id.err.txt"
        $outFile = Join-Path $tmpDir "$id.out.txt"
        # Paths are wrapped in quotes so spaces survive PowerShell 5.1's argument joining.
        $argList = @($ocr.Pre) + @($flags) + @(('"{0}"' -f $file.FullName), ('"{0}"' -f $outPath))
        try {
            $proc = Start-Process -FilePath $ocr.Exe -ArgumentList $argList -PassThru -NoNewWindow `
                -RedirectStandardError $errFile -RedirectStandardOutput $outFile
            $null = $proc.Handle   # without this, PowerShell 5.1 can lose the exit code
            [void]$jobs.Add([pscustomobject]@{ Proc = $proc; Name = $file.Name; Err = $errFile; Out = $outFile
                    In = $file.FullName; OutPath = $outPath; Done = $false
                })
            Write-Host ('[{0}/{1}] Started: {2}' -f $i, $files.Count, $file.Name)
        } catch {
            $stats.Failed++
            Write-Host ('[{0}/{1}] COULD NOT START: {2}' -f $i, $files.Count, $file.Name) -ForegroundColor Red
            Write-Host "          $($_.Exception.Message)" -ForegroundColor Red
            Write-RunLog $logFile "FAILED  $($file.Name) (could not start: $($_.Exception.Message))"
        }
    }
    while (@($jobs | Where-Object { -not $_.Done }).Count -gt 0) {
        Start-Sleep -Milliseconds 700
        Update-OcrJobs -Jobs $jobs -LogFile $logFile -Stats $stats
    }

    $elapsed = (Get-Date) - $started
    Write-RunLog $logFile ('=== Finished {0} | {1} OK | {2} to check | {3} failed | {4:hh\:mm\:ss} ===' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $stats.OK, $stats.Warn, $stats.Failed, $elapsed)
    Write-Host ''
    $colour = if ($stats.Failed) { 'Red' } elseif ($stats.Warn) { 'Yellow' } else { 'Green' }
    Write-Host ('Finished in {0:hh\:mm\:ss}.  OK: {1}   To check: {2}   Failed: {3}' -f $elapsed, $stats.OK, $stats.Warn, $stats.Failed) -ForegroundColor $colour
    if ($stats.Warn) {
        Write-Host ''
        Write-Host 'The "to check" files were OCR''d and every page is there. The PDF checker' -ForegroundColor Yellow
        Write-Host 'found a damaged image inside the original scan, which carries through to the' -ForegroundColor Yellow
        Write-Host 'new file. Open those pages and confirm the image still looks right.' -ForegroundColor Yellow
    }
    if ($stats.Failed) { Write-Host 'Check the log for the reason each file failed.' -ForegroundColor Red }
    Write-Host "Log: $logFile"
    Open-OutputFolder $outDir
}

# ======================================================================
# MERGE
# ======================================================================
function Invoke-MergeMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            $folder = Read-UserPath -Prompt 'Enter the folder that contains the PDFs:' -Type Folder
            $files = @(Get-ChildItem -LiteralPath $folder -Filter *.pdf -File)
            if ($files.Count -lt 2) { Write-Host 'Need at least 2 PDFs in that folder.' -ForegroundColor Yellow; Stop-Task }
            $s['Folder'] = $folder
            $s['All'] = $files
            Set-Recap $s "Folder: $folder ($($files.Count) PDFs)"
        },
        {
            param($s)
            Write-Host 'Order:'
            Write-Host '  1  By file name'
            Write-Host '  2  By date modified (oldest first)'
            Write-Host '  3  Choose my own order or a subset'
            Write-Host '  4  Arrange visually (shows the first page of each file)'
            $order = Read-Choice 'Choice' @('1', '2', '3', '4') '1'
            switch ($order) {
                '1' { $s['Files'] = @($s['All'] | Sort-Object Name); Set-Recap $s 'Order: by file name' }
                '2' { $s['Files'] = @($s['All'] | Sort-Object LastWriteTime); Set-Recap $s 'Order: by date modified' }
                '3' {
                    $sorted = @($s['All'] | Sort-Object Name)
                    Write-Host ''
                    Show-FileList $sorted
                    while ($true) {
                        $raw = Read-Text 'Enter the numbers in the order you want, e.g. 3,1,2 (files left out are skipped)'
                        $nums = @($raw -split '[,\s]+' | Where-Object { $_ })
                        $bad = @($nums | Where-Object { $_ -notmatch '^\d+$' -or [int]$_ -lt 1 -or [int]$_ -gt $sorted.Count })
                        if ($nums.Count -eq 0 -or $bad.Count -gt 0) { Write-Host '  Invalid list. Try again.' -ForegroundColor Yellow; continue }
                        $dupes = @($nums | Group-Object | Where-Object { $_.Count -gt 1 })
                        if ($dupes.Count -gt 0) { Write-Host "  Each file can only be used once. Repeated: $(($dupes | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Yellow; continue }
                        $s['Files'] = @($nums | ForEach-Object { $sorted[[int]$_ - 1] })
                        break
                    }
                    Set-Recap $s "Order: chosen by hand ($($s['Files'].Count) files)"
                }
                '4' {
                    $arranged = Get-FileOrderVisually -Files @($s['All'] | Sort-Object Name)
                    if (-not $arranged) { Write-Host '  Nothing arranged. Pick another option.' -ForegroundColor Yellow; Step-Back }
                    $s['Files'] = @($arranged)
                    Write-Host ''
                    Show-FileList $s['Files']
                    Set-Recap $s "Order: arranged visually ($($s['Files'].Count) files)"
                }
            }
        },
        {
            param($s)
            $name = Read-OutputName 'Combined.pdf'
            $outPath = Join-Path $s['Folder'] $name
            $s['Files'] = @($s['Files'] | Where-Object { $_.FullName -ine $outPath })
            if ($s['Files'].Count -lt 2) { Write-Host 'Need at least 2 PDFs to merge.' -ForegroundColor Yellow; Stop-Task }
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            $s['OutPath'] = $outPath
            Set-Recap $s "Output: $name"
        },
        {
            param($s)
            Write-Host 'Files will be merged in this order:'
            Show-FileList $s['Files']
            if (-not (Read-YesNo 'Merge these files?' 'Y')) { Step-Back }
            if (-not (Test-PdfsOpenable $s['Files'])) { return }
            Write-Host ''
            $list = New-ListFile @($s['Files'] | ForEach-Object { $_.FullName })
            $r = Invoke-Helper -Arguments @('merge', $s['OutPath'], $list)
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            if ($r.Code -eq 0) {
                Write-Host ''
                Write-Host "Saved: $($s['OutPath'])" -ForegroundColor Green
                Open-OutputFolder $s['Folder']
            }
        }
    )
    Invoke-Steps -Title 'Merge PDFs' -State $state -Steps $steps
}

# ======================================================================
# PAGES SUBMENU: split, extract, ranges, every N, delete
# ======================================================================
function Invoke-SplitMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s -Prompt 'Enter the PDF to split:' },
        {
            param($s)
            $outDir = Join-Path $s['Item'].DirectoryName "$($s['Item'].BaseName) pages"
            Write-Host 'One file per page will be saved to:'
            Write-Host "  $outDir"
            Write-Host ''
            if (-not (Read-YesNo "Split into $($s['Pages']) files?" 'Y')) { Step-Back }
            $r = Invoke-Helper -Arguments @('split', $s['File'], $outDir)
            if ($r.Code -eq 0) { Open-OutputFolder $outDir }
        }
    )
    Invoke-Steps -Title 'Split a PDF into single pages' -State $state -Steps $steps
}

function Invoke-ExtractMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s },
        {
            param($s)
            Write-Host "This PDF has $($s['Pages']) pages."
            Write-Host '  1  Type the page numbers'
            Write-Host '  2  Pick them from page previews'
            $how = Read-Choice 'Choice' @('1', '2') '1'
            $spec = $null
            if ($how -eq '2') {
                $spec = Select-PagesVisually -PdfPath $s['File'] -Instruction 'Tick the pages to extract. Ctrl+click for more than one, Shift+click for a run. Double-click a page to see it full size.'
                if (-not $spec) { Write-Host '  Nothing picked, so type the pages instead.' -ForegroundColor Yellow }
            }
            if (-not $spec) {
                Write-Host 'Examples:  1,3,5    1-10    1-5,10,15-20' -ForegroundColor DarkGray
                $spec = Read-Text 'Pages to extract'
                if (-not $spec) { Step-Back }
            }
            $s['Spec'] = $spec
            Set-Recap $s "Pages: $spec"
        },
        {
            param($s)
            $name = Read-OutputName "$($s['Item'].BaseName) extract.pdf"
            $outPath = Join-Path $s['Item'].DirectoryName $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            $r = Invoke-Helper -Arguments @('extract', $s['File'], $outPath, $s['Spec'])
            if ($r.Code -eq 2) { Write-Host ''; Write-Host 'Go back one step to fix the page numbers.' -ForegroundColor Yellow; return }
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Extract pages into one new PDF' -State $state -Steps $steps
}

function Invoke-RangesMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s },
        {
            param($s)
            Write-Host "This PDF has $($s['Pages']) pages."
            Write-Host '  1  Enter my own ranges (e.g. 1-10,11-25,26-50)'
            Write-Host '  2  Split into equal chunks of N pages'
            $how = Read-Choice 'Choice' @('1', '2') '1'
            if ($how -eq '1') {
                $spec = Read-Text 'Ranges'
                if (-not $spec) { Step-Back }
            } else {
                $sizeRaw = Read-Text 'Pages per file'
                if ($sizeRaw -notmatch '^\d+$' -or [int]$sizeRaw -lt 1) { Write-Host '  Enter a whole number.' -ForegroundColor Yellow; Step-Back }
                $size = [int]$sizeRaw
                $parts = @()
                for ($start = 1; $start -le $s['Pages']; $start += $size) {
                    $end = [Math]::Min($start + $size - 1, $s['Pages'])
                    $parts += "$start-$end"
                }
                $spec = $parts -join ','
                Write-Host "  That makes $($parts.Count) files: $spec" -ForegroundColor DarkGray
            }
            $s['Spec'] = $spec
            Set-Recap $s "Ranges: $spec"
        },
        {
            param($s)
            $outDir = Join-Path $s['Item'].DirectoryName "$($s['Item'].BaseName) parts"
            Write-Host "Files will be saved to: $outDir"
            Write-Host ''
            if (-not (Read-YesNo 'Continue?' 'Y')) { Step-Back }
            $r = Invoke-Helper -Arguments @('ranges', $s['File'], $outDir, $s['Spec'])
            if ($r.Code -eq 2) { Write-Host ''; Write-Host 'Go back and try different ranges.' -ForegroundColor Yellow; return }
            if ($r.Code -eq 0) { Open-OutputFolder $outDir }
        }
    )
    Invoke-Steps -Title 'Split a PDF into page ranges' -State $state -Steps $steps
}

function Invoke-DeletePagesMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s },
        {
            param($s)
            Write-Host "This PDF has $($s['Pages']) pages."
            Write-Host '  1  Type the page numbers'
            Write-Host '  2  Pick them from page previews'
            $how = Read-Choice 'Choice' @('1', '2') '1'
            $spec = $null
            if ($how -eq '2') {
                $spec = Select-PagesVisually -PdfPath $s['File'] -Instruction 'Tick the pages to REMOVE. Ctrl+click for more than one, Shift+click for a run. Double-click a page to see it full size.'
                if (-not $spec) { Write-Host '  Nothing picked, so type the pages instead.' -ForegroundColor Yellow }
            }
            if (-not $spec) {
                Write-Host 'Examples:  1    1-3    1,5,9-12' -ForegroundColor DarkGray
                $spec = Read-Text 'Pages to remove'
                if (-not $spec) { Step-Back }
            }
            $s['Spec'] = $spec
            Set-Recap $s "Remove: $spec"
        },
        {
            param($s)
            $name = Read-OutputName "$($s['Item'].BaseName) trimmed.pdf"
            $outPath = Join-Path $s['Item'].DirectoryName $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            Write-Host 'The original file is not changed.' -ForegroundColor DarkGray
            $r = Invoke-Helper -Arguments @('delete', $s['File'], $outPath, $s['Spec'])
            if ($r.Code -eq 2) { Write-Host ''; Write-Host 'Go back one step to fix the page numbers.' -ForegroundColor Yellow; return }
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Delete pages from a PDF' -State $state -Steps $steps
}

function Invoke-ReorderMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s -Prompt 'Enter the PDF whose pages you want to rearrange:' },
        {
            param($s)
            Write-Host 'A window will open with a preview of every page.' -ForegroundColor DarkGray
            Write-Host 'Click a page, then Move back or Move forward to place it.' -ForegroundColor DarkGray
            Write-Host ''
            if (-not (Read-YesNo 'Open the page view?' 'Y')) { Step-Back }
            $order = Get-PageOrderVisually -PdfPath $s['File']
            if (-not $order) { Write-Host '  Nothing to apply.' -ForegroundColor Yellow; Stop-Task }
            $s['Order'] = ($order -join ',')
            Set-Recap $s "New order: $($s['Order'])"
        },
        {
            param($s)
            $name = Read-OutputName "$($s['Item'].BaseName) reordered.pdf"
            $outPath = Join-Path $s['Item'].DirectoryName $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            Write-Host 'The original file is not changed.' -ForegroundColor DarkGray
            $r = Invoke-Helper -Arguments @('reorder', $s['File'], $outPath, $s['Order'])
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Rearrange pages' -State $state -Steps $steps
}

function Invoke-InsertMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s -Prompt 'Enter the PDF to insert INTO:' },
        {
            param($s)
            Write-Host 'What do you want to insert?'
            Write-Host '  1  Pages from another PDF'
            Write-Host '  2  An image (scan, screenshot, photo)'
            $kind = Read-Choice 'Choice' @('1', '2') '1'
            if ($kind -eq '1') {
                $src = Read-UserPath -Prompt 'Enter the PDF to take pages from:' -Type File -Extensions @('.pdf')
                $srcPages = Get-PdfPages $src
                if ($null -eq $srcPages) { Stop-Task }
                $s['Source'] = $src
                $s['SourcePages'] = $srcPages
                Set-Recap $s "Inserting from: $(Split-Path $src -Leaf) ($srcPages pages)"
            } else {
                $src = Read-UserPath -Prompt 'Enter the image file:' -Type File -Extensions $script:ImageExtensions
                $s['Source'] = $src
                $s['SourcePages'] = 1
                $s['Spec'] = 'all'
                Set-Recap $s "Inserting image: $(Split-Path $src -Leaf)"
            }
            $s['Kind'] = $kind
        },
        {
            param($s)
            if ($s['Kind'] -eq '2' -or $s['SourcePages'] -eq 1) {
                $s['Spec'] = 'all'
            } else {
                Write-Host "That PDF has $($s['SourcePages']) pages."
                Write-Host 'Enter "all", or pages like 1-3,7' -ForegroundColor DarkGray
                $spec = Read-Text 'Pages to take [all]'
                if (-not $spec) { $spec = 'all' }
                $s['Spec'] = $spec
            }
            Set-Recap $s "Taking: $($s['Spec'])"
        },
        {
            param($s)
            Write-Host "The target PDF has $($s['Pages']) pages. Where should they go?"
            Write-Host '  1  At the start'
            Write-Host '  2  At the end'
            Write-Host '  3  Before a specific page'
            $where = Read-Choice 'Choice' @('1', '2', '3') '2'
            if ($where -eq '1') { $s['Position'] = 'start' }
            elseif ($where -eq '2') { $s['Position'] = 'end' }
            else {
                $n = Read-Text "Insert before which page? (1 to $($s['Pages']))"
                if ($n -notmatch '^\d+$' -or [int]$n -lt 1 -or [int]$n -gt $s['Pages']) {
                    Write-Host '  That is not a page in this PDF.' -ForegroundColor Yellow
                    Step-Back
                }
                $s['Position'] = $n
            }
            $label = switch ($s['Position']) { 'start' { 'at the start' } 'end' { 'at the end' } default { "before page $($s['Position'])" } }
            Set-Recap $s "Position: $label"
        },
        {
            param($s)
            $name = Read-OutputName "$($s['Item'].BaseName) updated.pdf"
            $outPath = Join-Path $s['Item'].DirectoryName $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            Write-Host 'The original file is not changed.' -ForegroundColor DarkGray
            Write-Host ''
            $r = Invoke-Helper -Arguments @('insert', $s['File'], $outPath, $s['Source'], $s['Spec'], [string]$s['Position'])
            if ($r.Code -eq 2) { Write-Host ''; Write-Host 'Go back and change the pages or the position.' -ForegroundColor Yellow; return }
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Insert pages into a PDF' -State $state -Steps $steps
}

function Invoke-RotateMenu {
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s },
        {
            param($s)
            Write-Host 'Rotate by:'
            Write-Host '  1  90 degrees clockwise'
            Write-Host '  2  180 degrees'
            Write-Host '  3  90 degrees anticlockwise'
            $choice = Read-Choice 'Choice' @('1', '2', '3') '1'
            $s['Degrees'] = @{ '1' = 90; '2' = 180; '3' = 270 }[$choice]
            Set-Recap $s "Rotate: $($s['Degrees']) degrees"
        },
        {
            param($s)
            Write-Host "This PDF has $($s['Pages']) pages."
            Write-Host 'Enter "all", or pages like 1-3,7' -ForegroundColor DarkGray
            $spec = Read-Text 'Pages to rotate [all]'
            if (-not $spec) { $spec = 'all' }
            $s['Spec'] = $spec
            Set-Recap $s "Pages: $spec"
        },
        {
            param($s)
            $name = Read-OutputName "$($s['Item'].BaseName) rotated.pdf"
            $outPath = Join-Path $s['Item'].DirectoryName $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            $r = Invoke-Helper -Arguments @('rotate', $s['File'], $outPath, $s['Spec'], [string]$s['Degrees'])
            if ($r.Code -eq 2) { Write-Host ''; Write-Host 'Go back one step to fix the page numbers.' -ForegroundColor Yellow; return }
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Rotate pages' -State $state -Steps $steps
}

# ======================================================================
# COMPRESS
# ======================================================================
function Invoke-CompressMenu {
    if (-not (Get-GhostscriptExe)) {
        Write-Header 'Compress PDFs'
        Write-Host 'Ghostscript was not found. Run Install PDFToolbox.bat.' -ForegroundColor Red
        return
    }
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            Write-Host 'How small?'
            Write-Host '  1  Light    (300 dpi images, print quality)'
            Write-Host '  2  Balanced (150 dpi images, fine on screen and for email)'
            Write-Host '  3  Maximum  (72 dpi images, smallest, visibly softer)'
            $choice = Read-Choice 'Choice' @('1', '2', '3') '2'
            $s['Preset'] = @{ '1' = '/printer'; '2' = '/ebook'; '3' = '/screen' }[$choice]
            Set-Recap $s "Level: $(@{ '1' = 'Light'; '2' = 'Balanced'; '3' = 'Maximum' }[$choice])"
        },
        {
            param($s)
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'Compressed'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            if (-not (Read-YesNo "Compress $($s['Files'].Count) file(s)?" 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $s['OutDir'] | Out-Null
            Write-Host ''
            $before = 0; $after = 0; $ok = 0; $failed = 0
            foreach ($f in $s['Files']) {
                $out = Join-Path $s['OutDir'] $f.Name
                if ($out -ieq $f.FullName) { $out = Join-Path $s['OutDir'] "$($f.BaseName) compressed.pdf" }
                $r = Invoke-Ghostscript @('-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                    "-dPDFSETTINGS=$($s['Preset'])", "-sOutputFile=$out", $f.FullName)
                if ($r.Code -eq 0 -and (Test-Path -LiteralPath $out)) {
                    $sizeBefore = $f.Length
                    $sizeAfter = (Get-Item -LiteralPath $out).Length
                    $before += $sizeBefore; $after += $sizeAfter; $ok++
                    $pct = if ($sizeBefore -gt 0) { [Math]::Round(100 - ($sizeAfter / $sizeBefore * 100)) } else { 0 }
                    $colour = if ($sizeAfter -ge $sizeBefore) { 'Yellow' } else { 'Green' }
                    Write-Host ("  {0}  {1} -> {2} ({3}% smaller)" -f $f.Name, (Format-FileSize $sizeBefore), (Format-FileSize $sizeAfter), $pct) -ForegroundColor $colour
                } else {
                    $failed++
                    Write-Host "  FAILED: $($f.Name)" -ForegroundColor Red
                    $r.Lines | Select-Object -Last 2 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
                }
            }
            Write-Host ''
            if ($ok -gt 0) {
                $pct = if ($before -gt 0) { [Math]::Round(100 - ($after / $before * 100)) } else { 0 }
                Write-Host ("Total: {0} -> {1} ({2}% smaller). OK: {3}  Failed: {4}" -f (Format-FileSize $before), (Format-FileSize $after), $pct, $ok, $failed) -ForegroundColor Cyan
            }
            Write-Host 'Compression re-writes the PDF. Any OCR text layer is kept, but check one file before replacing originals.' -ForegroundColor DarkGray
            Open-OutputFolder $s['OutDir']
        }
    )
    Invoke-Steps -Title 'Compress PDFs' -State $state -Steps $steps
}

# ======================================================================
# SEARCH
# ======================================================================
function Invoke-SearchMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs to search:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            Write-Host 'Only searchable text is found. Scanned pages are reported separately.' -ForegroundColor DarkGray
            $term = Read-Text 'Text to find'
            if (-not $term) { Step-Back }
            $s['Term'] = $term
            Set-Recap $s "Looking for: $term"
        },
        {
            param($s)
            $list = New-ListFile @($s['Files'] | ForEach-Object { $_.FullName })
            $r = Invoke-Helper -Arguments @('search', $list, $s['Term']) -Quiet
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            if ($r.Code -ne 0) { $r.Lines | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }; return }
            $rows = @($r.Lines | Where-Object { $_ -match "`t" } | ForEach-Object {
                    $p = $_ -split "`t"; [pscustomobject]@{ File = $p[0]; Page = $p[1] } })
            $matches = @($rows | Where-Object { $_.Page -match '^\d+$' })
            $noText = @($rows | Where-Object { $_.Page -eq 'NOTEXT' })
            $locked = @($rows | Where-Object { $_.Page -eq 'LOCKED' })
            Write-Host ''
            if ($matches.Count -eq 0) { Write-Host '  No matches.' -ForegroundColor Yellow }
            foreach ($m in $matches) { Write-Host ('  page {0,4}  {1}' -f $m.Page, $m.File) }
            Write-Host ''
            Write-Host "  $($matches.Count) match(es) in $(@($matches | Group-Object File).Count) file(s)." -ForegroundColor Cyan
            if ($noText.Count -gt 0) {
                Write-Host "  $($noText.Count) file(s) have no searchable text and were not really searched:" -ForegroundColor Yellow
                $noText | ForEach-Object { Write-Host "    $($_.File)" -ForegroundColor Yellow }
                Write-Host '  OCR those first, then search again.' -ForegroundColor Yellow
            }
            if ($locked.Count -gt 0) {
                Write-Host "  $($locked.Count) file(s) are password protected and were skipped." -ForegroundColor Yellow
            }
            if ($matches.Count -gt 0) {
                Write-Host ''
                if (Read-YesNo 'Save these results as "Search results.csv"?' 'N') {
                    $csv = Join-Path $s['SourceFolder'] 'Search results.csv'
                    $matches | Select-Object File, Page | Export-Csv -LiteralPath $csv -NoTypeInformation
                    Write-Host "Saved: $csv" -ForegroundColor Green
                }
            }
        }
    )
    Invoke-Steps -Title 'Search text across PDFs' -State $state -Steps $steps
}

# ======================================================================
# COUNT / UNLOCK
# ======================================================================
function Invoke-CountMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            $list = New-ListFile @($s['Files'] | ForEach-Object { $_.FullName })
            $r = Invoke-Helper -Arguments @('count', $list) -Quiet
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            if ($r.Code -ne 0) { $r.Lines | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }; return }
            $rows = @($r.Lines | Where-Object { $_ -match "`t" } | ForEach-Object {
                    $p = $_ -split "`t"; [pscustomobject]@{ File = $p[0]; Pages = $p[1] } })
            $total = 0
            foreach ($row in $rows) {
                $colour = if ($row.Pages -match '^\d+$') { $total += [int]$row.Pages; 'Gray' } else { 'Yellow' }
                Write-Host ('  {0,6}  {1}' -f $row.Pages, $row.File) -ForegroundColor $colour
            }
            Write-Host ''
            Write-Host ('  {0,6}  TOTAL ({1} files)' -f $total, $rows.Count) -ForegroundColor Cyan
            if (@($rows | Where-Object { $_.Pages -notmatch '^\d+$' }).Count -gt 0) {
                Write-Host '  LOCKED = password protected, ERROR = could not be read. Not in the total.' -ForegroundColor Yellow
            }
            if ($rows.Count -gt 1) {
                Write-Host ''
                if (Read-YesNo 'Save this list as "Page count.csv" in the folder?' 'N') {
                    $csv = Join-Path $s['SourceFolder'] 'Page count.csv'
                    $rows | Export-Csv -LiteralPath $csv -NoTypeInformation
                    Write-Host "Saved: $csv" -ForegroundColor Green
                }
            }
        }
    )
    Invoke-Steps -Title 'Count pages' -State $state -Steps $steps
}

function Invoke-UnlockMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Write-Host 'You need to know the password. The original file is left untouched.' -ForegroundColor DarkGray
            Write-Host ''
            $file = Read-UserPath -Prompt 'Enter the password-protected PDF:' -Type File -Extensions @('.pdf')
            $info = Get-PdfInfo $file
            if ($null -eq $info) { Stop-Task }
            if (-not $info.Encrypted) { Write-Host 'This PDF is not password protected.' -ForegroundColor Yellow; Stop-Task }
            if ($null -eq (Get-PdfPages $file)) { Stop-Task }
            $s['File'] = $file
            $s['Item'] = Get-Item -LiteralPath $file
            Set-Recap $s "File: $(Split-Path $file -Leaf)"
        },
        {
            param($s)
            $outPath = Join-Path $s['Item'].DirectoryName "$($s['Item'].BaseName) unlocked.pdf"
            Write-Host 'Unlocked copy will be saved as:'
            Write-Host "  $outPath"
            Write-Host ''
            if (-not (Read-YesNo 'Continue?' 'Y')) { Step-Back }
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            $r = Invoke-Helper -Arguments @('unlock', $s['File'], $outPath)
            if ($r.Code -eq 0) { Write-Host ''; Write-Host "Saved: $outPath" -ForegroundColor Green }
        }
    )
    Invoke-Steps -Title 'Remove a PDF password' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: OFFICE TO PDF
# ======================================================================
function Convert-OfficeFiles {
    param([string]$App, $Files, [string]$OutDir, [hashtable]$Stats, [hashtable]$Used)
    if ($Files.Count -eq 0) { return }
    $missing = [Type]::Missing
    $noPw = 'pdftoolbox-no-password'   # makes protected files fail instead of waiting on a hidden prompt
    try {
        switch ($App) {
            'Word' { $com = New-Object -ComObject Word.Application; $com.Visible = $false; $com.DisplayAlerts = 0 }
            'Excel' { $com = New-Object -ComObject Excel.Application; $com.Visible = $false; $com.DisplayAlerts = $false }
            'PowerPoint' { $com = New-Object -ComObject PowerPoint.Application }
        }
    } catch {
        Write-Host "  $App is not installed or could not start. Skipped $($Files.Count) file(s)." -ForegroundColor Red
        $Stats.Failed += $Files.Count
        return
    }
    try {
        foreach ($f in $Files) {
            $out = Join-Path $OutDir ($f.BaseName + '.pdf')
            if ($Used.ContainsKey($out.ToLower())) {
                $out = Join-Path $OutDir ('{0} {1}.pdf' -f $f.BaseName, $f.Extension.TrimStart('.'))
            }
            $Used[$out.ToLower()] = $true
            $docObj = $null
            try {
                switch ($App) {
                    'Word' {
                        $docObj = $com.Documents.Open($f.FullName, $false, $true, $false, $noPw)
                        $docObj.ExportAsFixedFormat($out, 17)
                        $docObj.Close(0)
                    }
                    'Excel' {
                        $docObj = $com.Workbooks.Open($f.FullName, 0, $true, $missing, $noPw)
                        $docObj.ExportAsFixedFormat(0, $out)
                        $docObj.Close($false)
                    }
                    'PowerPoint' {
                        $docObj = $com.Presentations.Open($f.FullName, -1, 0, 0)
                        $docObj.SaveAs($out, 32)
                        $docObj.Close()
                    }
                }
                $Stats.OK++
                Write-Host "  Converted: $($f.Name)" -ForegroundColor Green
            } catch {
                $Stats.Failed++
                Write-Host "  FAILED:    $($f.Name) ($($_.Exception.Message))" -ForegroundColor Red
                if ($docObj) { try { $docObj.Close() } catch { } }
            } finally {
                if ($docObj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($docObj) }
            }
        }
    } finally {
        try { $com.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-OfficeToPdfMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a Word, Excel or PowerPoint file, or a folder:' -Extensions $script:OfficeExtensions -Kind 'Office'
            $s['Word'] = @($s['Files'] | Where-Object { @('.doc', '.docx') -contains $_.Extension.ToLower() })
            $s['Excel'] = @($s['Files'] | Where-Object { @('.xls', '.xlsx', '.xlsm') -contains $_.Extension.ToLower() })
            $s['Ppt'] = @($s['Files'] | Where-Object { @('.ppt', '.pptx') -contains $_.Extension.ToLower() })
        },
        {
            param($s)
            Write-Host "Found: $($s['Word'].Count) Word, $($s['Excel'].Count) Excel, $($s['Ppt'].Count) PowerPoint." -ForegroundColor Green
            if ($s['Excel'].Count) { Write-Host 'Excel files export every sheet, using each sheet''s print settings.' -ForegroundColor DarkGray }
            Write-Host ''
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'PDFs'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            if (-not (Read-YesNo "Convert $($s['Files'].Count) file(s) to PDF?" 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $s['OutDir'] | Out-Null
            Write-Host ''
            $stats = @{ OK = 0; Failed = 0 }
            $used = @{}
            Convert-OfficeFiles -App 'Word' -Files $s['Word'] -OutDir $s['OutDir'] -Stats $stats -Used $used
            Convert-OfficeFiles -App 'Excel' -Files $s['Excel'] -OutDir $s['OutDir'] -Stats $stats -Used $used
            Convert-OfficeFiles -App 'PowerPoint' -Files $s['Ppt'] -OutDir $s['OutDir'] -Stats $stats -Used $used
            Write-Host ''
            Write-Host ('OK: {0}   Failed: {1}' -f $stats.OK, $stats.Failed) -ForegroundColor $(if ($stats.Failed) { 'Red' } else { 'Green' })
            Open-OutputFolder $s['OutDir']
        }
    )
    Invoke-Steps -Title 'Convert Word / Excel / PowerPoint to PDF' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: PDF TO WORD (Word's own PDF reflow)
# ======================================================================
function Invoke-PdfToWordMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Write-Host 'Word rebuilds the document from the PDF. Simple text comes out well;' -ForegroundColor DarkGray
            Write-Host 'complex tables and columns often need tidying. Scanned PDFs must be OCR''d first.' -ForegroundColor DarkGray
            Write-Host ''
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'Word'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            if (-not (Read-YesNo "Convert $($s['Files'].Count) file(s) to Word?" 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $s['OutDir'] | Out-Null
            Write-Host ''
            try { $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0 }
            catch { Write-Host '  Word is not installed or could not start.' -ForegroundColor Red; return }
            $ok = 0; $failed = 0
            try {
                foreach ($f in $s['Files']) {
                    $out = Join-Path $s['OutDir'] ($f.BaseName + '.docx')
                    $doc = $null
                    try {
                        # ConfirmConversions = false stops Word's "convert this PDF" prompt.
                        $doc = $word.Documents.Open($f.FullName, $false, $true)
                        $doc.SaveAs2($out, 16)   # 16 = .docx
                        $doc.Close(0)
                        $ok++
                        Write-Host "  Converted: $($f.Name)" -ForegroundColor Green
                    } catch {
                        $failed++
                        Write-Host "  FAILED:    $($f.Name) ($($_.Exception.Message))" -ForegroundColor Red
                        if ($doc) { try { $doc.Close(0) } catch { } }
                    } finally {
                        if ($doc) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
                    }
                }
            } finally {
                try { $word.Quit() } catch { }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }
            Write-Host ''
            Write-Host ('OK: {0}   Failed: {1}' -f $ok, $failed) -ForegroundColor $(if ($failed) { 'Red' } else { 'Green' })
            Write-Host 'Check the result against the PDF before using it. Layout is rebuilt, not copied.' -ForegroundColor Yellow
            Open-OutputFolder $s['OutDir']
        }
    )
    Invoke-Steps -Title 'Convert PDF to Word' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: PDF TABLES TO EXCEL
# ======================================================================
function Invoke-PdfToExcelMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Write-Host 'Tables are detected from the PDF''s text layout and written to a workbook,' -ForegroundColor DarkGray
            Write-Host 'one sheet per table. Scanned PDFs give nothing until they are OCR''d.' -ForegroundColor DarkGray
            Write-Host ''
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'Excel'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            if (-not (Read-YesNo "Extract tables from $($s['Files'].Count) file(s)?" 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $s['OutDir'] | Out-Null
            Write-Host ''
            $ok = 0; $none = 0; $failed = 0
            foreach ($f in $s['Files']) {
                Write-Host "  $($f.Name)"
                $base = Join-Path $s['OutDir'] $f.BaseName
                $r = Invoke-Helper -Arguments @('tables', $f.FullName, $base) -Quiet
                if ($r.Code -eq 0) {
                    $ok++
                    $r.Lines | ForEach-Object { Write-Host "    $_" -ForegroundColor Green }
                } elseif ($r.Code -eq 2) {
                    $none++
                    Write-Host '    No tables found (scanned, or no table structure).' -ForegroundColor Yellow
                } else {
                    $failed++
                    $r.Lines | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
                }
            }
            Write-Host ''
            Write-Host ("Workbooks created: {0}   No tables: {1}   Errors: {2}" -f $ok, $none, $failed) -ForegroundColor Cyan
            Write-Host 'Table detection is a best guess. Tie the totals back to the PDF before using the figures.' -ForegroundColor Yellow
            Open-OutputFolder $s['OutDir']
        }
    )
    Invoke-Steps -Title 'Convert PDF tables to Excel' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: PDF TO TEXT
# ======================================================================
function Invoke-PdfToTextMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter a PDF file, or a folder of PDFs:' -Extensions @('.pdf') -Kind 'PDF'
        },
        {
            param($s)
            $s['OutDir'] = Select-OutputFolder -State $s -DefaultName 'Text'
            Set-Recap $s "Save to: $($s['OutDir'])"
        },
        {
            param($s)
            if (-not (Read-YesNo "Extract text from $($s['Files'].Count) file(s)?" 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $s['OutDir'] | Out-Null
            Write-Host ''
            $ok = 0; $failed = 0
            foreach ($f in $s['Files']) {
                $out = Join-Path $s['OutDir'] ($f.BaseName + '.txt')
                Write-Host "  $($f.Name)"
                $r = Invoke-Helper -Arguments @('text', $f.FullName, $out) -Quiet
                if ($r.Code -eq 0) {
                    $ok++
                    $r.Lines | ForEach-Object {
                        $colour = if ($_ -like 'WARNING:*') { 'Yellow' } else { 'Green' }
                        Write-Host "    $_" -ForegroundColor $colour
                    }
                } else {
                    $failed++
                    $r.Lines | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
                }
            }
            Write-Host ''
            Write-Host ("OK: {0}   Failed: {1}" -f $ok, $failed) -ForegroundColor Cyan
            Open-OutputFolder $s['OutDir']
        }
    )
    Invoke-Steps -Title 'Convert PDF to plain text' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: PDF TO IMAGES
# ======================================================================
function Invoke-PdfToImagesMenu {
    if (-not (Get-GhostscriptExe)) {
        Write-Header 'Convert PDF to images'
        Write-Host 'Ghostscript was not found. Run Install PDFToolbox.bat.' -ForegroundColor Red
        return
    }
    $state = @{}
    $steps = @(
        { param($s) Select-SinglePdf -State $s -Prompt 'Enter the PDF to turn into images:' },
        {
            param($s)
            Write-Host 'Image quality:'
            Write-Host '  1  150 dpi (screen, smaller files)'
            Write-Host '  2  300 dpi (print quality)'
            $choice = Read-Choice 'Choice' @('1', '2') '1'
            $s['Dpi'] = @{ '1' = 150; '2' = 300 }[$choice]
            Set-Recap $s "Quality: $($s['Dpi']) dpi"
        },
        {
            param($s)
            Write-Host 'Format:'
            Write-Host '  1  PNG (sharp text, larger files)'
            Write-Host '  2  JPEG (smaller files)'
            $choice = Read-Choice 'Choice' @('1', '2') '1'
            $s['Device'] = if ($choice -eq '1') { 'png16m' } else { 'jpeg' }
            $s['ImageExt'] = if ($choice -eq '1') { 'png' } else { 'jpg' }
            Set-Recap $s "Format: $($s['ImageExt'].ToUpper())"
        },
        {
            param($s)
            $outDir = Join-Path $s['Item'].DirectoryName "$($s['Item'].BaseName) images"
            Write-Host "$($s['Pages']) image(s) will be saved to:"
            Write-Host "  $outDir"
            Write-Host ''
            if (-not (Read-YesNo 'Continue?' 'Y')) { Step-Back }
            New-Item -ItemType Directory -Force -Path $outDir | Out-Null
            $pattern = Join-Path $outDir ("$($s['Item'].BaseName) p%03d." + $s['ImageExt'])
            Write-Host 'Working...'
            $r = Invoke-Ghostscript @("-sDEVICE=$($s['Device'])", "-r$($s['Dpi'])", "-sOutputFile=$pattern", $s['File'])
            $made = @(Get-ChildItem -LiteralPath $outDir -Filter "*.$($s['ImageExt'])" -File -ErrorAction SilentlyContinue)
            if ($r.Code -eq 0 -and $made.Count -gt 0) {
                Write-Host "Created $($made.Count) image(s)." -ForegroundColor Green
                Open-OutputFolder $outDir
            } else {
                Write-Host 'Ghostscript could not produce the images.' -ForegroundColor Red
                $r.Lines | Select-Object -Last 3 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
            }
        }
    )
    Invoke-Steps -Title 'Convert PDF to images' -State $state -Steps $steps
}

# ======================================================================
# CONVERT: IMAGES TO PDF
# ======================================================================
function Invoke-ImagesToPdfMenu {
    $state = @{}
    $steps = @(
        {
            param($s)
            Select-TargetFiles -State $s -Prompt 'Enter an image, or a folder of images:' -Extensions $script:ImageExtensions -Kind 'image'
        },
        {
            param($s)
            Write-Host 'Order:'
            Write-Host '  1  By file name'
            Write-Host '  2  By date taken or modified (oldest first)'
            $order = Read-Choice 'Choice' @('1', '2') '1'
            $s['Files'] = if ($order -eq '1') { @($s['Files'] | Sort-Object Name) } else { @($s['Files'] | Sort-Object LastWriteTime) }
            Write-Host ''
            Show-FileList $s['Files']
            Set-Recap $s ('Order: ' + $(if ($order -eq '1') { 'by file name' } else { 'by date' }))
        },
        {
            param($s)
            $name = Read-OutputName 'Scanned.pdf'
            $outPath = Join-Path $s['SourceFolder'] $name
            if (-not (Confirm-Overwrite $outPath)) { Step-Back }
            $s['OutPath'] = $outPath
            Set-Recap $s "Output: $name"
        },
        {
            param($s)
            if (-not (Read-YesNo "Make one PDF from $($s['Files'].Count) image(s)?" 'Y')) { Step-Back }
            Write-Host ''
            $list = New-ListFile @($s['Files'] | ForEach-Object { $_.FullName })
            $r = Invoke-Helper -Arguments @('imgtopdf', $list, $s['OutPath'])
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            if ($r.Code -eq 0) {
                Write-Host ''
                Write-Host "Saved: $($s['OutPath'])" -ForegroundColor Green
                Write-Host 'The pages are pictures. Run OCR on it to make the text searchable.' -ForegroundColor DarkGray
                Open-OutputFolder $s['SourceFolder']
            }
        }
    )
    Invoke-Steps -Title 'Convert images to one PDF' -State $state -Steps $steps
}

# ======================================================================
# VISUAL PAGE ORGANISER (WPF window, thumbnails rendered by Ghostscript)
# ======================================================================
$script:OrganiserXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PDF Toolbox" Height="740" Width="1060"
        WindowStartupLocation="CenterScreen" Background="#1A1A2E"
        FontFamily="Aptos Narrow, Aptos, Segoe UI" FontSize="13">
  <DockPanel Margin="12">
    <TextBlock x:Name="Info" DockPanel.Dock="Top" Foreground="#E6E9F0" FontSize="14"
               Margin="4,0,4,10" TextWrapping="Wrap"/>
    <Border DockPanel.Dock="Bottom" Background="#22223A" CornerRadius="6" Padding="8" Margin="0,10,0,0">
      <DockPanel>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
          <Button x:Name="PreviewButton" Content="Preview page" Width="120" Margin="2" Padding="4"/>
          <Button x:Name="AllButton" Content="Select all" Width="95" Margin="2" Padding="4"/>
          <Button x:Name="NoneButton" Content="Clear" Width="75" Margin="2" Padding="4"/>
          <Button x:Name="UpButton" Content="Move back" Width="95" Margin="2" Padding="4"/>
          <Button x:Name="DownButton" Content="Move forward" Width="110" Margin="2" Padding="4"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <TextBlock x:Name="Status" Foreground="#9AA3B8" VerticalAlignment="Center" Margin="0,0,14,0"/>
          <Button x:Name="OkButton" Content="Use this" Width="100" Margin="2" Padding="4" IsDefault="True"/>
          <Button x:Name="CancelButton" Content="Cancel" Width="80" Margin="2" Padding="4" IsCancel="True"/>
        </StackPanel>
      </DockPanel>
    </Border>
    <ListBox x:Name="List" Background="#141426" BorderBrush="#33334D" Foreground="#E6E9F0"
             ScrollViewer.HorizontalScrollBarVisibility="Disabled" SelectionMode="Extended">
      <ListBox.ItemsPanel>
        <ItemsPanelTemplate><WrapPanel Orientation="Horizontal"/></ItemsPanelTemplate>
      </ListBox.ItemsPanel>
      <ListBox.ItemTemplate>
        <DataTemplate>
          <StackPanel Width="152" Margin="6">
            <Border Background="White" Padding="3" CornerRadius="3">
              <Image Source="{Binding Thumb}" Height="185" Stretch="Uniform"/>
            </Border>
            <TextBlock Text="{Binding Label}" Foreground="#E6E9F0" HorizontalAlignment="Center"
                       Margin="0,6,0,0" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
        </DataTemplate>
      </ListBox.ItemTemplate>
    </ListBox>
  </DockPanel>
</Window>
'@

function Test-CanShowWindow {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        Write-Host '  The visual view needs an STA window. Type the page numbers instead.' -ForegroundColor Yellow
        return $false
    }
    try { Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction Stop; return $true }
    catch {
        Write-Host '  The visual view is not available on this machine. Type the page numbers instead.' -ForegroundColor Yellow
        return $false
    }
}

# WPF cannot bind Image.Source from a PSObject property, so grid items use a
# real .NET class with a typed ImageSource.
function Initialize-OrganiserItemType {
    if ('PdfToolbox.OrganiserItem' -as [type]) { return $true }
    try {
        Add-Type -ReferencedAssemblies PresentationCore, WindowsBase -TypeDefinition @'
using System.Windows.Media;
namespace PdfToolbox
{
    public class OrganiserItem
    {
        public string Label { get; set; }
        public ImageSource Thumb { get; set; }
        public object Tag { get; set; }
    }
}
'@ -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function New-OrganiserItem {
    param([string]$Label, $Thumb, $Tag)
    if (Initialize-OrganiserItemType) {
        $item = New-Object PdfToolbox.OrganiserItem
        $item.Label = $Label
        $item.Thumb = $Thumb
        $item.Tag = $Tag
        return $item
    }
    return [pscustomobject]@{ Label = $Label; Thumb = $Thumb; Tag = $Tag }
}

# Read the PNG into memory rather than pointing at the file: no file lock, and
# no URI handling to go wrong.
# DecodeWidth caps the decoded size, which keeps a grid of thumbnails light.
# Pass 0 for a full-resolution image, as the page preview needs.
function New-ThumbImage {
    param([string]$Path, [int]$DecodeWidth = 240)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -eq 0) { return $null }
        $stream = New-Object System.IO.MemoryStream (, $bytes)
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.StreamSource = $stream
        if ($DecodeWidth -gt 0) { $bitmap.DecodePixelWidth = $DecodeWidth }
        $bitmap.EndInit()
        $bitmap.Freeze()
        return $bitmap
    } catch {
        Write-Host "  Preview could not be loaded: $($_.Exception.Message)" -ForegroundColor DarkGray
        return $null
    }
}

function New-PdfThumbnails {
    param([string]$PdfPath, [int]$Dpi = 42, [int]$FirstPage = 0, [int]$LastPage = 0)
    $dir = Join-Path $script:WorkDir ('thumbs' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $arguments = @('-sDEVICE=png16m', "-r$Dpi", '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4')
    if ($FirstPage -gt 0) { $arguments += "-dFirstPage=$FirstPage" }
    if ($LastPage -gt 0) { $arguments += "-dLastPage=$LastPage" }
    $arguments += @(('-sOutputFile=' + (Join-Path $dir 'p%04d.png')), $PdfPath)
    $result = Invoke-Ghostscript $arguments
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.png' -File -ErrorAction SilentlyContinue | Sort-Object Name)
    return [pscustomobject]@{ Dir = $dir; Files = $files; Code = $result.Code }
}

function Remove-ThumbFolder {
    param([string]$Dir)
    try { Remove-Item -LiteralPath $Dir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}

# ----------------------------------------------------------------------
# UPDATE CHECK
# Reads a small JSON file from the web and compares it with this version.
# It never installs anything: it only reports and points at the download.
# ----------------------------------------------------------------------
function Get-LatestVersionInfo {
    if ($script:UpdateManifestUrl -like '*GITHUBUSER*') {
        return [pscustomobject]@{ Ok = $false; Message = 'No update location has been set up yet.' }
    }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'
    try {
        # PowerShell 5.1 can still default to old TLS, which GitHub refuses.
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
        $parameters = @{ Uri = $script:UpdateManifestUrl; TimeoutSec = 20 }
        # Go through the company proxy with the signed-in user's credentials.
        try {
            $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
            $proxyUri = $proxy.GetProxy($script:UpdateManifestUrl)
            if ($proxyUri -and $proxyUri.AbsoluteUri -ne $script:UpdateManifestUrl) {
                $parameters['Proxy'] = $proxyUri.AbsoluteUri
                $parameters['ProxyUseDefaultCredentials'] = $true
            }
        } catch { }
        $manifest = Invoke-RestMethod @parameters
        if (-not $manifest.version) { throw 'The update file did not contain a version.' }
        $latest = [version]("$($manifest.version)".Trim())
        $current = [version]$script:Version
        return [pscustomobject]@{
            Ok        = $true
            Newer     = ($latest -gt $current)
            Latest    = "$($manifest.version)".Trim()
            Released  = "$($manifest.released)".Trim()
            Notes     = "$($manifest.notes)".Trim()
            Download  = $(if ($manifest.download) { "$($manifest.download)".Trim() } else { $script:UpdateDownloadUrl })
            Message   = ''
        }
    } catch {
        return [pscustomobject]@{ Ok = $false; Message = $_.Exception.Message }
    } finally {
        $ErrorActionPreference = $previous
    }
}

# ----------------------------------------------------------------------
# PAGE PREVIEW
# Renders one page at a readable size and shows it in its own window, with
# Previous and Next so a document can be flicked through.
# ----------------------------------------------------------------------
$script:PreviewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preview" Height="940" Width="820" WindowStartupLocation="CenterScreen"
        Background="#1A1A2E" FontFamily="Aptos Narrow, Aptos, Segoe UI" FontSize="13">
  <DockPanel Margin="12" TextElement.Foreground="#E9ECF5">
    <TextBlock x:Name="PreviewTitle" DockPanel.Dock="Top" FontWeight="Bold" Margin="4,0,4,8" TextTrimming="CharacterEllipsis"/>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
      <StackPanel Orientation="Horizontal">
        <Button x:Name="PrevButton" Content="Previous" Width="100" Margin="3" Padding="6,4"/>
        <Button x:Name="NextButton" Content="Next" Width="100" Margin="3" Padding="6,4"/>
        <Button x:Name="ZoomOutButton" Content="Zoom out" Width="90" Margin="12,3,3,3" Padding="6,4"/>
        <Button x:Name="ZoomInButton" Content="Zoom in" Width="90" Margin="3" Padding="6,4"/>
        <Button x:Name="FitButton" Content="Fit page" Width="90" Margin="3" Padding="6,4"/>
        <Button x:Name="ActualButton" Content="100%" Width="70" Margin="3" Padding="6,4"/>
        <TextBlock x:Name="ZoomLabel" VerticalAlignment="Center" Margin="10,0,0,0" Foreground="#9AA3B8"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <TextBlock x:Name="PreviewStatus" VerticalAlignment="Center" Margin="0,0,14,0" Foreground="#9AA3B8"/>
        <Button x:Name="CloseButton" Content="Close" Width="100" Margin="3" Padding="6,4" IsCancel="True"/>
      </StackPanel>
    </DockPanel>
    <ScrollViewer x:Name="PreviewScroll" VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Auto" Background="#141426">
      <Border Background="White" Margin="10" Padding="4" HorizontalAlignment="Center" VerticalAlignment="Center">
        <Image x:Name="PreviewImage" Stretch="Uniform"
               RenderOptions.BitmapScalingMode="HighQuality"
               SnapsToDevicePixels="True"/>
      </Border>
    </ScrollViewer>
  </DockPanel>
</Window>
'@

function Show-PagePreview {
    param([string]$PdfPath, [int]$Page = 1, [string]$Title, [int]$TotalPages = 0)
    if (-not (Test-CanShowWindow)) { return }
    if (-not (Get-GhostscriptExe)) {
        Write-Host '  Ghostscript is needed to render a preview.' -ForegroundColor Yellow
        return
    }
    if ($TotalPages -lt 1) {
        $info = Get-PdfInfo $PdfPath
        $TotalPages = $(if ($info -and $info.Pages -gt 0) { $info.Pages } else { 1 })
    }
    if (-not $Title) { $Title = Split-Path $PdfPath -Leaf }

    try {
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$script:PreviewXaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
    } catch { return }

    $image = $window.FindName('PreviewImage')
    $titleBlock = $window.FindName('PreviewTitle')
    $status = $window.FindName('PreviewStatus')
    $prev = $window.FindName('PrevButton')
    $next = $window.FindName('NextButton')
    $close = $window.FindName('CloseButton')
    $scroll = $window.FindName('PreviewScroll')
    $zoomIn = $window.FindName('ZoomInButton')
    $zoomOut = $window.FindName('ZoomOutButton')
    $fitButton = $window.FindName('FitButton')
    $actualButton = $window.FindName('ActualButton')
    $zoomLabel = $window.FindName('ZoomLabel')
    $titleBlock.Text = $Title
    # The page number lives on an object: assigning to a plain variable inside
    # a click handler would only change a copy, leaving the viewer stuck.
    # Everything a handler changes lives on this object: assigning to a plain
    # variable inside a handler would only change a copy.
    $state = [pscustomobject]@{
        Page = [Math]::Max(1, [Math]::Min($Page, $TotalPages))
        Zoom = 1.0
        Fit  = $true
    }
    $folders = New-Object System.Collections.ArrayList
    $renderDpi = 200

    # 100% means the page at its real paper size: the render is 200 dpi and
    # WPF works in 96ths of an inch.
    $applyZoom = {
        if (-not $image.Source) { return }
        if ($state.Fit) {
            $scroll.HorizontalScrollBarVisibility = 'Disabled'
            $image.Width = [double]::NaN
            $image.Height = [double]::NaN
            $zoomLabel.Text = 'Fit page'
        } else {
            $scroll.HorizontalScrollBarVisibility = 'Auto'
            $naturalWidth = $image.Source.PixelWidth * (96.0 / $renderDpi)
            $image.Width = $naturalWidth * $state.Zoom
            $image.Height = [double]::NaN
            $zoomLabel.Text = "$([Math]::Round($state.Zoom * 100))%"
        }
    }

    $setZoom = {
        param($factor)
        $state.Fit = $false
        $state.Zoom = [Math]::Max(0.25, [Math]::Min(5.0, $factor))
        & $applyZoom
    }

    $render = {
        $status.Text = 'Rendering...'
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $thumbs = New-PdfThumbnails -PdfPath $PdfPath -Dpi 200 -FirstPage $state.Page -LastPage $state.Page
            [void]$folders.Add($thumbs.Dir)
            if ($thumbs.Files.Count -gt 0) {
                $bitmap = New-ThumbImage -Path $thumbs.Files[0].FullName -DecodeWidth 0
                if ($bitmap) { $image.Source = $bitmap }
            }
            $status.Text = "Page $($state.Page) of $TotalPages"
        } catch {
            $status.Text = "Could not render page $($state.Page)"
        } finally {
            $window.Cursor = $null
        }
        $prev.IsEnabled = ($state.Page -gt 1)
        $next.IsEnabled = ($state.Page -lt $TotalPages)
        & $applyZoom
    }

    $prev.Add_Click({ if ($state.Page -gt 1) { $state.Page = $state.Page - 1; & $render } })
    $next.Add_Click({ if ($state.Page -lt $TotalPages) { $state.Page = $state.Page + 1; & $render } })
    $close.Add_Click({ $window.Close() })
    $zoomIn.Add_Click({ & $setZoom ($(if ($state.Fit) { 1.0 } else { $state.Zoom }) * 1.25) })
    $zoomOut.Add_Click({ & $setZoom ($(if ($state.Fit) { 1.0 } else { $state.Zoom }) / 1.25) })
    $actualButton.Add_Click({ & $setZoom 1.0 })
    $fitButton.Add_Click({ $state.Fit = $true; & $applyZoom })

    # Ctrl and the mouse wheel zooms, like every other viewer.
    $scroll.Add_PreviewMouseWheel({
            param($sender, $e)
            if ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) {
                $base = $(if ($state.Fit) { 1.0 } else { $state.Zoom })
                & $setZoom $(if ($e.Delta -gt 0) { $base * 1.25 } else { $base / 1.25 })
                $e.Handled = $true
            }
        })

    & $render
    [void]$window.ShowDialog()
    foreach ($folder in $folders) { Remove-ThumbFolder $folder }
}

# ----------------------------------------------------------------------
# DRAG TO REORDER
# Handlers fire after the registering function has returned, so each one is
# built with .GetNewClosure() to keep hold of its collection. Drag state is
# global because a closure gets its own $script: scope.
# ----------------------------------------------------------------------
function Get-RowContainerAt {
    param($Control, $Point)
    try {
        $hit = [System.Windows.Media.VisualTreeHelper]::HitTest($Control, $Point)
        if (-not $hit) { return $null }
        $element = $hit.VisualHit
        while ($null -ne $element) {
            if ($element -is [System.Windows.Controls.ListBoxItem]) { return $element }
            $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
        }
    } catch { }
    return $null
}

# $OnExternalFiles keeps its own scope, so it can still see the caller's
# variables and functions. It receives the dropped paths as one argument.
function Enable-DragReorder {
    param($Control, $Collection, [scriptblock]$OnExternalFiles)
    $Control.AllowDrop = $true

    $Control.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            $global:PdfTbDragStart = $e.GetPosition($sender)
            $global:PdfTbDragItem = $null
            $container = Get-RowContainerAt -Control $sender -Point $global:PdfTbDragStart
            if ($container) { $global:PdfTbDragItem = $container.DataContext }
        })

    $Control.Add_PreviewMouseMove({
            param($sender, $e)
            if ($e.LeftButton -ne 'Pressed' -or $null -eq $global:PdfTbDragItem) { return }
            $position = $e.GetPosition($sender)
            $dx = [Math]::Abs($position.X - $global:PdfTbDragStart.X)
            $dy = [Math]::Abs($position.Y - $global:PdfTbDragStart.Y)
            if ($dx -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
                $dy -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }
            try {
                $payload = New-Object System.Windows.DataObject 'PdfToolboxRow', $global:PdfTbDragItem
                [void][System.Windows.DragDrop]::DoDragDrop($sender, $payload, [System.Windows.DragDropEffects]::Move)
            } catch { }
            $global:PdfTbDragItem = $null
        })

    $Control.Add_DragOver({
            param($sender, $e)
            if ($e.Data.GetDataPresent('PdfToolboxRow')) { $e.Effects = 'Move' }
            elseif ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { $e.Effects = 'Copy' }
            else { $e.Effects = 'None' }
            $e.Handled = $true
        }.GetNewClosure())

    $Control.Add_Drop({
            param($sender, $e)
            try {
                if ($e.Data.GetDataPresent('PdfToolboxRow')) {
                    $dragged = $e.Data.GetData('PdfToolboxRow')
                    $from = $Collection.IndexOf($dragged)
                    if ($from -lt 0) { return }
                    $container = Get-RowContainerAt -Control $sender -Point $e.GetPosition($sender)
                    $to = $(if ($container) { $Collection.IndexOf($container.DataContext) } else { $Collection.Count - 1 })
                    if ($to -lt 0) { $to = $Collection.Count - 1 }
                    if ($to -ne $from) { $Collection.Move($from, $to) }
                    $sender.SelectedIndex = $to
                } elseif ($null -ne $OnExternalFiles -and $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                    & $OnExternalFiles @($e.Data.GetData([System.Windows.DataFormats]::FileDrop))
                }
            } catch { }
            $global:PdfTbDragItem = $null
            $e.Handled = $true
        }.GetNewClosure())
}

# Mode 'Select' returns the tags of the ticked items.
# Mode 'Order'  returns every tag, in the order the user arranged them.
function Show-Organiser {
    param([string]$Instruction, $Items, [ValidateSet('Select', 'Order')][string]$Mode, [scriptblock]$OnPreview)
    if (-not (Test-CanShowWindow)) { return $null }
    [void](Initialize-OrganiserItemType)
    $script:OrganiserResult = $null
    try {
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$script:OrganiserXaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
    } catch {
        Write-Host "  Could not open the window: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
    $list = $window.FindName('List')
    $info = $window.FindName('Info')
    $status = $window.FindName('Status')
    $okButton = $window.FindName('OkButton')
    $cancelButton = $window.FindName('CancelButton')
    $allButton = $window.FindName('AllButton')
    $noneButton = $window.FindName('NoneButton')
    $upButton = $window.FindName('UpButton')
    $downButton = $window.FindName('DownButton')
    $previewButton = $window.FindName('PreviewButton')

    # Preview whatever is highlighted. Double-clicking an item does the same.
    if ($null -eq $OnPreview) {
        $previewButton.Visibility = 'Collapsed'
    } else {
        $previewButton.Add_Click({
                $chosen = $list.SelectedItem
                if ($null -eq $chosen) { return }
                & $OnPreview $chosen.Tag
            })
        $list.Add_MouseDoubleClick({
                $chosen = $list.SelectedItem
                if ($null -ne $chosen) { & $OnPreview $chosen.Tag }
            })
    }

    $collection = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($item in $Items) { $collection.Add($item) }
    $list.ItemsSource = $collection
    $info.Text = $Instruction

    if ($Mode -eq 'Select') {
        $upButton.Visibility = 'Collapsed'
        $downButton.Visibility = 'Collapsed'
        $status.Text = '0 selected'
        $list.Add_SelectionChanged({ $status.Text = "$($list.SelectedItems.Count) selected" })
        $allButton.Add_Click({ $list.SelectAll() })
        $noneButton.Add_Click({ $list.UnselectAll() })
    } else {
        $allButton.Visibility = 'Collapsed'
        $noneButton.Visibility = 'Collapsed'
        $list.SelectionMode = 'Single'
        $status.Text = "$($collection.Count) item(s)"
        Enable-DragReorder -Control $list -Collection $collection
        $upButton.Add_Click({
                $i = $list.SelectedIndex
                if ($i -gt 0) { $collection.Move($i, $i - 1); $list.SelectedIndex = $i - 1 }
            })
        $downButton.Add_Click({
                $i = $list.SelectedIndex
                if ($i -ge 0 -and $i -lt $collection.Count - 1) { $collection.Move($i, $i + 1); $list.SelectedIndex = $i + 1 }
            })
    }

    $okButton.Add_Click({
            if ($Mode -eq 'Select') { $script:OrganiserResult = @($list.SelectedItems | ForEach-Object { $_.Tag }) }
            else { $script:OrganiserResult = @($collection | ForEach-Object { $_.Tag }) }
            $window.DialogResult = $true
            $window.Close()
        })
    $cancelButton.Add_Click({ $script:OrganiserResult = $null; $window.Close() })

    [void]$window.ShowDialog()
    return $script:OrganiserResult
}

# 1,2,3,7 -> "1-3,7"
function ConvertTo-PageSpec {
    param([int[]]$Pages)
    $sorted = @($Pages | Sort-Object -Unique)
    if ($sorted.Count -eq 0) { return '' }
    $parts = @()
    $start = $sorted[0]
    $previous = $sorted[0]
    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $page = $sorted[$i]
        if ($page -eq $previous + 1) { $previous = $page; continue }
        $parts += $(if ($start -eq $previous) { "$start" } else { "$start-$previous" })
        $start = $page
        $previous = $page
    }
    $parts += $(if ($start -eq $previous) { "$start" } else { "$start-$previous" })
    return ($parts -join ',')
}

function Select-PagesVisually {
    param([string]$PdfPath, [string]$Instruction)
    if (-not (Test-CanShowWindow)) { return $null }
    if (-not (Get-GhostscriptExe)) {
        Write-Host '  Ghostscript is needed for page previews. Type the page numbers instead.' -ForegroundColor Yellow
        return $null
    }
    Write-Host '  Rendering page previews, one moment...' -ForegroundColor DarkGray
    $thumbs = New-PdfThumbnails -PdfPath $PdfPath
    if ($thumbs.Files.Count -eq 0) {
        Remove-ThumbFolder $thumbs.Dir
        Write-Host '  Could not render previews for this PDF. Type the page numbers instead.' -ForegroundColor Yellow
        return $null
    }
    $items = @()
    $blank = 0
    for ($i = 0; $i -lt $thumbs.Files.Count; $i++) {
        $image = New-ThumbImage $thumbs.Files[$i].FullName
        if (-not $image) { $blank++ }
        $items += New-OrganiserItem -Label "Page $($i + 1)" -Thumb $image -Tag ($i + 1)
    }
    if ($blank -gt 0) { Write-Host "  $blank preview(s) could not be loaded from: $($thumbs.Dir)" -ForegroundColor Yellow }
    $picked = Show-Organiser -Instruction $Instruction -Items $items -Mode 'Select' -OnPreview {
        param($page)
        Show-PagePreview -PdfPath $PdfPath -Page $page -Title (Split-Path $PdfPath -Leaf) -TotalPages $thumbs.Files.Count
    }
    if ($blank -eq 0) { Remove-ThumbFolder $thumbs.Dir }
    if (-not $picked -or $picked.Count -eq 0) { return $null }
    return (ConvertTo-PageSpec ([int[]]$picked))
}

function Get-PageOrderVisually {
    param([string]$PdfPath)
    if (-not (Test-CanShowWindow)) { return $null }
    if (-not (Get-GhostscriptExe)) {
        Write-Host '  Ghostscript is needed for page previews.' -ForegroundColor Yellow
        return $null
    }
    Write-Host '  Rendering page previews, one moment...' -ForegroundColor DarkGray
    $thumbs = New-PdfThumbnails -PdfPath $PdfPath
    if ($thumbs.Files.Count -eq 0) {
        Remove-ThumbFolder $thumbs.Dir
        Write-Host '  Could not render previews for this PDF.' -ForegroundColor Yellow
        return $null
    }
    $items = @()
    for ($i = 0; $i -lt $thumbs.Files.Count; $i++) {
        $items += New-OrganiserItem -Label "Page $($i + 1)" -Thumb (New-ThumbImage $thumbs.Files[$i].FullName) -Tag ($i + 1)
    }
    $order = Show-Organiser -Mode 'Order' -Items $items -OnPreview {
        param($page)
        Show-PagePreview -PdfPath $PdfPath -Page $page -Title (Split-Path $PdfPath -Leaf) -TotalPages $thumbs.Files.Count
    } -Instruction 'Drag a page where you want it, or click it and use Move back and Move forward. Double-click a page to see it full size. Click Use this when the order is right.'
    Remove-ThumbFolder $thumbs.Dir
    if (-not $order -or $order.Count -eq 0) { return $null }
    return ,([int[]]$order)
}

function Get-FileOrderVisually {
    param($Files)
    if (-not (Test-CanShowWindow)) { return $null }
    if (-not (Get-GhostscriptExe)) {
        Write-Host '  Ghostscript is needed for previews.' -ForegroundColor Yellow
        return $null
    }
    Write-Host "  Rendering the first page of $($Files.Count) file(s), one moment..." -ForegroundColor DarkGray
    $items = @()
    $dirs = @()
    foreach ($file in $Files) {
        $thumbs = New-PdfThumbnails -PdfPath $file.FullName -FirstPage 1 -LastPage 1
        $dirs += $thumbs.Dir
        $thumb = if ($thumbs.Files.Count -gt 0) { New-ThumbImage $thumbs.Files[0].FullName } else { $null }
        $items += New-OrganiserItem -Label $file.Name -Thumb $thumb -Tag $file
    }
    $order = Show-Organiser -Mode 'Order' -Items $items -OnPreview {
        param($file)
        Show-PagePreview -PdfPath $file.FullName -Page 1 -Title $file.Name
    } -Instruction 'Drag a file where you want it, or click it and use Move back and Move forward. Double-click a file to page through it. Click Use this when the order is right.'
    foreach ($dir in $dirs) { Remove-ThumbFolder $dir }
    if (-not $order -or $order.Count -eq 0) { return $null }
    return @($order)
}

# ======================================================================
# SUBMENUS
# ======================================================================
function Show-SubMenu {
    param([string]$Title, [array]$Items)
    # Items: ordered array of @{ Key; Text; Action }
    while ($true) {
        Write-Header $Title
        foreach ($item in $Items) { Write-Host ('   {0}  {1}' -f $item.Key, $item.Text) }
        Write-Host '   0  Back to the main menu'
        Write-Host ''
        $valid = @($Items | ForEach-Object { $_.Key }) + @('0')
        try { $choice = Read-Choice 'Choose an option' $valid }
        catch {
            if ($_.Exception.Message -eq 'PDFTB_BACK' -or $_.Exception.Message -eq 'PDFTB_MENU') { return }
            throw
        }
        if ($choice -eq '0') { return }
        $action = ($Items | Where-Object { $_.Key -eq $choice } | Select-Object -First 1).Action
        try { & $action }
        catch {
            $message = $_.Exception.Message
            if ($message -ne 'PDFTB_BACK' -and $message -ne 'PDFTB_MENU') {
                Write-Host ''
                Write-Host "Something went wrong: $message" -ForegroundColor Red
                Wait-ForEnter
                continue
            }
        } finally { Clear-PdfPassword }
        Wait-ForEnter
    }
}

function Invoke-PagesSubMenu {
    Show-SubMenu -Title 'Split, extract or remove pages' -Items @(
        @{ Key = '1'; Text = 'Split into single pages'; Action = { Invoke-SplitMenu } },
        @{ Key = '2'; Text = 'Extract pages into one new PDF'; Action = { Invoke-ExtractMenu } },
        @{ Key = '3'; Text = 'Split into page ranges or equal chunks'; Action = { Invoke-RangesMenu } },
        @{ Key = '4'; Text = 'Delete pages'; Action = { Invoke-DeletePagesMenu } },
        @{ Key = '5'; Text = 'Rotate pages'; Action = { Invoke-RotateMenu } },
        @{ Key = '6'; Text = 'Insert pages or an image into a PDF'; Action = { Invoke-InsertMenu } },
        @{ Key = '7'; Text = 'Rearrange pages (visual page view)'; Action = { Invoke-ReorderMenu } }
    )
}

function Invoke-ConvertSubMenu {
    Show-SubMenu -Title 'Convert' -Items @(
        @{ Key = '1'; Text = 'Word / Excel / PowerPoint  ->  PDF'; Action = { Invoke-OfficeToPdfMenu } },
        @{ Key = '2'; Text = 'PDF  ->  Word (editable, layout rebuilt)'; Action = { Invoke-PdfToWordMenu } },
        @{ Key = '3'; Text = 'PDF  ->  Excel (tables only)'; Action = { Invoke-PdfToExcelMenu } },
        @{ Key = '4'; Text = 'PDF  ->  plain text'; Action = { Invoke-PdfToTextMenu } },
        @{ Key = '5'; Text = 'PDF  ->  images (PNG or JPEG)'; Action = { Invoke-PdfToImagesMenu } },
        @{ Key = '6'; Text = 'Images  ->  one PDF'; Action = { Invoke-ImagesToPdfMenu } }
    )
}

# ======================================================================
# SETUP CHECK
# ======================================================================
function Write-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail)
    $tag = if ($Ok) { '  OK      ' } else { '  MISSING ' }
    Write-Host $tag -NoNewline -ForegroundColor $(if ($Ok) { 'Green' } else { 'Red' })
    Write-Host ('{0,-14} {1}' -f $Name, $Detail)
}

function Test-PyModule {
    param([string]$Module)
    if (-not $script:PyExe) { return $false }
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $script:PyExe -c "import $Module" *> $null; return ($LASTEXITCODE -eq 0) }
    finally { $ErrorActionPreference = $previous }
}

function Invoke-SetupCheck {
    Write-Header 'Check setup'
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $pyVer = if ($script:PyExe) { "$(& $script:PyExe --version 2>&1)" } else { '' }
        $ocr = Get-OcrCommand
        $tess = Get-Command tesseract -ErrorAction SilentlyContinue
        $gs = Get-GhostscriptExe
    } finally { $ErrorActionPreference = $previous }
    Write-Check 'Python' ([bool]$script:PyExe) $pyVer
    Write-Check 'pypdf' (Test-PyModule 'pypdf') 'merge, split, extract, delete, rotate, count'
    Write-Check 'cryptography' (Test-PyModule 'cryptography') 'password-protected PDFs'
    Write-Check 'OCRmyPDF' ([bool]$ocr) 'OCR'
    Write-Check 'Tesseract' ([bool]$tess) $(if ($tess) { $tess.Source } else { '' })
    Write-Check 'Ghostscript' ([bool]$gs) $(if ($gs) { $gs } else { 'compress, PDF to images' })
    Write-Check 'pdfplumber' (Test-PyModule 'pdfplumber') 'PDF tables to Excel'
    Write-Check 'openpyxl' (Test-PyModule 'openpyxl') 'writes .xlsx (without it, CSV is used)'
    Write-Check 'Pillow' (Test-PyModule 'PIL') 'images to PDF'
    Write-Check 'Word' (Test-Path 'Registry::HKEY_CLASSES_ROOT\Word.Application') 'Office to PDF, PDF to Word'
    Write-Check 'Excel' (Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application') ''
    Write-Check 'PowerPoint' (Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application') ''
    Write-Host ''
    Write-Host 'Anything missing: close the toolbox and run "Install PDFToolbox.bat".' -ForegroundColor Cyan
}

# ======================================================================
# START-UP AND MAIN MENU
# ======================================================================
function Initialize-Toolbox {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
    $env:PYTHONIOENCODING = 'utf-8'
    New-Item -ItemType Directory -Force -Path $script:WorkDir | Out-Null
    $script:WorkDir = Expand-LongPath $script:WorkDir
    $script:Helper = Join-Path $script:WorkDir 'pdftoolboxhelper.py'
    Set-Content -LiteralPath $script:Helper -Value $script:HelperCode -Encoding ASCII

    if (Get-Command py -ErrorAction SilentlyContinue) { $script:PyExe = 'py' }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { $script:PyExe = 'python' }

    # If Tesseract / Ghostscript are installed but not on PATH, add them for this session.
    if (-not (Get-Command tesseract -ErrorAction SilentlyContinue)) {
        $t = 'C:\Program Files\Tesseract-OCR'
        if (Test-Path -LiteralPath (Join-Path $t 'tesseract.exe')) { $env:Path += ";$t" }
    }
    if (-not (Get-Command gswin64c -ErrorAction SilentlyContinue)) {
        $g = Get-ChildItem 'C:\Program Files\gs' -Recurse -Filter gswin64c.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($g) { $env:Path += ";$($g.DirectoryName)" }
    }
    Get-ToolboxSettings
}

function Invoke-About {
    Write-Header 'About' -Subtitle "Built by $script:Author"
    Write-Host "  Version $script:Version, released $script:ReleaseDate"
    Write-Host "  $script:Author    $script:AuthorEmail"
    Write-Host ''
    Write-Host '  WHAT IT DOES' -ForegroundColor Cyan
    Write-Host '  OCR, merge, split, extract, delete, rotate, compress, search and'
    Write-Host '  count PDFs, remove known passwords, and convert to and from PDF.'
    Write-Host '  Built for audit work: scanned bank statements, working papers and'
    Write-Host '  supporting documents.'
    Write-Host ''
    Write-Host '  WHAT IT RUNS ON' -ForegroundColor Cyan
    Write-Host '  OCRmyPDF, Tesseract and Ghostscript do the OCR and compression.'
    Write-Host '  pypdf, pdfplumber, openpyxl and Pillow handle the page and table work.'
    Write-Host '  Microsoft Word, Excel and PowerPoint do the Office conversions.'
    Write-Host '  All open-source except Office. Everything runs on this machine;'
    Write-Host '  no file is uploaded anywhere.'
    Write-Host ''
    Write-Host '  KNOW THIS BEFORE YOU TRUST THE OUTPUT' -ForegroundColor Yellow
    Write-Host '  - A scanned PDF has no text until you OCR it. Text extraction,'
    Write-Host '    table extraction and search all return nothing until then.'
    Write-Host '  - PDF to Excel guesses where the columns are. Tie the totals back'
    Write-Host '    to the PDF before using any figure.'
    Write-Host '  - PDF to Word rebuilds the layout, it does not copy it. Complex'
    Write-Host '    tables and columns usually need tidying.'
    Write-Host '  - OCR is not perfect. Spot-check figures on anything that matters.'
    Write-Host '  - Compressing re-writes the PDF. Check one file before you replace'
    Write-Host '    originals.'
    Write-Host '  - Every task writes a new file. Originals are left alone.'
    Write-Host ''
    Write-Host '  PROBLEMS' -ForegroundColor Cyan
    Write-Host '  Run option 9 (Check setup) first: most faults are a missing install.'
    Write-Host "  Still stuck, or want something added: $script:AuthorEmail"
    Write-Host "  Quote the version ($script:Version) and what you were doing."
    if ($script:UpdateDownloadUrl -notlike '*GITHUBUSER*') {
        Write-Host ''
        Write-Host '  UPDATES' -ForegroundColor Cyan
        Write-Host "  Newer versions are published at: $script:UpdateDownloadUrl"
    }
}

function Show-MainMenu {
    Write-Header 'PDF Toolbox' -Subtitle "Built by $script:Author"
    if ($script:LastFolder) { Write-Host "  Last folder: $($script:LastFolder)" -ForegroundColor DarkGray; Write-Host '' }
    Write-Host '   1  OCR PDFs (make scanned PDFs searchable)'
    Write-Host '   2  Merge PDFs in a folder'
    Write-Host '   3  Pages: split, extract, delete, rotate, insert, rearrange'
    Write-Host '   4  Convert (to PDF and from PDF)'
    Write-Host '   5  Compress PDFs'
    Write-Host '   6  Search text across PDFs'
    Write-Host '   7  Count pages'
    Write-Host '   8  Remove a PDF password'
    Write-Host '   9  Check setup'
    Write-Host '  10  About this tool'
    Write-Host '   0  Exit'
    Write-Host ''
    Write-Host ''
}

function Start-Toolbox {
    Initialize-Toolbox
    while ($true) {
        Show-MainMenu
        $choice = Read-Choice 'Choose an option' @('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '0')
        if ($choice -eq '0') { break }
        $pauseAfter = $true
        try {
            switch ($choice) {
                '1' { Invoke-OcrMenu }
                '2' { Invoke-MergeMenu }
                '3' { Invoke-PagesSubMenu; $pauseAfter = $false }
                '4' { Invoke-ConvertSubMenu; $pauseAfter = $false }
                '5' { Invoke-CompressMenu }
                '6' { Invoke-SearchMenu }
                '7' { Invoke-CountMenu }
                '8' { Invoke-UnlockMenu }
                '9' { Invoke-SetupCheck }
                '10' { Invoke-About }
            }
        } catch {
            $message = $_.Exception.Message
            if ($message -eq 'PDFTB_MENU' -or $message -eq 'PDFTB_BACK') { continue }
            Write-Host ''
            Write-Host "Something went wrong: $message" -ForegroundColor Red
        } finally {
            Clear-PdfPassword
        }
        if ($pauseAfter) { Wait-ForEnter }
    }
    Clear-PdfPassword
}

# Run the menu unless the file is being dot-sourced (used for testing).
if ($MyInvocation.InvocationName -ne '.') { Start-Toolbox }
