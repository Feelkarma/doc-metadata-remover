# Doc Metadata Remover

A lightweight, cross-platform desktop app that **strips metadata from Office
documents and PDFs** — modern `.docx` / `.pptx` / `.xlsx`, legacy
`.doc` / `.ppt` / `.xls`, and `.pdf` — without touching a single byte of their
visible content or formatting.

For modern Office files it works by editing the Office Open XML package directly
(the file *is* a ZIP archive): only the metadata parts are rewritten or removed,
and every content part is copied through verbatim — **no Microsoft Office and no
LibreOffice required**. Legacy binary formats and PDFs are handled by two
optional helpers (LibreOffice and `pikepdf`); see
[Supported formats](#supported-formats).

When you start the app, a short **welcome dialog** explains in plain language
what the tool does and which formats are currently available on your computer.

---

## Features

- 🧹 **Removes all common metadata**
  - Author / creator and last-modified-by
  - Created, modified and last-printed timestamps
  - Revision number
  - Title, subject, keywords, description, category, content status
  - Application name & version, company, manager (`app.xml`)
  - **Custom properties** (`custom.xml` removed entirely)
  - **Comments** (Word, PowerPoint and Excel comment parts + references)
  - **Tracked-change identity** — the *who* (`author`) and *when* (`date`) are
    erased from revision markup while the change itself is left intact
- 🔒 **Never overwrites your originals** — output is saved as
  `filename_clean.docx` (a free name is chosen automatically if it already
  exists)
- 🎯 **Content-safe** — content parts (`document.xml`, slides, worksheets,
  formulas, images, styles) are copied byte-for-byte, so nothing about how the
  document looks or reads changes
- 🖱️ **Modern tkinter GUI** — drag-and-drop, file picker, batch processing,
  progress bar, live log, and optional output-folder selection
- 🗂️ **Legacy formats too** — `.doc`, `.ppt` and `.xls` are cleaned by
  round-tripping through LibreOffice (if it is installed) and are saved back in
  their original format
- 📄 **PDF support** — document info and XMP metadata are stripped with
  `pikepdf` (if it is installed) while pages and content are preserved
- 👋 **Plain-language welcome dialog** on startup that explains how the app
  works and shows which formats are available (dismissable with a
  "Don't show this again" option)
- ⌨️ **Command-line mode** for scripting and automation
- 📦 **Packageable as a standalone EXE** with PyInstaller
- 🐍 **Zero required dependencies for modern Office files** — that path uses only
  the Python standard library; legacy formats and PDFs use optional helpers

---

## Supported formats

| Format | Extension | Handler | Requirement | Notes |
| ------ | --------- | ------- | ----------- | ----- |
| Word         | `.docx` | Built-in (ZIP/XML) | None | core/app/custom props, comments, tracked-change identity |
| PowerPoint   | `.pptx` | Built-in (ZIP/XML) | None | core/app/custom props, comments |
| Excel        | `.xlsx` | Built-in (ZIP/XML) | None | core/app/custom props, comments/threaded comments |
| Word (legacy)       | `.doc` | LibreOffice round-trip | LibreOffice installed | converted to `.docx`, cleaned, converted back |
| PowerPoint (legacy) | `.ppt` | LibreOffice round-trip | LibreOffice installed | converted to `.pptx`, cleaned, converted back |
| Excel (legacy)      | `.xls` | LibreOffice round-trip | LibreOffice installed | converted to `.xlsx`, cleaned, converted back |
| PDF          | `.pdf` | `pikepdf` | `pip install pikepdf` | strips document info + XMP metadata; pages/content preserved |

**Why the legacy path is different.** Old `.doc` / `.ppt` / `.xls` files are not
ZIP/XML packages, so they can't be edited in place. The app hands them to
LibreOffice, converts them to the modern equivalent, cleans that with the
built-in surgical cleaner, then converts the clean result back to the original
format. This requires **LibreOffice** to be installed (a free, separate
download). If LibreOffice isn't found, modern formats and PDFs still work — only
legacy files are reported as unavailable.

**PDFs** need the `pikepdf` package (`pip install pikepdf`). If it isn't
installed, everything else still works and PDFs are reported as unavailable.

---

## Installation

You need **Python 3.8+**.

```bash
git clone https://github.com/Feelkarma/doc-metadata-remover.git
cd doc-metadata-remover

# (Optional) create a virtual environment
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# Optional extras: drag-and-drop support + the EXE builder
pip install -r requirements.txt
```

Cleaning **modern** Office files needs **no third-party packages**. The optional
extras enable more:

- `tkinterdnd2` — drag-and-drop onto the GUI window
- `pyinstaller` — building a standalone EXE
- `pikepdf` — PDF metadata removal
- **LibreOffice** (a separate program, not a pip package) — cleaning legacy
  `.doc` / `.ppt` / `.xls` files. Download it from
  [libreoffice.org](https://www.libreoffice.org/).

Whatever you don't install simply shows up as "unavailable" in the welcome
dialog; everything else keeps working.

---

## Usage

### GUI

```bash
python app.py
```

1. Drag files onto the window, or click **Add files**.
2. (Optional) pick an **Output folder** — otherwise cleaned files are written
   next to the originals.
3. Click **Remove metadata**.

Each cleaned file is saved as `originalname_clean.ext`.

### Command line

```bash
# Clean specific files
python app.py report.docx deck.pptx budget.xlsx

# Send output to a folder
python app.py report.docx --output ./cleaned

# Clean every supported file in a folder (recursively)
python app.py ./my_documents --output ./cleaned

# Custom suffix
python app.py report.docx --suffix _scrubbed
```

### As a library

```python
from metadata_remover import clean_document

result = clean_document("report.docx", output_dir="cleaned")
print(result.success, result.output_path)
print("removed:", result.removed_parts)
print("cleaned:", result.cleaned_parts)
```

---

## Building a standalone executable

A [PyInstaller](https://pyinstaller.org/) spec is included.

```bash
pip install pyinstaller tkinterdnd2
pyinstaller build_exe.spec
```

The binary appears in `dist/`:

- **Windows** → `dist/DocMetadataRemover.exe`
- **macOS**   → `dist/DocMetadataRemover`
- **Linux**   → `dist/DocMetadataRemover`

> Build on the OS you want to target — PyInstaller does not cross-compile. Build
> the `.exe` on Windows, the macOS binary on macOS, etc.

One-liner without the spec file:

```bash
pyinstaller --onefile --windowed --name DocMetadataRemover app.py
```

---

## How it works

An Office Open XML file is a ZIP archive of XML parts. Metadata lives in a small
set of well-known parts:

```
docProps/core.xml     ← author, dates, revision, title, subject, keywords …
docProps/app.xml      ← application name/version, company, manager …
docProps/custom.xml   ← user-defined custom properties
word/comments*.xml    ← Word comments (and PowerPoint/Excel equivalents)
```

Doc Metadata Remover:

1. Opens the archive in memory.
2. Replaces `core.xml` with an empty, valid properties part.
3. Rewrites `app.xml` with identifying fields blanked out.
4. Deletes `custom.xml` and all comment parts — and also removes their
   relationship entries (`*.rels`) and content-type overrides
   (`[Content_Types].xml`) so the package stays valid.
5. Strips `author` / `date` / `initials` attributes from tracked-change and
   comment markup in the document body.
6. Re-zips, preserving the original member order and per-part compression.

Because only metadata parts are modified, the document's content and formatting
are guaranteed to be unchanged.

### Legacy `.doc` / `.ppt` / `.xls`

These older files are binary, not ZIP/XML, so they can't be edited in place.
The app:

1. Asks LibreOffice to convert the file to its modern equivalent
   (`.doc → .docx`, `.ppt → .pptx`, `.xls → .xlsx`).
2. Runs the same surgical XML cleaner described above on that copy.
3. Asks LibreOffice to convert the cleaned copy back to the original format.

The cleaned file keeps its original extension. (LibreOffice's *same-format*
conversion preserves metadata, which is exactly why the round-trip through the
modern format is required.)

### PDF

PDF metadata lives in two places: the document-info dictionary and an XMP
metadata stream. Using `pikepdf`, the app deletes both and re-saves the file.
Pages, text and layout are left untouched.

---

## Testing

```bash
pip install pytest python-docx python-pptx openpyxl
pytest -q
```

The test suite builds documents with known metadata and verifies that the
metadata is removed while the content survives.

---

## License

[MIT](LICENSE) © 2026 Feelkarma
