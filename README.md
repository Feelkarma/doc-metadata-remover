# Doc Metadata Remover

A lightweight, cross-platform desktop app that **strips metadata from Office
documents and PDFs** — modern `.docx` / `.pptx` / `.xlsx`, legacy
`.doc` / `.ppt` / `.xls`, and `.pdf` — without touching a single byte of their
visible content or formatting.

For modern Office files it works by editing the Office Open XML package directly
(the file *is* a ZIP archive): only the metadata parts are rewritten or removed,
and every content part is copied through verbatim. Legacy binary formats are
cleaned the same surgical way — directly inside the file — using the tiny
pure-Python `olefile` package, and PDFs use `pikepdf`. **No Microsoft Office and
no LibreOffice are required for any format.** See
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
  - **Comment identity** — comments are **kept** (Word, PowerPoint and Excel,
    classic *and* threaded), but the author name, initials, e-mail, user id and
    timestamp attached to each one are blanked, so the feedback stays while the
    person behind it becomes anonymous
  - **Tracked-change identity** — the *who* (`author`) and *when* (`date`) are
    erased from revision markup while the change itself is left intact and
    visible
- 🔒 **Never overwrites your originals** — output is saved as
  `filename_clean.docx` (a free name is chosen automatically if it already
  exists)
- 🎯 **Content-safe** — content parts (`document.xml`, slides, worksheets,
  formulas, images, styles) are copied byte-for-byte, so nothing about how the
  document looks or reads changes
- 🖱️ **Modern tkinter GUI** — drag-and-drop, file picker, batch processing,
  progress bar, live log, and optional output-folder selection
- 🗂️ **Legacy formats too** — `.doc`, `.ppt` and `.xls` are cleaned in place
  with `olefile` (the standard document-property streams are wiped) and saved
  back in their original format — instantly, with no external programs
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
| Word         | `.docx` | Built-in (ZIP/XML) | None | core/app/custom props; comments & tracked changes kept but anonymised |
| PowerPoint   | `.pptx` | Built-in (ZIP/XML) | None | core/app/custom props; comments kept but anonymised |
| Excel        | `.xlsx` | Built-in (ZIP/XML) | None | core/app/custom props; comments/threaded comments kept but anonymised |
| Word (legacy)       | `.doc` | In-place OLE strip (`olefile`) | `pip install olefile` | wipes document-property streams; in-body comment/revision names not touched |
| PowerPoint (legacy) | `.ppt` | In-place OLE strip (`olefile`) | `pip install olefile` | wipes document-property streams; in-body comment/revision names not touched |
| Excel (legacy)      | `.xls` | In-place OLE strip (`olefile`) | `pip install olefile` | wipes document-property streams; in-body comment/revision names not touched |
| PDF          | `.pdf` | `pikepdf` | `pip install pikepdf` | strips document info + XMP metadata; pages/content preserved |

**How the legacy path works now.** Old `.doc` / `.ppt` / `.xls` files aren't
ZIP/XML packages — they are OLE2 "compound" files. Their metadata lives in two
standard streams (`\x05SummaryInformation` and `\x05DocumentSummaryInformation`),
which the app rewrites as empty in place, leaving every content stream
byte-for-byte identical. This uses the tiny pure-Python **`olefile`** package —
**no LibreOffice, no conversion, and no printer prompts** — so it's instant. If
`olefile` isn't installed, everything else still works and legacy files are
reported as unavailable.

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
- `olefile` — cleaning legacy `.doc` / `.ppt` / `.xls` files (tiny, pure-Python)
- `pikepdf` — PDF metadata removal

All optional extras are ordinary `pip` packages — **no separate programs such
as LibreOffice are needed**. Whatever you don't install simply shows up as
"unavailable" in the welcome dialog; everything else keeps working.

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
4. Deletes `custom.xml` — and removes its relationship entry (`*.rels`) and
   content-type override (`[Content_Types].xml`) so the package stays valid.
5. **Anonymises comments (they are kept, not deleted):** in every comment /
   author / people part (Word, PowerPoint and Excel, classic *and* threaded) it
   blanks the author name, initials, e-mail, user id and timestamp while leaving
   the comment text and its anchor in place.
6. **Anonymises tracked changes:** strips `author` / `date` / `initials`
   attributes from revision markup in the document body so the change stays
   visible but the identity is gone.
7. Re-zips, preserving the original member order and per-part compression.

Because only metadata/identity fields are modified, the document's visible
content and formatting are guaranteed to be unchanged.

### Legacy `.doc` / `.ppt` / `.xls`

These older files are OLE2 *compound files* (a little filesystem-in-a-file), not
ZIP/XML packages. Their document-property metadata lives in two standard
streams:

```
\x05SummaryInformation          ← title, subject, author, keywords, comments,
                                  last-saved-by, revision number, app name, dates
\x05DocumentSummaryInformation  ← category, company, manager, custom properties
```

Using the pure-Python [`olefile`](https://pypi.org/project/olefile/) package,
the app rewrites each of those two streams **in place** with a valid but empty
property set (padded to the exact original length, so the file layout is
unchanged). Every other stream — the actual document content and formatting —
is left byte-for-byte identical.

Because it edits the file directly, there is **no conversion, no LibreOffice,
and no printer prompt** — cleaning is instant and the file keeps its original
extension.

> **Scope note.** For these legacy binary formats the app clears the standard
> document-property streams (author, company, dates, last-saved-by, etc.).
> Unlike the modern formats, it does **not** reach into the file body to
> anonymise names attached to inline comments or tracked changes — the old
> binary layout makes that risky to edit in place. If you need that, open the
> file in Word and save it as `.docx` first, then run it through the app.

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
