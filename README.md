# Doc Metadata Remover

A lightweight, cross-platform desktop app that **strips metadata from Office
documents** — `.docx`, `.pptx`, and `.xlsx` — without touching a single byte of
their visible content or formatting.

It works by editing the Office Open XML package directly (the file *is* a ZIP
archive): only the metadata parts are rewritten or removed, and every content
part is copied through verbatim. **No Microsoft Office and no LibreOffice
required.**

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
- ⌨️ **Command-line mode** for scripting and automation
- 📦 **Packageable as a standalone EXE** with PyInstaller
- 🐍 **Zero required dependencies** — the core and GUI use only the Python
  standard library

---

## Supported formats

| Format | Extension | Notes                                   |
| ------ | --------- | --------------------------------------- |
| Word         | `.docx` | core/app/custom props, comments, tracked-change identity |
| PowerPoint   | `.pptx` | core/app/custom props, comments         |
| Excel        | `.xlsx` | core/app/custom props, comments/threaded comments |

> Legacy binary formats (`.doc`, `.ppt`, `.xls`) are **not** supported — they
> are not ZIP/XML based. Save them as the modern format first.

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

The app runs with **no third-party packages** — `pip install` is only needed
for the optional drag-and-drop feature (`tkinterdnd2`) and for building an EXE
(`pyinstaller`).

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
