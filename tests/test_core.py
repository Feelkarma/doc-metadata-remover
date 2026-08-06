"""Tests for the metadata-removal core.

These tests build small Office documents with known metadata (using
python-docx / python-pptx / openpyxl, which are dev-only dependencies), run the
cleaner, and assert that (a) metadata is gone and (b) content is preserved.

Run with:  pytest -q
"""

import datetime
import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from metadata_remover import (  # noqa: E402
    clean_document,
    libreoffice_available,
    pdf_support_available,
)

docx = pytest.importorskip("docx")
pptx = pytest.importorskip("pptx")
openpyxl = pytest.importorskip("openpyxl")


def _read(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode("utf-8", "ignore") if part in z.namelist() else None


def test_docx_metadata_removed_and_content_preserved(tmp_path):
    from docx import Document

    src = tmp_path / "in.docx"
    d = Document()
    d.add_paragraph("Body text that must survive.")
    cp = d.core_properties
    cp.author = "Secret Author"
    cp.last_modified_by = "Secret Editor"
    cp.title = "Secret Title"
    cp.revision = 9
    cp.created = datetime.datetime(2020, 1, 1)
    d.save(str(src))

    result = clean_document(str(src), output_dir=str(tmp_path))
    assert result.success, result.error
    assert os.path.exists(result.output_path)
    assert os.path.exists(str(src))  # original preserved

    core = _read(result.output_path, "docProps/core.xml")
    for leaked in ("Secret Author", "Secret Editor", "Secret Title", "2020"):
        assert leaked not in core

    out = Document(result.output_path)
    assert [p.text for p in out.paragraphs] == ["Body text that must survive."]
    assert out.core_properties.author in ("", None)
    assert out.core_properties.last_modified_by in ("", None)


def test_pptx_metadata_removed_and_content_preserved(tmp_path):
    from pptx import Presentation

    src = tmp_path / "in.pptx"
    p = Presentation()
    slide = p.slides.add_slide(p.slide_layouts[0])
    slide.shapes.title.text = "Deck Title"
    p.core_properties.author = "Deck Author"
    p.save(str(src))

    result = clean_document(str(src), output_dir=str(tmp_path))
    assert result.success, result.error

    core = _read(result.output_path, "docProps/core.xml")
    assert "Deck Author" not in core

    out = Presentation(result.output_path)
    titles = [
        sh.text_frame.text
        for s in out.slides
        for sh in s.shapes
        if sh.has_text_frame
    ]
    assert "Deck Title" in titles


def test_xlsx_metadata_removed_and_content_preserved(tmp_path):
    from openpyxl import Workbook, load_workbook

    src = tmp_path / "in.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Header"
    ws["A2"] = 42
    ws["B2"] = "=A2*2"
    wb.properties.creator = "XL Creator"
    wb.properties.title = "XL Secret"
    wb.save(str(src))

    result = clean_document(str(src), output_dir=str(tmp_path))
    assert result.success, result.error

    core = _read(result.output_path, "docProps/core.xml")
    assert "XL Creator" not in core and "XL Secret" not in core

    out = load_workbook(result.output_path)
    ws2 = out.active
    assert ws2["A1"].value == "Header"
    assert ws2["A2"].value == 42
    assert ws2["B2"].value == "=A2*2"


def test_unsupported_extension(tmp_path):
    src = tmp_path / "note.txt"
    src.write_text("hello")
    result = clean_document(str(src), output_dir=str(tmp_path))
    assert not result.success
    assert "Unsupported" in (result.error or "")


@pytest.mark.skipif(not pdf_support_available(), reason="pikepdf not installed")
def test_pdf_metadata_removed_and_content_preserved(tmp_path):
    import pikepdf

    src = tmp_path / "in.pdf"
    pdf = pikepdf.new()
    pdf.add_blank_page(page_size=(200, 200))
    with pdf.open_metadata() as meta:
        meta["dc:creator"] = ["Secret PDF Author"]
        meta["dc:title"] = "Secret PDF Title"
    pdf.docinfo["/Author"] = "Secret PDF Author"
    pdf.docinfo["/Title"] = "Secret PDF Title"
    pdf.save(str(src))
    pdf.close()

    result = clean_document(str(src), output_dir=str(tmp_path))
    assert result.success, result.error
    assert os.path.exists(str(src))  # original preserved

    out = pikepdf.open(result.output_path)
    try:
        assert len(out.pages) == 1  # content preserved
        assert len(out.docinfo) == 0  # doc info stripped
        assert "/Metadata" not in out.Root  # XMP stream removed
    finally:
        out.close()


@pytest.mark.skipif(
    not libreoffice_available(), reason="LibreOffice not installed"
)
def test_legacy_doc_metadata_removed_and_content_preserved(tmp_path):
    from docx import Document

    from metadata_remover.core import find_soffice, _run_soffice_convert

    soffice = find_soffice()
    assert soffice, "LibreOffice reported available but soffice not found"

    # Build a modern .docx with metadata, then convert it to legacy .doc.
    modern = tmp_path / "seed.docx"
    d = Document()
    d.add_paragraph("Legacy body text.")
    d.core_properties.author = "Legacy Secret Author"
    d.save(str(modern))

    _run_soffice_convert(soffice, str(modern), str(tmp_path), "doc:MS Word 97")
    legacy = tmp_path / "seed.doc"
    assert legacy.exists(), "LibreOffice failed to produce the .doc"

    result = clean_document(str(legacy), output_dir=str(tmp_path))
    assert result.success, result.error
    assert result.output_path.lower().endswith(".doc")
    assert os.path.exists(str(legacy))  # original preserved

    # Convert the cleaned .doc back to text and confirm secrets are gone but
    # the body survives.
    _run_soffice_convert(soffice, result.output_path, str(tmp_path), "txt:Text")
    txt = os.path.splitext(result.output_path)[0] + ".txt"
    text = open(txt, encoding="utf-8", errors="ignore").read()
    assert "Legacy body text." in text
    assert "Legacy Secret Author" not in text


def test_original_never_overwritten(tmp_path):
    from docx import Document

    src = tmp_path / "keep.docx"
    Document().save(str(src))
    result = clean_document(str(src), output_dir=str(tmp_path))
    assert result.success
    assert os.path.abspath(result.output_path) != os.path.abspath(str(src))
