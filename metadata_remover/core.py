"""Core metadata-removal logic for Office Open XML documents.

The approach is deliberately surgical: we open the OOXML package (a ZIP
archive) and only rewrite the *metadata* parts, leaving every content part
(``word/document.xml``, ``ppt/slides/*``, ``xl/worksheets/*`` ...) byte-for-byte
untouched.  This guarantees that visible content and formatting are preserved.

What gets removed / cleaned:
    * ``docProps/core.xml``   -> author (creator), last-modified-by, created and
                                 modified timestamps, revision number, title,
                                 subject, keywords, description, category,
                                 content status, last-printed date.
    * ``docProps/app.xml``    -> application name & version, company, manager,
                                 and other identifying application fields.
    * ``docProps/custom.xml`` -> all custom properties (part removed entirely,
                                 together with its relationship and content-type
                                 override).
    * Comments                -> comments are *kept* but anonymised: the author
                                 name, initials, e-mail, user id and timestamp
                                 are blanked in every comment / author / people
                                 part (Word/PowerPoint/Excel, classic & threaded)
                                 while the comment text and anchors stay intact.
    * Tracked changes         -> revisions are *kept* but anonymised: the
                                 identifying ``w:author`` / ``w:date`` /
                                 ``w:initials`` attributes are stripped from the
                                 revision markup so *who* and *when* is erased
                                 while the changes stay visible and the document
                                 structure is preserved.

Only these metadata parts are modified; all other archive members are copied
verbatim.

Beyond the modern OOXML formats this module also supports:

    * Legacy binary Office formats (``.doc``, ``.ppt``, ``.xls``) via a
      surgical edit of the OLE2 compound file with :mod:`olefile`.  The two
      standard document-property streams (``\\x05SummaryInformation`` and
      ``\\x05DocumentSummaryInformation``) are rewritten as empty property sets
      of identical length, while every content stream is left byte-for-byte
      untouched.  No LibreOffice / conversion is involved, so it is instant and
      never triggers a printer query.  Requires the small, pure-Python
      ``olefile`` package.
    * PDF files (``.pdf``) via :mod:`pikepdf`: the document-information
      dictionary and the XMP metadata stream are removed while every page and
      all visible content are preserved.  This requires the optional
      ``pikepdf`` package.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

# Modern Office Open XML formats — handled with zero dependencies.
OOXML_EXTENSIONS = (".docx", ".pptx", ".xlsx")
# Legacy binary Office formats — handled in-place with olefile.
LEGACY_EXTENSIONS = (".doc", ".ppt", ".xls")
# PDF — handled through pikepdf.
PDF_EXTENSIONS = (".pdf",)

SUPPORTED_EXTENSIONS = OOXML_EXTENSIONS + LEGACY_EXTENSIONS + PDF_EXTENSIONS


# ---------------------------------------------------------------------------
# Capability detection (which optional format engines are available)
# ---------------------------------------------------------------------------

def legacy_support_available() -> bool:
    """True if the optional ``olefile`` package is importable.

    ``olefile`` is a tiny, pure-Python library used to strip metadata from
    legacy ``.doc`` / ``.ppt`` / ``.xls`` files in place.
    """
    try:
        import olefile  # noqa: F401
        return True
    except Exception:
        return False


def pdf_support_available() -> bool:
    """True if the optional ``pikepdf`` package is importable (needed for PDF)."""
    try:
        import pikepdf  # noqa: F401
        return True
    except Exception:
        return False


def capabilities() -> Dict[str, bool]:
    """Return a map of which format families can currently be processed."""
    return {
        "ooxml": True,  # always available (pure standard library)
        "legacy": legacy_support_available(),
        "pdf": pdf_support_available(),
    }


# ---------------------------------------------------------------------------
# Clean replacement content for the standard property parts
# ---------------------------------------------------------------------------

# A minimal but valid core-properties part with every field emptied.
_CLEAN_CORE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '</cp:coreProperties>'
).encode("utf-8")


def _clean_app_xml(original: bytes) -> bytes:
    """Return an app.xml with identifying fields emptied.

    We keep the part structurally valid (the same root/namespace) but blank out
    every element that can carry identifying information.  Elements that only
    describe document geometry (e.g. ``<Pages>``, ``<Words>``) are harmless, but
    to be safe and predictable we simply emit a clean, minimal app.xml.
    """
    # Detect the correct namespace flavour (Word/Excel/PowerPoint share it).
    ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Properties xmlns="{ns}" xmlns:vt="{vt}">'
        '<Application></Application>'
        '<Company></Company>'
        '</Properties>'
    ).encode("utf-8")


# ---------------------------------------------------------------------------
# Helpers for editing the small XML "wiring" parts with regex
# ---------------------------------------------------------------------------

def _remove_relationships(rels_xml: bytes, target_substrings: List[str]) -> bytes:
    """Remove <Relationship> elements whose Target matches any substring."""
    text = rels_xml.decode("utf-8")
    pattern = re.compile(r"<Relationship\b[^>]*?/>", re.IGNORECASE)

    def keep(match: re.Match) -> str:
        el = match.group(0)
        m = re.search(r'Target="([^"]*)"', el, re.IGNORECASE)
        if m and any(sub.lower() in m.group(1).lower() for sub in target_substrings):
            return ""
        return el

    return pattern.sub(keep, text).encode("utf-8")


def _remove_content_type_overrides(ct_xml: bytes, part_substrings: List[str]) -> bytes:
    """Remove <Override> entries in [Content_Types].xml for given parts."""
    text = ct_xml.decode("utf-8")
    pattern = re.compile(r"<Override\b[^>]*?/>", re.IGNORECASE)

    def keep(match: re.Match) -> str:
        el = match.group(0)
        m = re.search(r'PartName="([^"]*)"', el, re.IGNORECASE)
        if m and any(sub.lower() in m.group(1).lower() for sub in part_substrings):
            return ""
        return el

    return pattern.sub(keep, text).encode("utf-8")


def _strip_tracked_change_identity(doc_xml: bytes) -> bytes:
    """Remove author/date/id attributes from tracked-change markup.

    This erases *who* made a change and *when*, without accepting or rejecting
    the change itself, so document content and structure are preserved.  The
    revisions themselves stay in the document and remain visible as tracked
    changes — only the person's identity is blanked.
    """
    text = doc_xml.decode("utf-8")
    text = re.sub(r'\s+w:author="[^"]*"', ' w:author=""', text)
    text = re.sub(r'\s+w:date="[^"]*"', "", text)
    # Anonymise initials used by comments/annotations too.
    text = re.sub(r'\s+w:initials="[^"]*"', "", text)
    return text.encode("utf-8")


# Attributes that carry a person's identity inside comment / author / people
# parts.  These are blanked (value emptied) so the comment stays intact and
# anchored while the name/e-mail/user-id disappears.
_AUTHOR_NAME_ATTRS = (
    "w:author", "w15:author", "w16cid:author",
    "author", "name", "displayName", "userId", "email", "userName",
    "w15:userId", "w15:providerId",
)
_AUTHOR_DATE_ATTRS = ("w:date", "date", "dT", "dateTime", "created", "w16du:dateUtc")
_AUTHOR_INITIALS_ATTRS = ("w:initials", "initials")


def _anonymize_comment_authors(xml_bytes: bytes) -> bytes:
    """Blank out identity fields in a comment / author / people XML part.

    The comment text and structure are kept intact — only the author name,
    initials, e-mail, user id and timestamp are removed.  This is applied only
    to comment/author/people parts, so blanking generic attributes such as
    ``name`` or ``displayName`` is safe here.
    """
    text = xml_bytes.decode("utf-8")

    # Blank identity-bearing attributes (keep the attribute, empty its value so
    # structural references such as ids stay valid).
    for attr in _AUTHOR_NAME_ATTRS:
        text = re.sub(
            rf'({re.escape(attr)})="[^"]*"',
            r'\1=""',
            text,
        )
    # Remove initials and date attributes entirely.
    for attr in _AUTHOR_INITIALS_ATTRS + _AUTHOR_DATE_ATTRS:
        text = re.sub(rf'\s+{re.escape(attr)}="[^"]*"', "", text)

    # Excel legacy comments store the name as element text: <author>Jane</author>.
    text = re.sub(r"(<author\b[^>]*>)[^<]*(</author>)", r"\1\2", text)

    return text.encode("utf-8")


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------

@dataclass
class CleanResult:
    """Outcome of cleaning a single document."""

    input_path: str
    output_path: str
    removed_parts: List[str] = field(default_factory=list)
    cleaned_parts: List[str] = field(default_factory=list)
    success: bool = True
    error: Optional[str] = None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _output_path_for(input_path: str, output_dir: Optional[str], suffix: str) -> str:
    base = os.path.basename(input_path)
    name, ext = os.path.splitext(base)
    out_name = f"{name}{suffix}{ext}"
    target_dir = output_dir if output_dir else os.path.dirname(os.path.abspath(input_path))
    os.makedirs(target_dir, exist_ok=True)
    return os.path.join(target_dir, out_name)


def clean_document(
    input_path: str,
    output_dir: Optional[str] = None,
    suffix: str = "_clean",
    overwrite: bool = False,
) -> CleanResult:
    """Strip metadata from a single document.

    Supported formats:
        * ``.docx`` / ``.pptx`` / ``.xlsx`` -> surgical ZIP/XML cleaning
          (no external tools required).
        * ``.doc`` / ``.ppt`` / ``.xls``    -> in-place OLE metadata strip
          (requires the ``olefile`` package; no LibreOffice needed).
        * ``.pdf``                          -> pikepdf metadata strip
          (requires the ``pikepdf`` package).

    Args:
        input_path: Path to the source document.
        output_dir: Directory for the cleaned file. Defaults to the source
            file's own directory.
        suffix: Suffix appended to the file name (default ``_clean``). The
            original file is never modified.
        overwrite: If True, allow overwriting an existing output file.

    Returns:
        A :class:`CleanResult` describing what happened.
    """
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error=f"Unsupported file type: {ext or '(none)'}",
        )

    if not os.path.isfile(input_path):
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error="File not found.",
        )

    # Capability gating for the optional engines.
    if ext in LEGACY_EXTENSIONS and not legacy_support_available():
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error=(
                "Cleaning legacy .doc/.ppt/.xls files requires the 'olefile' "
                "package. Install it with: pip install olefile"
            ),
        )
    if ext in PDF_EXTENSIONS and not pdf_support_available():
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error=(
                "PDF support requires the 'pikepdf' package. Install it with: "
                "pip install pikepdf"
            ),
        )

    output_path = _output_path_for(input_path, output_dir, suffix)
    if os.path.abspath(output_path) == os.path.abspath(input_path):
        # Never write over the original.
        output_path = _output_path_for(input_path, output_dir, suffix + "_1")
    if os.path.exists(output_path) and not overwrite:
        # Find a free name so we never clobber an existing file.
        i = 1
        base_out = output_path
        while os.path.exists(output_path):
            root, e = os.path.splitext(base_out)
            output_path = f"{root}({i}){e}"
            i += 1

    try:
        if ext in OOXML_EXTENSIONS:
            removed_parts, cleaned_parts = _clean_ooxml(input_path, output_path)
        elif ext in LEGACY_EXTENSIONS:
            removed_parts, cleaned_parts = _clean_legacy(input_path, output_path)
        else:  # PDF
            removed_parts, cleaned_parts = _clean_pdf(input_path, output_path)

        return CleanResult(
            input_path=input_path,
            output_path=output_path,
            removed_parts=removed_parts,
            cleaned_parts=cleaned_parts,
            success=True,
        )

    except zipfile.BadZipFile:
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error="Not a valid Office document (corrupt or not an OOXML file).",
        )
    except Exception as exc:  # pragma: no cover - defensive
        return CleanResult(
            input_path=input_path,
            output_path="",
            success=False,
            error=str(exc),
        )


# ---------------------------------------------------------------------------
# Engine 1: modern OOXML (.docx/.pptx/.xlsx) — surgical ZIP/XML cleaning
# ---------------------------------------------------------------------------

def _clean_ooxml(input_path: str, output_path: str) -> Tuple[List[str], List[str]]:
    """Surgically strip metadata from an OOXML package, writing ``output_path``.

    Returns ``(removed_parts, cleaned_parts)``. May raise ``zipfile.BadZipFile``
    for a corrupt / non-OOXML file.
    """
    removed_parts: List[str] = []
    cleaned_parts: List[str] = []

    with zipfile.ZipFile(input_path, "r") as zin:
        names = zin.namelist()
        infos = {info.filename: info for info in zin.infolist()}
        data: Dict[str, bytes] = {name: zin.read(name) for name in names}

    # Identify comment / author / people parts (Word/PowerPoint/Excel).  These
    # are kept in the document but have their author identities blanked, so the
    # comments and tracked changes remain visible while becoming anonymous.
    author_parts = [
        n for n in list(data)
        if re.search(
            r"(word/comments.*\.xml$|"
            r"word/(commentsExtended|commentsIds|commentsExtensible)\.xml$|"
            r"word/people\.xml$|"
            r"ppt/comments/.*\.xml$|"
            r"ppt/(commentAuthors|authors)\.xml$|"
            r"ppt/modernComments/.*\.xml$|"
            r"xl/comments.*\.xml$|"
            r"xl/threadedComments/.*\.xml$|"
            r"xl/persons/.*\.xml$|"
            r".*/threadedComments/.*\.xml$|"
            r".*/persons/.*\.xml$)",
            n,
            re.IGNORECASE,
        )
    ]

    # 1. Drop custom properties only (comments are kept, just anonymised).
    drop = set()
    if "docProps/custom.xml" in data:
        drop.add("docProps/custom.xml")

    for part in sorted(drop):
        data.pop(part, None)
        removed_parts.append(part)

    # 2. Clean core.xml.
    if "docProps/core.xml" in data:
        data["docProps/core.xml"] = _CLEAN_CORE_XML
        cleaned_parts.append("docProps/core.xml")

    # 3. Clean app.xml.
    if "docProps/app.xml" in data:
        data["docProps/app.xml"] = _clean_app_xml(data["docProps/app.xml"])
        cleaned_parts.append("docProps/app.xml")

    # 4. Update [Content_Types].xml (remove overrides for dropped parts).
    if "[Content_Types].xml" in data and drop:
        override_targets = ["/" + p for p in drop]
        data["[Content_Types].xml"] = _remove_content_type_overrides(
            data["[Content_Types].xml"], override_targets
        )

    # 5. Update relationship parts (remove rels to dropped parts only).
    rel_targets = ["custom.xml"]
    for rels_name in [n for n in data if n.endswith(".rels")]:
        data[rels_name] = _remove_relationships(data[rels_name], rel_targets)

    # 6. Anonymise the author identity inside every comment / author / people
    #    part.  The comments themselves are preserved.
    for part in author_parts:
        new_bytes = _anonymize_comment_authors(data[part])
        if new_bytes != data[part]:
            data[part] = new_bytes
            if part not in cleaned_parts:
                cleaned_parts.append(part)

    # 7. DOCX: anonymise tracked-change identity in document.xml and
    #    headers/footers/notes.  Comment anchors are left in place so the
    #    (now anonymous) comments stay attached to the right text.
    for part in [n for n in data if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]:
        new_bytes = _strip_tracked_change_identity(data[part])
        if new_bytes != data[part]:
            data[part] = new_bytes
            if part not in cleaned_parts:
                cleaned_parts.append(part)

    # Write the new package, preserving original compression per member.
    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        # Preserve original member order for maximum compatibility.
        for name in names:
            if name not in data:
                continue  # dropped part
            info = infos.get(name)
            compress = (
                info.compress_type if info is not None else zipfile.ZIP_DEFLATED
            )
            zout.writestr(name, data[name], compress_type=compress)

    shutil.move(tmp_path, output_path)
    return removed_parts, cleaned_parts


# ---------------------------------------------------------------------------
# Engine 2: legacy Office (.doc/.ppt/.xls) — in-place OLE metadata strip
# ---------------------------------------------------------------------------
#
# Legacy 97-2003 files are OLE2 compound files. Their document metadata lives in
# two standard property-set streams at the root of the container:
#     \x05SummaryInformation          title, subject, author, keywords,
#                                      comments, template, last-saved-by,
#                                      revision number, application name,
#                                      create/last-saved timestamps ...
#     \x05DocumentSummaryInformation  category, company, manager, and any
#                                      user-defined custom properties ...
# We rewrite each of those streams in place with a valid but EMPTY property set
# (property count = 0) padded to the original byte length, so all the document's
# real content streams stay byte-for-byte identical. No external program
# (LibreOffice) is used, so there is no printer query and no conversion delay.

_OLE_METADATA_STREAMS = ("\x05SummaryInformation", "\x05DocumentSummaryInformation")


def _empty_property_set(original: bytes) -> bytes:
    """Return a valid OLE property-set blob holding ZERO properties.

    The result is padded with null bytes to exactly ``len(original)`` so it can
    be written back into the stream in place (``olefile`` requires the new data
    to match the existing stream size). The byte-order mark, version, system
    identifier, CLSID and the format ID(s) are copied from the original header
    so the stream stays structurally valid; only the property *values* are
    dropped by setting each section's property count to 0.
    """
    import struct

    size = len(original)
    if size < 48:  # too small to be a real property set — just zero it out
        return b"\x00" * size

    bom, ver, sysid = struct.unpack_from("<HHI", original, 0)
    clsid = original[8:24]
    num_sets = struct.unpack_from("<I", original, 24)[0]
    if num_sets not in (1, 2):
        num_sets = 1
    fmtids = [original[28 + i * 20: 28 + i * 20 + 16] for i in range(num_sets)]

    header = (
        struct.pack("<HHI", bom, ver, sysid) + clsid + struct.pack("<I", num_sets)
    )
    first_section_offset = 28 + num_sets * 20
    set_table = b""
    body = b""
    cursor = first_section_offset
    for i in range(num_sets):
        set_table += fmtids[i] + struct.pack("<I", cursor)
        body += struct.pack("<II", 8, 0)  # section size = 8, property count = 0
        cursor += 8

    blob = header + set_table + body
    if len(blob) > size:  # pragma: no cover - defensive
        blob = blob[:size]
    else:
        blob += b"\x00" * (size - len(blob))
    return blob


def _clean_legacy(input_path: str, output_path: str) -> Tuple[List[str], List[str]]:
    """Strip document-property metadata from a legacy .doc/.ppt/.xls file.

    Works directly on the OLE2 compound file with ``olefile``: the two standard
    metadata streams are rewritten as empty property sets while every content
    stream is preserved byte-for-byte. No LibreOffice / conversion is involved.
    Returns ``(removed_parts, cleaned_parts)``.
    """
    import olefile

    # Work on a copy so the original is never touched and a failure can't leave
    # a half-written file at the destination.
    shutil.copy2(input_path, output_path)

    try:
        ole = olefile.OleFileIO(output_path, write_mode=True)
    except Exception as exc:
        if os.path.exists(output_path):
            os.remove(output_path)
        raise RuntimeError(
            "Not a valid legacy Office file (could not read its structure)."
        ) from exc

    removed: List[str] = []
    try:
        for stream in _OLE_METADATA_STREAMS:
            if ole.exists(stream):
                original = ole.openstream(stream).read()
                ole.write_stream(stream, _empty_property_set(original))
                removed.append(stream.lstrip("\x05"))
    finally:
        ole.close()

    if not removed:
        return ["no metadata streams found (file was already clean)"], []
    return ["document properties: " + ", ".join(removed)], []


# ---------------------------------------------------------------------------
# Engine 3: PDF (.pdf) — pikepdf metadata strip
# ---------------------------------------------------------------------------

def _clean_pdf(input_path: str, output_path: str) -> Tuple[List[str], List[str]]:
    """Strip the document-info dictionary and XMP metadata from a PDF.

    Every page and all visible content are preserved; only metadata is removed.
    Returns ``(removed_parts, cleaned_parts)``.
    """
    import pikepdf  # imported lazily; availability checked by the caller

    removed: List[str] = []
    with pikepdf.open(input_path) as pdf:
        # 1. Document information dictionary (Author, Title, Producer, ...).
        try:
            if pdf.docinfo and len(pdf.docinfo) > 0:
                keys = ", ".join(str(k).lstrip("/") for k in pdf.docinfo.keys())
                removed.append(f"document info ({keys})")
            del pdf.docinfo
        except Exception:
            pass

        # 2. XMP metadata stream (extended/embedded metadata).
        try:
            if "/Metadata" in pdf.Root:
                del pdf.Root.Metadata
                removed.append("XMP metadata stream")
        except Exception:
            pass

        pdf.save(output_path)

    if not removed:
        removed.append("no metadata found (file was already clean)")
    return removed, []


def clean_documents(
    input_paths: List[str],
    output_dir: Optional[str] = None,
    suffix: str = "_clean",
    progress_callback: Optional[Callable[[int, int, CleanResult], None]] = None,
) -> List[CleanResult]:
    """Clean multiple documents, reporting progress via an optional callback.

    The callback receives ``(index, total, result)`` after each file.
    """
    results: List[CleanResult] = []
    total = len(input_paths)
    for i, path in enumerate(input_paths, start=1):
        result = clean_document(path, output_dir=output_dir, suffix=suffix)
        results.append(result)
        if progress_callback:
            progress_callback(i, total, result)
    return results
