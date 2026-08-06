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
    * Comments                -> comment parts removed (Word/PowerPoint/Excel)
                                 plus their references, relationships and
                                 content-type overrides.
    * Tracked changes         -> the identifying ``w:author`` / ``w:date`` /
                                 ``w:id`` attributes are stripped from revision
                                 markup so *who* and *when* is erased while the
                                 document structure stays intact.

Only these metadata parts are modified; all other archive members are copied
verbatim.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

SUPPORTED_EXTENSIONS = (".docx", ".pptx", ".xlsx")

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


def _strip_docx_comment_refs(doc_xml: bytes) -> bytes:
    """Remove comment range/reference markers from word/document.xml."""
    text = doc_xml.decode("utf-8")
    # Self-closing comment markers.
    text = re.sub(r"<w:commentRangeStart\b[^>]*/>", "", text)
    text = re.sub(r"<w:commentRangeEnd\b[^>]*/>", "", text)
    # A run that only carries a comment reference -> drop the whole run.
    text = re.sub(
        r"<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:commentReference\b[^>]*/>.*?</w:r>",
        "",
        text,
        flags=re.DOTALL,
    )
    # Any stray comment reference left behind.
    text = re.sub(r"<w:commentReference\b[^>]*/>", "", text)
    return text.encode("utf-8")


def _strip_tracked_change_identity(doc_xml: bytes) -> bytes:
    """Remove author/date/id attributes from tracked-change markup.

    This erases *who* made a change and *when*, without accepting or rejecting
    the change itself, so document content and structure are preserved.
    """
    text = doc_xml.decode("utf-8")
    text = re.sub(r'\s+w:author="[^"]*"', ' w:author=""', text)
    text = re.sub(r'\s+w:date="[^"]*"', "", text)
    # Anonymise initials used by comments/annotations too.
    text = re.sub(r'\s+w:initials="[^"]*"', "", text)
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
    """Strip metadata from a single OOXML document.

    Args:
        input_path: Path to the source .docx/.pptx/.xlsx file.
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

    removed_parts: List[str] = []
    cleaned_parts: List[str] = []

    try:
        with zipfile.ZipFile(input_path, "r") as zin:
            names = zin.namelist()
            infos = {info.filename: info for info in zin.infolist()}
            data: Dict[str, bytes] = {name: zin.read(name) for name in names}

        # Identify comment-related parts to drop (Word/PowerPoint/Excel).
        comment_parts = [
            n for n in list(data)
            if re.search(
                r"(word/comments.*\.xml$|"
                r"word/(commentsExtended|commentsIds|commentsExtensible)\.xml$|"
                r"ppt/comments/.*\.xml$|"
                r"ppt/slides/_rels/.*comment.*|"
                r"xl/comments.*\.xml$|"
                r"xl/threadedComments/.*\.xml$|"
                r".*/threadedComments/.*\.xml$)",
                n,
                re.IGNORECASE,
            )
        ]

        # 1. Drop custom properties + comment parts.
        drop = set()
        if "docProps/custom.xml" in data:
            drop.add("docProps/custom.xml")
        drop.update(comment_parts)

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

        # 5. Update relationship parts (remove rels to dropped parts).
        rel_targets = ["custom.xml"] + [os.path.basename(p) for p in comment_parts]
        for rels_name in [n for n in data if n.endswith(".rels")]:
            data[rels_name] = _remove_relationships(data[rels_name], rel_targets)

        # 6. DOCX: strip comment refs + tracked-change identity from document.xml
        #    and headers/footers.
        for part in [n for n in data if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]:
            new_bytes = _strip_docx_comment_refs(data[part])
            new_bytes = _strip_tracked_change_identity(new_bytes)
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
