"""doc-metadata-remover

Strip metadata from documents without altering visible content or formatting:

    * Modern Office (DOCX, PPTX, XLSX) via direct ZIP/XML manipulation.
    * Legacy Office (DOC, PPT, XLS) via an in-place OLE metadata strip (olefile).
    * PDF via pikepdf.
"""

from .core import (
    CleanResult,
    LEGACY_EXTENSIONS,
    OOXML_EXTENSIONS,
    PDF_EXTENSIONS,
    SUPPORTED_EXTENSIONS,
    capabilities,
    clean_document,
    clean_documents,
    legacy_support_available,
    pdf_support_available,
)

__version__ = "1.3.0"
__all__ = [
    "clean_document",
    "clean_documents",
    "CleanResult",
    "SUPPORTED_EXTENSIONS",
    "OOXML_EXTENSIONS",
    "LEGACY_EXTENSIONS",
    "PDF_EXTENSIONS",
    "capabilities",
    "legacy_support_available",
    "pdf_support_available",
]
