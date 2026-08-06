"""doc-metadata-remover

Strip metadata from Office Open XML documents (DOCX, PPTX, XLSX) via direct
ZIP/XML manipulation, without altering the visible content or formatting.
"""

from .core import clean_document, CleanResult, SUPPORTED_EXTENSIONS

__version__ = "1.0.0"
__all__ = ["clean_document", "CleanResult", "SUPPORTED_EXTENSIONS"]
