"""Command-line interface for doc-metadata-remover.

Usage:
    python -m metadata_remover.cli file1.docx file2.xlsx --output ./cleaned
"""

from __future__ import annotations

import argparse
import os
from typing import List, Optional

from .core import SUPPORTED_EXTENSIONS, clean_document


def _expand(paths: List[str]) -> List[str]:
    """Expand directories into their supported files."""
    out: List[str] = []
    for p in paths:
        if os.path.isdir(p):
            for root, _dirs, files in os.walk(p):
                for f in files:
                    if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS:
                        out.append(os.path.join(root, f))
        else:
            out.append(p)
    return out


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="doc-metadata-remover",
        description="Strip metadata from DOCX, PPTX and XLSX files without "
        "altering their content.",
    )
    parser.add_argument("files", nargs="+", help="Files or folders to clean.")
    parser.add_argument(
        "-o", "--output", default=None,
        help="Output folder (default: alongside each source file).",
    )
    parser.add_argument(
        "-s", "--suffix", default="_clean",
        help="Suffix for cleaned files (default: _clean).",
    )
    ns = parser.parse_args(argv)

    files = _expand(ns.files)
    if not files:
        print("No files to process.")
        return 1

    ok = 0
    for path in files:
        result = clean_document(path, output_dir=ns.output, suffix=ns.suffix)
        if result.success:
            ok += 1
            print(f"[OK]   {path} -> {result.output_path}")
            if result.removed_parts:
                print(f"       removed: {', '.join(result.removed_parts)}")
        else:
            print(f"[FAIL] {path} -- {result.error}")

    print(f"\nDone: {ok}/{len(files)} file(s) cleaned.")
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
