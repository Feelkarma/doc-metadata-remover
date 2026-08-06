#!/usr/bin/env python3
"""Entry point for the Doc Metadata Remover desktop application.

Run with no arguments to launch the GUI:

    python app.py

Or pass files on the command line for headless/batch use:

    python app.py file1.docx file2.xlsx --output ./cleaned
"""

import sys


def main() -> int:
    # If files were passed, run in CLI mode; otherwise launch the GUI.
    args = sys.argv[1:]
    if args:
        from metadata_remover.cli import main as cli_main

        return cli_main(args)

    from metadata_remover.gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
