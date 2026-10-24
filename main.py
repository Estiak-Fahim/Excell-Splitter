"""Excel Splitter -- application entry point.

Run from source with:
    python main.py

(See README.md for creating a standalone ExcelSplitter.exe with PyInstaller.)
"""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from app.logging_config import configure_logging
from app.main_window import MainWindow


def main() -> int:
    configure_logging()

    app = QApplication(sys.argv)
    app.setApplicationName("Excel Splitter")
    app.setOrganizationName("ExcelSplitter")

    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
