"""Start the desktop app. Any file paths on the command line are added and
converted straight away - that's what the Explorer "Send to" shortcut uses."""

from __future__ import annotations

import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv

    from . import APP_NAME
    from .logsetup import setup_logging

    setup_logging()

    if sys.platform == "win32":
        # Own taskbar icon/grouping instead of "python.exe".
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Doc2Word.App")
        except Exception:  # noqa: BLE001
            pass

    from PySide6.QtWidgets import QApplication

    from .gui.main_window import MainWindow

    app = QApplication(sys.argv[:1])
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_NAME)
    if sys.platform != "win32":
        app.setStyle("Fusion")  # native "windows11" style is used on Windows

    window = MainWindow()
    window.show()

    files = [Path(a) for a in argv if not a.startswith("-")]
    if files and window.add_paths(files):
        window.start()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
