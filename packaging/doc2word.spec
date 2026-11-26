# PyInstaller spec: builds dist\Doc2Word\ containing
#   Doc2Word.exe      - the desktop app (no console window)
#   doc2word-cli.exe  - command-line converter (for IT / batch jobs)
# Run from the repo root:  pyinstaller packaging\doc2word.spec --noconfirm

from pathlib import Path

ROOT = Path(SPECPATH).parent
ICON = str(ROOT / "packaging" / "doc2word.ico")

hidden = [
    "win32com", "win32com.client", "pythoncom", "pywintypes", "win32timezone",
    "pillow_heif", "pdf2docx", "doc2word.engines.pdf2docx_engine",
]
# Big Qt modules we never use - keeps the install much smaller.
excludes = [
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.Qt3DCore",
    "PySide6.QtQuick", "PySide6.QtQml", "PySide6.QtMultimedia", "PySide6.QtCharts",
    "PySide6.QtDataVisualization", "PySide6.QtPdf", "tkinter", "pytest",
]


def analysis(script):
    return Analysis(
        [str(ROOT / "packaging" / script)],
        pathex=[str(ROOT)],
        hiddenimports=hidden,
        excludes=excludes,
        noarchive=False,
    )


gui_a = analysis("launch_gui.py")
cli_a = analysis("launch_cli.py")

gui_exe = EXE(
    PYZ(gui_a.pure), gui_a.scripts, [],
    exclude_binaries=True, name="Doc2Word", icon=ICON,
    console=False, upx=False,
    version=str(ROOT / "packaging" / "version_info.txt"),
)
cli_exe = EXE(
    PYZ(cli_a.pure), cli_a.scripts, [],
    exclude_binaries=True, name="doc2word-cli", icon=ICON,
    console=True, upx=False,
)

coll = COLLECT(
    gui_exe, gui_a.binaries, gui_a.datas,
    cli_exe, cli_a.binaries, cli_a.datas,
    name="Doc2Word", upx=False,
)
