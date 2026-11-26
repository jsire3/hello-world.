"""The main Doc2Word window."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import QObject, QSettings, Qt, QUrl, Signal
from PySide6.QtGui import QAction, QColor, QIcon, QDesktopServices, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QStyle,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .. import APP_NAME, __version__
from ..converter import describe
from ..formats import SUPPORTED_EXTS, expand_inputs, is_supported
from ..worker import ConversionQueue, Job, SessionFactory, State, default_session_factory

COL_FILE, COL_TYPE, COL_STATUS, COL_NOTES = range(4)
READY = "Ready"

STATE_COLORS = {
    State.DONE: "#1a7f37",
    State.WARNING: "#9a6700",
    State.FAILED: "#cf222e",
    State.CANCELLED: "#6e7781",
    State.WORKING: "#0969da",
}

FILE_FILTER = (
    "Supported files ("
    + " ".join(f"*{e}" for e in sorted(SUPPORTED_EXTS))
    + ");;PDF (*.pdf);;Images (*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.gif *.webp *.heic)"
    ";;Office documents (*.doc *.docx *.rtf *.odt *.txt *.htm *.html *.xls *.xlsx *.ppt *.pptx)"
    ";;All files (*)"
)


class _Bridge(QObject):
    """Carries job updates from the worker thread to the GUI thread."""

    updated = Signal(int, str, str, str, str)  # id, state, message, output, detail


class DropZone(QFrame):
    """Big dashed area that accepts dropped files/folders and opens a picker on click."""

    dropped = Signal(list)
    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("DropZone")
        self.setMinimumHeight(130)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        lay = QVBoxLayout(self)
        title = QLabel("Drop files or folders here")
        title.setObjectName("DropTitle")
        sub = QLabel("PDF, Word, Excel, PowerPoint, RTF, text, web pages and images - or click to browse")
        sub.setObjectName("DropSub")
        for w in (title, sub):
            w.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lay.addWidget(w)
        self._set_hover(False)

    def _set_hover(self, on: bool) -> None:
        self.setProperty("hover", on)
        self.style().unpolish(self)
        self.style().polish(self)

    def mouseReleaseEvent(self, event):  # noqa: N802 (Qt naming)
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

    def dragEnterEvent(self, event: QDragEnterEvent):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_hover(True)

    def dragLeaveEvent(self, event):  # noqa: N802
        self._set_hover(False)

    def dropEvent(self, event: QDropEvent):  # noqa: N802
        self._set_hover(False)
        paths = [Path(u.toLocalFile()) for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.dropped.emit(paths)
            event.acceptProposedAction()


STYLE = """
#DropZone { border: 2px dashed #8c959f; border-radius: 10px; background: palette(base); }
#DropZone[hover="true"] { border-color: #0969da; background: palette(alternate-base); }
#DropTitle { font-size: 16pt; font-weight: 600; }
#DropSub { color: palette(placeholder-text); }
#Banner { background: #fff8c5; color: #4d2d00; border: 1px solid #d4a72c;
          border-radius: 6px; padding: 8px; }
QPushButton#Primary { font-weight: 600; padding: 6px 18px; }
"""


class MainWindow(QMainWindow):
    def __init__(self, session_factory: SessionFactory = default_session_factory):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} - Convert to Word")
        self.resize(920, 600)
        self.setAcceptDrops(True)
        self.settings = QSettings(APP_NAME, APP_NAME)

        self._items: dict[int, QTreeWidgetItem] = {}   # job id -> row
        self._pending: list[QTreeWidgetItem] = []        # added, not yet submitted
        self._batch_outputs: list[str] = []
        self._batch_size = 0

        self._bridge = _Bridge()
        self._bridge.updated.connect(self._on_job_updated)
        self.queue = ConversionQueue(self._from_worker, session_factory=session_factory)

        self._build_ui()
        self._load_settings()
        self._refresh_buttons()

    # -- UI construction ----------------------------------------------------
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        self.setStyleSheet(STYLE)
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 12)
        root.setSpacing(10)

        self.banner = QLabel()
        self.banner.setObjectName("Banner")
        self.banner.setWordWrap(True)
        self.banner.hide()
        root.addWidget(self.banner)
        self._check_office()

        self.drop = DropZone()
        self.drop.dropped.connect(self.add_paths)
        self.drop.clicked.connect(self._browse_files)
        root.addWidget(self.drop)

        add_row = QHBoxLayout()
        b_files = QPushButton("Add files…")
        b_files.clicked.connect(self._browse_files)
        b_folder = QPushButton("Add folder…")
        b_folder.clicked.connect(self._browse_folder)
        add_row.addWidget(b_files)
        add_row.addWidget(b_folder)
        add_row.addStretch()
        self.b_clear = QPushButton("Clear finished")
        self.b_clear.clicked.connect(self._clear_finished)
        add_row.addWidget(self.b_clear)
        root.addLayout(add_row)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["File", "Type", "Status", "Notes"])
        self.tree.setRootIsDecorated(False)
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._context_menu)
        self.tree.itemDoubleClicked.connect(lambda item, _c: self._open_output(item))
        hdr = self.tree.header()
        hdr.setSectionResizeMode(COL_FILE, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(COL_NOTES, QHeaderView.ResizeMode.Stretch)
        self.tree.setColumnWidth(COL_FILE, 300)
        self.tree.setColumnWidth(COL_TYPE, 150)
        self.tree.setColumnWidth(COL_STATUS, 120)
        root.addWidget(self.tree, 1)

        opts = QHBoxLayout()
        opts.addWidget(QLabel("Save to:"))
        self.save_mode = QComboBox()
        self.save_mode.addItems(["Same folder as the original", "A folder I choose…"])
        self.save_mode.activated.connect(self._on_save_mode)
        opts.addWidget(self.save_mode)
        self.folder_label = QLabel()
        self.folder_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        opts.addWidget(self.folder_label, 1)
        self.open_when_done = QCheckBox("Open when finished")
        opts.addWidget(self.open_when_done)
        root.addLayout(opts)

        actions = QHBoxLayout()
        self.summary = QLabel()
        actions.addWidget(self.summary, 1)
        about = QPushButton("About")
        about.setFlat(True)
        about.clicked.connect(self._about)
        actions.addWidget(about)
        self.b_cancel = QPushButton("Stop")
        self.b_cancel.clicked.connect(self._cancel)
        actions.addWidget(self.b_cancel)
        self.b_convert = QPushButton("Convert to Word")
        self.b_convert.setObjectName("Primary")
        self.b_convert.setDefault(True)
        self.b_convert.clicked.connect(self.start)
        actions.addWidget(self.b_convert)
        root.addLayout(actions)

    def _check_office(self) -> None:
        if sys.platform != "win32":
            return
        from ..engines.office import office_available

        if not office_available():
            self.banner.setText(
                "Microsoft Word wasn't found on this PC. Images can still be converted, "
                "but PDFs and Office files need Word installed."
            )
            self.banner.show()

    # -- settings -----------------------------------------------------------
    def _load_settings(self) -> None:
        self._out_dir = self.settings.value("out_dir", "", str)
        mode = self.settings.value("save_mode", 0, int)
        if mode == 1 and not self._out_dir:
            mode = 0
        self.save_mode.setCurrentIndex(mode)
        self.open_when_done.setChecked(self.settings.value("open_when_done", True, bool))
        self._update_folder_label()

    def _save_settings(self) -> None:
        self.settings.setValue("out_dir", self._out_dir)
        self.settings.setValue("save_mode", self.save_mode.currentIndex())
        self.settings.setValue("open_when_done", self.open_when_done.isChecked())

    def _on_save_mode(self, index: int) -> None:
        if index == 1:
            folder = QFileDialog.getExistingDirectory(self, "Save Word files to", self._out_dir or str(Path.home()))
            if folder:
                self._out_dir = folder
            elif not self._out_dir:
                self.save_mode.setCurrentIndex(0)
        self._update_folder_label()
        self._save_settings()

    def _update_folder_label(self) -> None:
        if self.save_mode.currentIndex() == 1 and self._out_dir:
            self.folder_label.setText(self._out_dir)
        else:
            self.folder_label.setText("")

    def _current_out_dir(self) -> Path | None:
        return Path(self._out_dir) if self.save_mode.currentIndex() == 1 and self._out_dir else None

    # -- adding files ---------------------------------------------------------
    def _browse_files(self) -> None:
        start = self.settings.value("last_dir", str(Path.home()), str)
        files, _ = QFileDialog.getOpenFileNames(self, "Choose files to convert", start, FILE_FILTER)
        if files:
            self.settings.setValue("last_dir", str(Path(files[0]).parent))
            self.add_paths([Path(f) for f in files])

    def _browse_folder(self) -> None:
        start = self.settings.value("last_dir", str(Path.home()), str)
        folder = QFileDialog.getExistingDirectory(self, "Choose a folder to convert", start)
        if folder:
            self.settings.setValue("last_dir", folder)
            self.add_paths([Path(folder)])

    def add_paths(self, paths: list[Path]) -> int:
        files = expand_inputs(paths)
        already = {item.data(COL_FILE, Qt.ItemDataRole.UserRole) for item in self._pending}
        added = 0
        for f in files:
            if str(f) in already:
                continue
            _kind, detail = describe(f)
            item = QTreeWidgetItem([f.name, detail, READY, ""])
            item.setToolTip(COL_FILE, str(f))
            item.setData(COL_FILE, Qt.ItemDataRole.UserRole, str(f))
            item.setIcon(COL_FILE, self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
            self.tree.addTopLevelItem(item)
            self._pending.append(item)
            added += 1
        skipped = sum(1 for p in paths if p.is_file() and not is_supported(p))
        if added == 0 and paths:
            self.summary.setText("No supported files found in what was dropped.")
        elif skipped > 0:
            self.summary.setText(f"Added {added} file(s); {skipped} unsupported file(s) skipped.")
        else:
            self._update_summary()
        self._refresh_buttons()
        return added

    # -- running --------------------------------------------------------------
    def start(self) -> None:
        if not self._pending:
            return
        out_dir = self._current_out_dir()
        if out_dir is not None and not out_dir.is_dir():
            QMessageBox.warning(self, APP_NAME, f"The folder {out_dir} doesn't exist. Please choose another.")
            return
        if not self.queue.busy:
            self._batch_outputs, self._batch_size = [], 0
        for item in self._pending:
            src = Path(item.data(COL_FILE, Qt.ItemDataRole.UserRole))
            job = self.queue.submit(src, out_dir)
            self._items[job.id] = item
            self._set_state(item, State.QUEUED)
            self._batch_size += 1
        self._pending.clear()
        self._save_settings()
        self._refresh_buttons()
        self._update_summary()

    def _cancel(self) -> None:
        self.queue.cancel_all()

    def _from_worker(self, job: Job) -> None:
        # Runs on the worker thread: copy plain values and hop to the GUI thread.
        out = str(job.result.output) if job.result else ""
        detail = job.result.detail if job.result else ""
        self._bridge.updated.emit(job.id, job.state.value, job.message, out, detail)

    def _on_job_updated(self, job_id: int, state: str, message: str, output: str, detail: str) -> None:
        item = self._items.get(job_id)
        if item is None:
            return
        st = State(state)
        self._set_state(item, st)
        item.setText(COL_NOTES, message)
        item.setToolTip(COL_NOTES, message)
        if detail:
            item.setText(COL_TYPE, detail)
        if output:
            item.setData(COL_STATUS, Qt.ItemDataRole.UserRole, output)
            item.setToolTip(COL_STATUS, f"Saved as {output}\nDouble-click to open")
            self._batch_outputs.append(output)
        if st in (State.DONE, State.WARNING, State.FAILED, State.CANCELLED):
            self._maybe_finish_batch()
        self._update_summary()
        self._refresh_buttons()

    def _maybe_finish_batch(self) -> None:
        if self.queue.busy:
            return
        if self.open_when_done.isChecked() and self._batch_outputs:
            if len(self._batch_outputs) == 1:
                _open_path(Path(self._batch_outputs[0]))
            else:
                folders = {str(Path(o).parent) for o in self._batch_outputs}
                for folder in sorted(folders)[:3]:  # don't open dozens of windows
                    _open_path(Path(folder))
        self._batch_outputs, self._batch_size = [], 0

    def _set_state(self, item: QTreeWidgetItem, state: State) -> None:
        item.setText(COL_STATUS, state.value)
        color = STATE_COLORS.get(state)
        item.setForeground(COL_STATUS, QColor(color) if color else self.palette().text().color())
        icons = {
            State.DONE: QStyle.StandardPixmap.SP_DialogApplyButton,
            State.WARNING: QStyle.StandardPixmap.SP_MessageBoxWarning,
            State.FAILED: QStyle.StandardPixmap.SP_MessageBoxCritical,
            State.WORKING: QStyle.StandardPixmap.SP_BrowserReload,
        }
        if state in icons:
            item.setIcon(COL_STATUS, self.style().standardIcon(icons[state]))
        else:
            item.setIcon(COL_STATUS, QIcon())

    def _update_summary(self) -> None:
        counts: dict[str, int] = {}
        for i in range(self.tree.topLevelItemCount()):
            s = self.tree.topLevelItem(i).text(COL_STATUS)
            counts[s] = counts.get(s, 0) + 1
        wording = [
            (READY, "ready"), (State.QUEUED.value, "waiting"), (State.WORKING.value, "converting"),
            (State.DONE.value, "done"), (State.WARNING.value, "to check"), (State.FAILED.value, "failed"),
        ]
        parts = [f"{counts[label]} {word}" for label, word in wording if counts.get(label)]
        self.summary.setText(" · ".join(parts))

    def _refresh_buttons(self) -> None:
        self.b_convert.setEnabled(bool(self._pending))
        self.b_cancel.setEnabled(self.queue.busy)
        done = any(
            self.tree.topLevelItem(i).text(COL_STATUS)
            in (State.DONE.value, State.WARNING.value, State.FAILED.value, State.CANCELLED.value)
            for i in range(self.tree.topLevelItemCount())
        )
        self.b_clear.setEnabled(done)

    # -- row actions ------------------------------------------------------------
    def _output_of(self, item: QTreeWidgetItem) -> Path | None:
        out = item.data(COL_STATUS, Qt.ItemDataRole.UserRole)
        return Path(out) if out else None

    def _open_output(self, item: QTreeWidgetItem) -> None:
        out = self._output_of(item)
        if out and out.exists():
            _open_path(out)

    def _context_menu(self, pos) -> None:
        item = self.tree.itemAt(pos)
        if item is None:
            return
        menu = QMenu(self)
        out = self._output_of(item)
        src = Path(item.data(COL_FILE, Qt.ItemDataRole.UserRole))
        a_open = QAction("Open Word file", self, enabled=bool(out and out.exists()))
        a_open.triggered.connect(lambda: self._open_output(item))
        a_show = QAction("Show in folder", self, enabled=bool(out and out.exists()))
        a_show.triggered.connect(lambda: _reveal(out))
        a_orig = QAction("Open original", self)
        a_orig.triggered.connect(lambda: _open_path(src))
        a_remove = QAction("Remove from list", self, enabled=item.text(COL_STATUS) not in
                           (State.QUEUED.value, State.WORKING.value))
        a_remove.triggered.connect(lambda: self._remove(item))
        for a in (a_open, a_show, a_orig):
            menu.addAction(a)
        menu.addSeparator()
        menu.addAction(a_remove)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def _remove(self, item: QTreeWidgetItem) -> None:
        if item in self._pending:
            self._pending.remove(item)
        self._items = {k: v for k, v in self._items.items() if v is not item}
        self.tree.takeTopLevelItem(self.tree.indexOfTopLevelItem(item))
        self._update_summary()
        self._refresh_buttons()

    def _clear_finished(self) -> None:
        finished = (State.DONE.value, State.WARNING.value, State.FAILED.value, State.CANCELLED.value)
        for i in reversed(range(self.tree.topLevelItemCount())):
            item = self.tree.topLevelItem(i)
            if item.text(COL_STATUS) in finished:
                self._remove(item)

    def _about(self) -> None:
        from ..logsetup import log_dir

        QMessageBox.about(
            self,
            f"About {APP_NAME}",
            f"<b>{APP_NAME} {__version__}</b><br>Converts files to Word documents on this PC. "
            "Nothing is uploaded anywhere.<br><br>"
            f"Log files: {log_dir()}",
        )

    # -- window-level drag & drop + close ----------------------------------------
    def dragEnterEvent(self, event: QDragEnterEvent):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):  # noqa: N802
        paths = [Path(u.toLocalFile()) for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.add_paths(paths)

    def closeEvent(self, event):  # noqa: N802
        if self.queue.busy:
            answer = QMessageBox.question(
                self, APP_NAME, "Files are still converting. Stop and close?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self._save_settings()
        self.queue.shutdown()
        event.accept()


def _open_path(path: Path) -> None:
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


def _reveal(path: Path | None) -> None:
    if not path:
        return
    if sys.platform == "win32":
        subprocess.Popen(["explorer", "/select,", str(path)])  # noqa: S603,S607
    else:
        _open_path(path.parent)

