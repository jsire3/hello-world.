"""Drive the installed Microsoft Office (Word, Excel, PowerPoint) in the background.

Design notes
------------
* We start our **own** hidden Word instance with ``DispatchEx`` so we never
  touch documents the user has open in their own Word window.
* Every dialog Word could show is suppressed: alerts off, macros disabled,
  PDF-conversion warning disabled, a dummy password so protected files fail
  instead of prompting.
* Word can hang on a bad file. The worker thread runs the COM call, and a
  watchdog (see ``OfficeSession.kill``) can terminate our Word process; the
  blocked COM call then raises and the session is restarted for the next job.
* The Word process runs at below-normal priority so the PC stays responsive.

This module only works on Windows with Office installed. It imports
``win32com`` lazily so the rest of the app (and the tests) run anywhere.
"""

from __future__ import annotations

import logging
import os
import shutil
import sys
import tempfile
import threading
import time
from pathlib import Path

from .base import ConversionError

log = logging.getLogger(__name__)

# Word constants (from the Word object model).
WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE = 0
WD_FORMAT_DOCUMENT_DEFAULT = 16  # .docx
WD_CURRENT_COMPAT = 65535        # wdCurrent - newest compatibility mode
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Excel / PowerPoint export constants.
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
PP_SAVE_AS_PDF = 32
MSO_FALSE = 0

# Passed as the document password so protected files raise instead of prompting.
DUMMY_PASSWORD = "__doc2word_no_password__"

# Restart Word after this many documents to keep memory use flat.
RECYCLE_AFTER = 25


def _require_windows() -> None:
    if sys.platform != "win32":
        raise ConversionError("Microsoft Office conversion only works on Windows.")


def disable_pdf_reflow_prompt() -> None:
    """Turn off Word's 'Word will now convert your PDF...' confirmation.

    This is a per-user Word setting (HKCU), the same one the checkbox in the
    dialog sets. It's harmless and makes unattended PDF conversion possible.
    """
    if sys.platform != "win32":
        return
    import winreg

    for version in ("16.0", "15.0"):  # Office 2016+ / 365, and Office 2013
        key_path = rf"Software\Microsoft\Office\{version}\Word\Options"
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, 1)
        except OSError:
            log.debug("Could not set PDF warning key for Office %s", version)


def _office_pids(image_name: str) -> set[int]:
    """PIDs of Office processes started by automation (not the user's own)."""
    import psutil

    pids = set()
    for proc in psutil.process_iter(["name", "cmdline"]):
        try:
            if (proc.info["name"] or "").lower() != image_name:
                continue
            cmd = " ".join(proc.info["cmdline"] or []).lower()
            if "/automation" in cmd or "-embedding" in cmd:
                pids.add(proc.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


def _lower_priority(pid: int) -> None:
    try:
        import psutil

        psutil.Process(pid).nice(psutil.BELOW_NORMAL_PRIORITY_CLASS)
    except Exception:  # noqa: BLE001 - best effort only
        pass


def _com_error_message(exc: Exception) -> str:
    """Turn a pywintypes.com_error into something a person can read."""
    args = getattr(exc, "args", ())
    # com_error args: (hresult, text, (wcode, source, description, ...), argerr)
    if len(args) >= 3 and args[2] and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return str(exc)


class _App:
    """One automated Office application and the PID that backs it."""

    def __init__(self, prog_id: str, image_name: str):
        import win32com.client

        before = _office_pids(image_name)
        self.app = win32com.client.DispatchEx(prog_id)
        self.pid: int | None = None
        # Give Windows a moment to register the new process.
        for _ in range(20):
            new = _office_pids(image_name) - before
            if new:
                self.pid = min(new)
                break
            time.sleep(0.1)
        if self.pid:
            _lower_priority(self.pid)
        self.uses = 0

    def quit(self) -> None:
        try:
            self.app.Quit()
        except Exception:  # noqa: BLE001 - already gone or hung
            pass
        self.app = None

    def kill(self) -> None:
        if not self.pid:
            return
        try:
            import psutil

            psutil.Process(self.pid).kill()
        except Exception:  # noqa: BLE001
            pass


class OfficeSession:
    """Owns hidden Word/Excel/PowerPoint instances for one worker thread.

    Must be created, used and closed on the same thread (COM apartment rules).
    """

    def __init__(self) -> None:
        _require_windows()
        import pythoncom

        pythoncom.CoInitialize()
        disable_pdf_reflow_prompt()
        self._word: _App | None = None
        self._excel: _App | None = None
        self._ppt: _App | None = None
        self._lock = threading.Lock()  # guards the app references for kill()

    # -- application lifecycle -------------------------------------------
    def _get_word(self):
        if self._word and self._word.uses >= RECYCLE_AFTER:
            self._close_app("_word")
        if self._word is None:
            app = _App("Word.Application", "winword.exe")
            w = app.app
            w.Visible = False
            w.DisplayAlerts = WD_ALERTS_NONE
            w.ScreenUpdating = False
            try:
                w.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:  # noqa: BLE001 - older Word
                pass
            # Note: never change w.Options here - those are the user's own
            # Word settings and persist after we quit.
            with self._lock:
                self._word = app
        self._word.uses += 1
        return self._word.app

    def _get_excel(self):
        if self._excel is None:
            app = _App("Excel.Application", "excel.exe")
            app.app.Visible = False
            app.app.DisplayAlerts = False
            try:
                app.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:  # noqa: BLE001
                pass
            with self._lock:
                self._excel = app
        return self._excel.app

    def _get_ppt(self):
        if self._ppt is None:
            app = _App("PowerPoint.Application", "powerpnt.exe")
            # PowerPoint refuses Visible = False on the app; we open windowless instead.
            try:
                app.app.DisplayAlerts = 1  # ppAlertsNone
            except Exception:  # noqa: BLE001
                pass
            with self._lock:
                self._ppt = app
        return self._ppt.app

    def _close_app(self, attr: str) -> None:
        with self._lock:
            app: _App | None = getattr(self, attr)
            setattr(self, attr, None)
        if app:
            app.quit()
            app.kill()  # make sure nothing is left behind

    def kill(self) -> None:
        """Called from the watchdog thread when a conversion hangs."""
        with self._lock:
            apps = [a for a in (self._word, self._excel, self._ppt) if a]
        for app in apps:
            app.kill()

    def reset(self) -> None:
        """Drop all app references after a crash/kill so they restart cleanly."""
        for attr in ("_word", "_excel", "_ppt"):
            self._close_app(attr)

    def close(self) -> None:
        self.reset()
        import pythoncom

        pythoncom.CoUninitialize()

    # -- conversions ---------------------------------------------------------
    def word_to_docx(self, src: Path, dest: Path) -> None:
        """Open anything Word understands (incl. PDF via reflow) and save as .docx."""
        word = self._get_word()
        doc = None
        try:
            doc = word.Documents.Open(
                FileName=str(src),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=DUMMY_PASSWORD,
                Revert=False,
                Visible=False,
                NoEncodingDialog=True,
            )
            doc.SaveAs2(
                FileName=str(dest),
                FileFormat=WD_FORMAT_DOCUMENT_DEFAULT,
                AddToRecentFiles=False,
                CompatibilityMode=WD_CURRENT_COMPAT,
            )
        except Exception as exc:  # noqa: BLE001
            raise ConversionError(_friendly(exc)) from exc
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE)
                except Exception:  # noqa: BLE001
                    pass

    def spreadsheet_to_pdf(self, src: Path, pdf: Path) -> None:
        excel = self._get_excel()
        wb = None
        try:
            wb = excel.Workbooks.Open(
                str(src), UpdateLinks=0, ReadOnly=True,
                Password=DUMMY_PASSWORD, AddToMru=False,
            )
            wb.ExportAsFixedFormat(
                Type=XL_TYPE_PDF, Filename=str(pdf), Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=False, IgnorePrintAreas=False, OpenAfterPublish=False,
            )
        except Exception as exc:  # noqa: BLE001
            raise ConversionError(_friendly(exc)) from exc
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass

    def presentation_to_pdf(self, src: Path, pdf: Path) -> None:
        ppt = self._get_ppt()
        pres = None
        try:
            pres = ppt.Presentations.Open(
                str(src), ReadOnly=True, Untitled=False, WithWindow=MSO_FALSE,
            )
            pres.SaveAs(str(pdf), PP_SAVE_AS_PDF)
        except Exception as exc:  # noqa: BLE001
            raise ConversionError(_friendly(exc)) from exc
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:  # noqa: BLE001
                    pass


def _friendly(exc: Exception) -> str:
    msg = _com_error_message(exc)
    low = msg.lower()
    if "password" in low:
        return "The file is password protected."
    if "rpc server" in low or "disconnected" in low or "-2147023174" in low:
        return "Office stopped responding while converting this file."
    if "corrupt" in low or "unreadable" in low:
        return "The file appears to be damaged."
    return msg or "Office could not convert this file."


class TempDir:
    """A private temp folder for intermediate files, removed afterwards."""

    def __enter__(self) -> Path:
        self.path = Path(tempfile.mkdtemp(prefix="doc2word_"))
        return self.path

    def __exit__(self, *exc) -> None:
        shutil.rmtree(self.path, ignore_errors=True)


def office_available() -> bool:
    """True if Word is registered on this PC."""
    if sys.platform != "win32":
        return False
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID").Close()
        return True
    except OSError:
        return False


# Re-exported for callers that only need to know the platform.
IS_WINDOWS = os.name == "nt"
