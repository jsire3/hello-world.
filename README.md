# Doc2Word

A Windows desktop app that converts PDFs, images and Office files into editable
Word documents, **entirely offline** on the user's PC. It drives the copy of
Microsoft Word already installed on each machine, so there is nothing to license.

## What it converts

| Input | How | Result |
|---|---|---|
| **PDF (digital)** | Word **and** pdf2docx at the same time; the better result is kept | Editable text, headings, tables, images |
| **PDF (scanned / mixed)** | Word's PDF reflow; each page is checked first | Typed pages editable; scanned pages kept as pictures *(OCR comes in Phase 3)* |
| **DOC, DOCX, RTF, ODT, TXT, HTML, MHT, XML, WPD** | Word opens and re-saves as .docx | Near-perfect |
| **XLS, XLSX, CSV, ODS** | Excel → PDF → Word | Tables as laid out for printing |
| **PPT, PPTX, ODP** | PowerPoint → PDF → Word | One page per slide |
| **JPG, PNG, TIFF (multi-page), BMP, GIF, WEBP, HEIC** | Placed on a correctly oriented page | Exact look *(editable text from images comes in Phase 3)* |

### How "best of both" works

Each digital PDF is converted by two engines in parallel: Word's PDF reflow and
[pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) (running in a separate
background process). Both results are scored against the original PDF:

| Check | Weight | What it catches |
|---|---|---|
| Text coverage | 35% | Missing text |
| Reading order | 25% | Multi-column pages that came out interleaved |
| Tables | 20% | Tables flattened into loose text or tabs |
| Images | 10% | Missing pictures and logos |
| Editability | 10% | Text trapped in floating text boxes |

The higher score wins; Word wins near-ties (within 2 points) because its output
uses cleaner styles. The **Engine** column shows which one was kept (hover for
both scores). To compare, right-click a file → **Save the Word/pdf2docx version too**.
The other version is kept for 7 days.

If one engine fails or hangs, the other one's result is used and the file is marked
**Done - check**. PDFs still convert on a PC without Word (pdf2docx only).
Use **PDF engine: Word only** for the fastest conversion of big batches.

## Using it

1. Drag files or whole folders onto the window (or click to browse).
2. Click **Convert to Word**.
3. The `.docx` is saved next to the original (or a folder you choose). Existing
   files are never overwritten: you get `report (1).docx` instead.

Status column: **Done**, **Done - check** (converted, but read the note, e.g.
scanned pages or text that may be missing), **Failed** (with the reason).
Double-click a row to open the result; right-click for more options.

The app never changes the original file, never uploads anything, and runs Word
hidden and separately from any Word window the user has open.

## Rolling it out (IT)

The installer is built automatically by GitHub Actions on every push
(**Actions → Build Windows installer → Artifacts**), or locally with
`packaging\build.ps1`.

```text
Doc2Word-Setup-0.2.0.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

* Installs per-machine into `C:\Program Files\Doc2Word` (admin rights needed), with a Start-menu shortcut.
* Works with Intune (Win32 app), Group Policy, SCCM or a shared-drive link.
* Uninstall: `"C:\Program Files\Doc2Word\unins000.exe" /VERYSILENT`.
* Requires Windows 10/11 x64 and Microsoft Word 2013 or later (Office 365 recommended).
  Excel/PowerPoint are only needed for spreadsheets/presentations.
* Logs: `%LOCALAPPDATA%\Doc2Word\logs\doc2word.log` (includes both engines' scores per PDF).
* The losing engine's versions are kept in `%LOCALAPPDATA%\Doc2Word\alternates` for 7 days, then deleted.
* The app sets one per-user Word option, `DisableConvertPdfWarning`, so PDFs convert
  without Word's "Word will now convert your PDF" prompt.

The installer isn't code-signed yet, so Windows SmartScreen may warn on first
run. Signing it with your company's certificate (or allow-listing it in Intune)
removes that.

A command-line converter is installed too, for batch jobs:

```text
"C:\Program Files\Doc2Word\doc2word-cli.exe" "\\server\scans" --out D:\Converted
```

## Developing

```bash
python -m venv .venv && .venv\Scripts\activate
pip install -e ".[dev]"
pytest -q                 # 34 tests; runs on Windows, Linux, macOS (Office is faked)
python -m doc2word        # start the app
```

Build the installer on Windows (needs Python 3.12 and [Inno Setup 6](https://jrsoftware.org/isdl.php)):

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build.ps1
```

### Layout

```text
doc2word/
  formats.py         file-type detection (extension + magic bytes), folder expansion
  pdfinspect.py      per-page digital vs scanned detection (ignores invisible OCR text)
  converter.py       routes each file to an engine; safe temp copies; never overwrites
  quality.py         text-coverage check (single-engine mode)
  scoring.py         scores each engine's .docx against the PDF and picks the best
  worker.py          background queue + watchdog that stops a hung Word after 5 min
  engines/office.py  hidden Word/Excel/PowerPoint automation (Windows only)
  engines/pdf2docx_engine.py  pdf2docx in a killable background process
  engines/image.py   images -> .docx pages
  gui/main_window.py PySide6 window
  cli.py, app.py     entry points
packaging/           PyInstaller spec, Inno Setup script, build script, icon
tests/               pytest suite with generated sample files and a fake Office
```

## Roadmap

- [x] **Phase 1:** app, drag-and-drop queue, Word-based conversion, installer, CI build
- [x] **Phase 2:** second PDF engine (pdf2docx) and automatic best-result selection
- [ ] **Phase 3:** offline OCR for scans and images (image clean-up, Docling layout and tables), "Exact look" mode
- [ ] **Phase 4:** Explorer right-click "Convert to Word", side-by-side preview, auto-update, low-confidence highlighting, OpenVINO acceleration
