# Changelog

## 1.1.0 — 2026-11-14

- Merged the two near-identical applications into one: a shared core
  (`converter.py`, no GUI dependencies) plus a single GUI app
  (`word_number_converter_gui.py`) with a direction selector for
  English → Persian and Persian/Arabic → English. Fixes now only need
  to be made once.
- Coverage extended: text boxes are now converted, and footnotes,
  endnotes and comments are converted when the document has them.
  The README now states precisely what is and is not converted
  (field results such as page numbers and TOC are not).
- File dialog and drag & drop now accept only `.docx` files
  (previously any file could be selected and only failed later, at
  conversion time).
- Added `requirements.txt` and a unit-test suite
  (`tests/test_converter.py`) covering both directions, Arabic
  digits, idempotency, output naming, tables, headers and text
  boxes, plus a full round-trip.
- Folder and file names are now plain English (the Persian folder
  names with spaces caused quoting/URL friction).
- README fixed: the Project Structure section previously listed two
  script names that did not exist.

## 1.0.0 — 2026-11-11

- Initial release: two separate GUI applications
  (English → Persian and Persian → English) with .exe builds,
  format preservation, table and header/footer support, and
  drag & drop.
