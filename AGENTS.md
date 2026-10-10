# Project context

This is a small Python/Tkinter desktop app for church donation QR entry into an existing XLSX workbook. Read README.md before changes. Run with `.venv/bin/python main.py`; run checks with `.venv/bin/python -m unittest discover -s tests -v`.

- Keep donor data local. Do not commit the supplied XLSX/JPEG or backups without an explicit request.
- Test workbook writes on temporary copies. Do not alter the supplied reference file as part of code verification.
- Discover sheet, name, year, QR and date columns from actual workbook headers. Avoid hardcoded row limits, years or name columns.
- Match by validated category-ID QR, never by donor name. Reject ambiguous duplicates.
- Guard existing deposits, missing date columns, formula destinations and external file changes.
- Preserve untouched XLSX ZIP members, especially external links, styles, comments and VML. Save with backup and atomic replacement.
- Keep Tk widget operations on the main thread; run slow file/image work on the serial executor.
- Do not add subagents or dependencies unless the task needs them. Preserve the existing `DateFucntion.py` import name for compatibility.
