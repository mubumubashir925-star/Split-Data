# Split Data Pro — Desktop (Phase 1)

Windows 10/11 desktop application built with Python and PySide6.

## Start

```powershell
cd desktop_app
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Phase 1 includes a modern navigable dashboard, file picker, drag/drop, workbook preview, and working basic row split to new XLSX workbooks. Combine, Cleanup and File Tools are marked as upcoming; original VBA sources in repository root are untouched. Files are processed locally. Split outputs are written to a folder you choose. Original workbook is never overwritten.
