# Invoice to Excel — חשבונית לאקסל

A simple Windows desktop application for non-technical Hebrew-speaking users.  
Upload a PDF, Excel, or CSV invoice → map columns → export a clean Excel file.

---

## Supported input formats

| Format | Notes |
|--------|-------|
| PDF (`.pdf`) | Tables are extracted automatically. Scanned PDFs require OCR (not yet supported — see below). |
| Excel (`.xlsx`, `.xls`) | First sheet loaded by default; multi-sheet files show a sheet selector. |
| CSV (`.csv`) | Handles UTF-8, Windows-1255, and most common Hebrew encodings. |

**Target output columns (in fixed order):**

1. קוד
2. תיאור
3. תמונה
4. כמות
5. מחיר יחידה כולל מע״מ
6. אחוז הנחה
7. סה״כ

---

## Running in development

### Requirements

- Python 3.9 or higher
- `pip`

### Install dependencies

```bash
pip install -r requirements.txt
```

### Start the app

```bash
python app.py
```

---

## Building a Windows EXE

> Run this on a **Windows** machine. The resulting `.exe` only works on Windows.

1. Open a Command Prompt in the project folder.
2. Run:

```bat
build_exe.bat
```

The script will:
1. Install all Python dependencies via `pip`.
2. Run PyInstaller to produce a single-file executable.

### Where the EXE is created

```
dist\InvoiceToExcel.exe
```

---

## Sending the EXE to another Windows user

### Recommended method: cloud storage

Email clients and some messaging apps **block or strip `.exe` attachments** because they can contain malware.  
Windows Defender / SmartScreen may also show a warning when the file is first run.

**Recommended transfer methods:**

| Service | How |
|---------|-----|
| Google Drive | Upload → Share link |
| WeTransfer | No account needed, up to 2 GB |
| Dropbox | Upload → Copy link |
| OneDrive | Upload → Share |

### The recipient's first run

Windows SmartScreen may show: *"Windows protected your PC"*.  
Click **"More info"** → **"Run anyway"** to proceed.  
This warning appears because the EXE is not digitally signed (code signing requires a paid certificate).

---

## Project structure

```
invoice_to_excel/
├── app.py               # Tkinter UI and main entry point
├── requirements.txt     # Python dependencies
├── build_exe.bat        # One-click Windows build script
├── README.md
├── .gitignore
└── src/
    ├── __init__.py
    ├── file_loader.py   # PDF / Excel / CSV loading
    ├── mapper.py        # Target schema + column mapping
    └── exporter.py      # Styled Excel export
```

---

## Adding OCR support (future)

For scanned PDFs, replace or extend the `_load_pdf` function in `src/file_loader.py`.  
A drop-in integration point:

```python
# src/file_loader.py  –  _load_pdf()
# If pdfplumber finds no tables, fall through to an OCR engine, e.g.:
#   import pytesseract, pdf2image
#   pages = pdf2image.convert_from_path(path)
#   text  = pytesseract.image_to_data(pages[0], output_type=Output.DATAFRAME, lang='heb')
```

No other files need to change.

---

## License

MIT
