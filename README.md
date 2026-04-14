# PDF2Word Pro

> Convert PDF files to editable Microsoft Word documents — with formatting preserved.

[![Build & Release](https://github.com/ratnadipsinha/pdftoword/actions/workflows/build-release.yml/badge.svg)](https://github.com/ratnadipsinha/pdftoword/actions/workflows/build-release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)

---

## Download

**No Python needed** — grab the latest `.exe` from [Releases](../../releases/latest) and double-click to run.

---

## Features

| Feature | Details |
|---------|---------|
| Font fidelity | Name, size, bold, italic, color preserved |
| Layout | Tables, columns, headings, alignment |
| Images | Embedded images at correct position/size |
| OCR | Scanned/image PDFs converted via Tesseract |
| Batch mode | Convert an entire folder of PDFs at once |
| Privacy | 100% local — no cloud upload |

---

## Screenshots

| Drop zone | Conversion log |
|-----------|---------------|
| _(drag & drop PDFs, pick output folder)_ | _(per-file status, progress bar)_ |

---

## Usage

### GUI (double-click)
1. Download `PDF2Word-Pro-vX.X.X-Windows.exe` from [Releases](../../releases/latest)
2. Double-click to launch
3. Drag PDF files onto the drop zone (or click to browse)
4. Select an output folder
5. Click **Convert to Word**

### CLI

```bash
# Single file
"PDF2Word Pro.exe" report.pdf

# Single file with custom output
"PDF2Word Pro.exe" report.pdf C:\output\report.docx

# Batch — convert all PDFs in a folder
"PDF2Word Pro.exe" --batch C:\my-pdfs\ C:\output\
```

---

## Run from Source

```bash
# 1. Clone
git clone https://github.com/ratnadipsinha/pdftoword.git
cd pdftoword

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch GUI
python main.py

# 4. Or use CLI
python main.py invoice.pdf
python main.py --batch ./pdfs/ ./output/
```

### OCR support (scanned PDFs)

Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) and make sure it is on your `PATH`.
The app auto-detects scanned pages and runs OCR on them.

---

## Build .exe locally

```bash
pip install pyinstaller
pyinstaller pdf2word.spec --clean --noconfirm
# Output: dist/PDF2Word Pro.exe
```

---

## How it works

```
PDF file
  └── PDFAnalyzer (PyMuPDF + pdfplumber)
        ├── Text blocks  → fonts, sizes, colors, alignment
        ├── Tables       → rows, cells, borders
        ├── Images       → extracted at original resolution
        └── Scanned?     → Tesseract OCR
              ↓
        WordBuilder (python-docx)
              ↓
        .docx file
```

---

## Tech stack

| Library | Purpose |
|---------|---------|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF parsing, image extraction, font metadata |
| [pdfplumber](https://github.com/jsvine/pdfplumber) | Table detection and extraction |
| [python-docx](https://python-docx.readthedocs.io/) | Word document generation |
| [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) | Text recognition in scanned PDFs |
| [Pillow](https://pillow.readthedocs.io/) | Image processing |
| Tkinter | Desktop GUI |
| PyInstaller | Single-file .exe packaging |

---

## License

MIT — free for personal and commercial use.

---

## Contributing

Pull requests welcome. Please open an issue first to discuss major changes.
