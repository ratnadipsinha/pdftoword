# PDF2Word Pro

> Convert PDF files to editable Microsoft Word documents — fonts, tables and images preserved.

[![CI](https://github.com/ratnadipsinha/pdftoword/actions/workflows/build-release.yml/badge.svg)](https://github.com/ratnadipsinha/pdftoword/actions/workflows/build-release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://python.org)

---

## Live Demo

**[pdftoword-6-xc09.onrender.com](https://pdftoword-6-xc09.onrender.com)** — open in any browser, no install needed.

---

## Run locally in 3 commands

```bash
git clone https://github.com/ratnadipsinha/pdftoword.git
cd pdftoword
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:5000** in your browser.

---

## How to use

1. Drag & drop (or click to browse) your PDF
2. Optionally type a local folder path to save directly — or leave empty to download
3. Pick OCR language (for scanned PDFs)
4. Click **Convert to Word**
5. Download your `.docx` — or find it in the folder you specified

---

## Features

| Feature | Detail |
|---------|--------|
| Font fidelity | Name, size, bold, italic, color |
| Layout | Tables, columns, headings, alignment |
| Images | Extracted and positioned correctly |
| OCR | Scanned/image-only PDFs via Tesseract |
| Batch CLI | `python main.py --batch ./pdfs/ ./output/` |
| Privacy | 100% local — files never leave your machine |

---

## CLI (no browser needed)

```bash
# Single file
python main.py invoice.pdf

# Single file with custom output
python main.py invoice.pdf C:\output\invoice.docx

# Batch folder
python main.py --batch C:\my-pdfs\ C:\output\
```

---

## OCR for scanned PDFs

Install [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and add it to your PATH.
The app auto-detects image-only pages and runs OCR on them.

---

## Tech stack

| Library | Purpose |
|---------|---------|
| [Flask](https://flask.palletsprojects.com/) | Web server |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF parsing, fonts, images |
| [pdfplumber](https://github.com/jsvine/pdfplumber) | Table extraction |
| [python-docx](https://python-docx.readthedocs.io/) | Word document generation |
| [Tesseract](https://github.com/tesseract-ocr/tesseract) | OCR for scanned PDFs |
| [Pillow](https://pillow.readthedocs.io/) | Image handling |

---

## License

MIT — free for personal and commercial use.
