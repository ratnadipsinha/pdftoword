"""
Converter — orchestrates PDF analysis → Word building.
Supports single file and batch folder conversion.
"""

import os
import time
from pathlib import Path
from typing import Callable, Optional

from .analyzer import PDFAnalyzer
from .builder import WordBuilder


class ConversionResult:
    def __init__(self, input_path: str, output_path: str, success: bool,
                 error: str = "", duration: float = 0.0, page_count: int = 0):
        self.input_path = input_path
        self.output_path = output_path
        self.success = success
        self.error = error
        self.duration = duration
        self.page_count = page_count

    def __repr__(self):
        status = "OK" if self.success else f"FAIL({self.error})"
        return f"<ConversionResult {Path(self.input_path).name} → {status} in {self.duration:.1f}s>"


class PDF2WordConverter:
    """
    Main converter. Usage:

        converter = PDF2WordConverter()
        result = converter.convert("input.pdf", "output.docx")
        results = converter.convert_batch("./pdfs/", "./output/")
    """

    def __init__(self, ocr_lang: str = "eng", progress_cb: Optional[Callable] = None):
        """
        Args:
            ocr_lang:    Tesseract language code (e.g. "eng", "fra", "deu").
            progress_cb: Optional callback(current, total, filename) for progress updates.
        """
        self.ocr_lang = ocr_lang
        self.progress_cb = progress_cb

    # ------------------------------------------------------------------
    # Single file conversion
    # ------------------------------------------------------------------

    def convert(self, pdf_path: str, output_path: Optional[str] = None) -> ConversionResult:
        """
        Convert a single PDF to .docx.

        Args:
            pdf_path:    Path to the source PDF.
            output_path: Destination .docx path. If None, uses same dir as PDF.

        Returns:
            ConversionResult
        """
        pdf_path = str(Path(pdf_path).resolve())

        if not os.path.isfile(pdf_path):
            return ConversionResult(pdf_path, "", False, error=f"File not found: {pdf_path}")

        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix(".docx"))
        else:
            output_path = str(Path(output_path).resolve())

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        start = time.time()
        try:
            with PDFAnalyzer(pdf_path) as analyzer:
                pages = analyzer.analyze()

            builder = WordBuilder(pages, ocr_lang=self.ocr_lang)
            builder.build(output_path)

            duration = time.time() - start
            return ConversionResult(
                input_path=pdf_path,
                output_path=output_path,
                success=True,
                duration=duration,
                page_count=len(pages),
            )

        except Exception as exc:
            duration = time.time() - start
            return ConversionResult(
                input_path=pdf_path,
                output_path=output_path,
                success=False,
                error=str(exc),
                duration=duration,
            )

    # ------------------------------------------------------------------
    # Batch conversion
    # ------------------------------------------------------------------

    def convert_batch(
        self,
        input_dir: str,
        output_dir: str,
        recursive: bool = False,
    ) -> list[ConversionResult]:
        """
        Convert all PDFs in input_dir to output_dir.

        Args:
            input_dir:  Folder containing PDF files.
            output_dir: Folder for output .docx files (created if missing).
            recursive:  If True, recurse into subdirectories.

        Returns:
            List of ConversionResult
        """
        input_dir = Path(input_dir)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        pattern = "**/*.pdf" if recursive else "*.pdf"
        pdf_files = sorted(input_dir.glob(pattern))

        if not pdf_files:
            return []

        results = []
        total = len(pdf_files)

        for idx, pdf_file in enumerate(pdf_files, start=1):
            # Mirror subfolder structure in output
            rel = pdf_file.relative_to(input_dir)
            out_file = output_dir / rel.with_suffix(".docx")
            out_file.parent.mkdir(parents=True, exist_ok=True)

            if self.progress_cb:
                self.progress_cb(idx, total, pdf_file.name)

            result = self.convert(str(pdf_file), str(out_file))
            results.append(result)

        return results

    # ------------------------------------------------------------------
    # Summary helper
    # ------------------------------------------------------------------

    @staticmethod
    def print_summary(results: list[ConversionResult]):
        ok = [r for r in results if r.success]
        fail = [r for r in results if not r.success]
        print(f"\n{'='*50}")
        print(f"Conversion Summary: {len(ok)}/{len(results)} succeeded")
        print(f"{'='*50}")
        for r in ok:
            print(f"  [OK]   {Path(r.input_path).name}  ({r.page_count}p, {r.duration:.1f}s)")
        for r in fail:
            print(f"  [FAIL] {Path(r.input_path).name}  — {r.error}")
        print()
