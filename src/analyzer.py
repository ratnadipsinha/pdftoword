"""
PDF Analyzer — detects PDF type and extracts structured content blocks.
Handles: text-based PDFs and scanned/image PDFs (via OCR flag).
"""

import fitz  # PyMuPDF
import pdfplumber
from dataclasses import dataclass, field
from typing import Optional
import os
import tempfile
from PIL import Image
import io


@dataclass
class TextBlock:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    font_name: str = "Arial"
    font_size: float = 11.0
    bold: bool = False
    italic: bool = False
    color: tuple = (0, 0, 0)     # RGB 0-255
    align: str = "left"           # left | center | right | justify
    page_num: int = 0


@dataclass
class TableBlock:
    rows: list                    # list of list of str (cell text)
    x0: float = 0
    y0: float = 0
    x1: float = 0
    y1: float = 0
    page_num: int = 0


@dataclass
class ImageBlock:
    image_bytes: bytes
    x0: float
    y0: float
    x1: float
    y1: float
    width_pt: float
    height_pt: float
    page_num: int = 0
    ext: str = "png"


@dataclass
class PageData:
    page_num: int
    width: float
    height: float
    text_blocks: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    images: list = field(default_factory=list)
    is_scanned: bool = False


class PDFAnalyzer:
    """Analyzes a PDF and returns structured page data."""

    SCANNED_THRESHOLD = 0.1   # If text char density < this, treat as scanned

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._fitz_doc = fitz.open(pdf_path)

    def close(self):
        self._fitz_doc.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def analyze(self) -> list[PageData]:
        """Return a list of PageData, one per PDF page."""
        pages = []
        with pdfplumber.open(self.pdf_path) as plumber_doc:
            for page_num, (plumber_page, fitz_page) in enumerate(
                zip(plumber_doc.pages, self._fitz_doc)
            ):
                page_data = self._analyze_page(page_num, plumber_page, fitz_page)
                pages.append(page_data)
        return pages

    # ------------------------------------------------------------------
    # Per-page analysis
    # ------------------------------------------------------------------

    def _analyze_page(self, page_num: int, plumber_page, fitz_page) -> PageData:
        width = float(plumber_page.width)
        height = float(plumber_page.height)
        page_data = PageData(page_num=page_num, width=width, height=height)

        # Detect scanned pages
        page_data.is_scanned = self._is_scanned(fitz_page)

        if page_data.is_scanned:
            # OCR path — extract as one image, mark for OCR in builder
            img_block = self._extract_page_as_image(fitz_page, page_num)
            if img_block:
                page_data.images.append(img_block)
        else:
            # Extract tables first (so we can skip those regions for text)
            page_data.tables = self._extract_tables(plumber_page, page_num)
            table_regions = [(t.x0, t.y0, t.x1, t.y1) for t in page_data.tables]

            # Extract text blocks (skip table areas)
            page_data.text_blocks = self._extract_text_blocks(
                fitz_page, page_num, table_regions
            )

            # Extract embedded images
            page_data.images = self._extract_images(fitz_page, page_num)

        return page_data

    # ------------------------------------------------------------------
    # Scanned detection
    # ------------------------------------------------------------------

    def _is_scanned(self, fitz_page) -> bool:
        """Return True if the page has almost no selectable text."""
        text = fitz_page.get_text("text").strip()
        page_area = fitz_page.rect.width * fitz_page.rect.height
        if page_area == 0:
            return True
        char_density = len(text) / page_area
        return char_density < self.SCANNED_THRESHOLD

    # ------------------------------------------------------------------
    # Text extraction
    # ------------------------------------------------------------------

    def _extract_text_blocks(
        self, fitz_page, page_num: int, skip_regions: list
    ) -> list[TextBlock]:
        """Extract styled text blocks using PyMuPDF dict mode."""
        blocks = []
        raw = fitz_page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in raw.get("blocks", []):
            if block.get("type") != 0:   # 0 = text block
                continue

            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if not text:
                        continue

                    bbox = span.get("bbox", (0, 0, 0, 0))
                    x0, y0, x1, y1 = bbox

                    # Skip if inside a table region
                    if self._in_region((x0, y0, x1, y1), skip_regions):
                        continue

                    font_name = span.get("font", "Arial")
                    font_size = span.get("size", 11.0)
                    flags = span.get("flags", 0)

                    # PyMuPDF flags: bit 1=italic, bit 4=bold
                    bold = bool(flags & (1 << 4))
                    italic = bool(flags & (1 << 1))

                    # Color is int 0xRRGGBB
                    color_int = span.get("color", 0)
                    color = self._int_to_rgb(color_int)

                    blocks.append(TextBlock(
                        text=text,
                        x0=x0, y0=y0, x1=x1, y1=y1,
                        font_name=font_name,
                        font_size=round(font_size, 1),
                        bold=bold,
                        italic=italic,
                        color=color,
                        align=self._detect_alignment(x0, x1, fitz_page.rect.width),
                        page_num=page_num,
                    ))

        return blocks

    # ------------------------------------------------------------------
    # Table extraction
    # ------------------------------------------------------------------

    def _extract_tables(self, plumber_page, page_num: int) -> list[TableBlock]:
        """Extract tables using pdfplumber's table finder."""
        result = []
        tables = plumber_page.extract_tables()
        table_settings = plumber_page.find_tables()

        for i, rows in enumerate(tables):
            if not rows:
                continue
            bbox = table_settings[i].bbox if i < len(table_settings) else (0, 0, 0, 0)
            cleaned_rows = [
                [cell if cell else "" for cell in row]
                for row in rows
            ]
            result.append(TableBlock(
                rows=cleaned_rows,
                x0=bbox[0], y0=bbox[1], x1=bbox[2], y1=bbox[3],
                page_num=page_num,
            ))
        return result

    # ------------------------------------------------------------------
    # Image extraction
    # ------------------------------------------------------------------

    def _extract_images(self, fitz_page, page_num: int) -> list[ImageBlock]:
        """Extract embedded images from a page."""
        images = []
        img_list = fitz_page.get_images(full=True)

        for img_info in img_list:
            xref = img_info[0]
            try:
                base_image = self._fitz_doc.extract_image(xref)
                image_bytes = base_image["image"]
                ext = base_image["ext"]

                # Get position via image rect
                rects = fitz_page.get_image_rects(xref)
                if not rects:
                    continue
                rect = rects[0]

                images.append(ImageBlock(
                    image_bytes=image_bytes,
                    x0=rect.x0, y0=rect.y0, x1=rect.x1, y1=rect.y1,
                    width_pt=rect.width,
                    height_pt=rect.height,
                    page_num=page_num,
                    ext=ext,
                ))
            except Exception:
                continue  # Skip unreadable images

        return images

    def _extract_page_as_image(self, fitz_page, page_num: int) -> Optional[ImageBlock]:
        """Render an entire page as an image (for scanned PDFs)."""
        try:
            mat = fitz.Matrix(2.0, 2.0)   # 2x zoom = ~144 dpi
            pix = fitz_page.get_pixmap(matrix=mat)
            image_bytes = pix.tobytes("png")
            w = fitz_page.rect.width
            h = fitz_page.rect.height
            return ImageBlock(
                image_bytes=image_bytes,
                x0=0, y0=0, x1=w, y1=h,
                width_pt=w, height_pt=h,
                page_num=page_num,
                ext="png",
            )
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _int_to_rgb(color_int: int) -> tuple:
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8) & 0xFF
        b = color_int & 0xFF
        return (r, g, b)

    @staticmethod
    def _detect_alignment(x0: float, x1: float, page_width: float) -> str:
        center = (x0 + x1) / 2
        page_center = page_width / 2
        left_margin = x0
        right_margin = page_width - x1

        if abs(center - page_center) < 30 and left_margin > 50:
            return "center"
        if right_margin < 30 and left_margin > 50:
            return "right"
        return "left"

    @staticmethod
    def _in_region(bbox: tuple, regions: list) -> bool:
        x0, y0, x1, y1 = bbox
        for rx0, ry0, rx1, ry1 in regions:
            if x0 >= rx0 - 5 and y0 >= ry0 - 5 and x1 <= rx1 + 5 and y1 <= ry1 + 5:
                return True
        return False
