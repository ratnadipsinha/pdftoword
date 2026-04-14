"""
Word Document Builder — converts extracted PageData into a .docx file.
Preserves: fonts, sizes, bold/italic/color, alignment, tables, images.
Handles OCR for scanned pages.
"""

import io
import os
import tempfile
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .analyzer import PageData, TextBlock, TableBlock, ImageBlock

try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class WordBuilder:
    """Builds a .docx document from a list of PageData objects."""

    # Font size thresholds for heading detection
    HEADING_SIZES = {
        24: 1,   # H1
        20: 1,
        18: 2,   # H2
        16: 2,
        14: 3,   # H3
    }

    def __init__(self, pages: list[PageData], ocr_lang: str = "eng"):
        self.pages = pages
        self.ocr_lang = ocr_lang
        self.doc = Document()
        self._setup_default_styles()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def build(self, output_path: str):
        """Build and save the Word document to output_path."""
        self._clear_default_paragraph()

        for page in self.pages:
            if page.is_scanned:
                self._handle_scanned_page(page)
            else:
                self._handle_text_page(page)

            # Add page break between pages (except last)
            if page.page_num < len(self.pages) - 1:
                self._add_page_break()

        self.doc.save(output_path)

    # ------------------------------------------------------------------
    # Page handlers
    # ------------------------------------------------------------------

    def _handle_text_page(self, page: PageData):
        """Process a text-based page: text blocks, tables, images."""
        # Sort all elements by vertical position (top to bottom)
        elements = []
        for tb in page.text_blocks:
            elements.append(("text", tb.y0, tb))
        for table in page.tables:
            elements.append(("table", table.y0, table))
        for img in page.images:
            elements.append(("image", img.y0, img))

        elements.sort(key=lambda e: e[1])

        for kind, _, obj in elements:
            if kind == "text":
                self._add_text_block(obj)
            elif kind == "table":
                self._add_table(obj)
            elif kind == "image":
                self._add_image(obj)

    def _handle_scanned_page(self, page: PageData):
        """OCR a scanned page and add extracted text, or embed as image."""
        if not page.images:
            return

        img_block = page.images[0]  # Full-page render

        if OCR_AVAILABLE:
            try:
                pil_img = Image.open(io.BytesIO(img_block.image_bytes))
                ocr_text = pytesseract.image_to_string(pil_img, lang=self.ocr_lang)
                if ocr_text.strip():
                    para = self.doc.add_paragraph(ocr_text.strip())
                    para.style = self.doc.styles["Normal"]
                    return
            except Exception:
                pass  # Fall through to image embed

        # If OCR failed or not available — embed image
        self._add_image(img_block)

    # ------------------------------------------------------------------
    # Text block rendering
    # ------------------------------------------------------------------

    def _add_text_block(self, tb: TextBlock):
        """Add a styled paragraph for a text span."""
        heading_level = self._detect_heading(tb)

        if heading_level:
            para = self.doc.add_heading("", level=heading_level)
        else:
            para = self.doc.add_paragraph()
            para.style = self.doc.styles["Normal"]

        para.alignment = ALIGN_MAP.get(tb.align, WD_ALIGN_PARAGRAPH.LEFT)

        run = para.add_run(tb.text)
        run.bold = tb.bold
        run.italic = tb.italic
        run.font.name = self._safe_font(tb.font_name)
        run.font.size = Pt(tb.font_size)

        r, g, b = tb.color
        if (r, g, b) != (0, 0, 0):
            run.font.color.rgb = RGBColor(r, g, b)

    def _detect_heading(self, tb: TextBlock) -> int | None:
        """Return heading level if text looks like a heading, else None."""
        if tb.font_size >= 24 and tb.bold:
            return 1
        if tb.font_size >= 18 and tb.bold:
            return 2
        if tb.font_size >= 14 and tb.bold:
            return 3
        if tb.bold and tb.font_size >= 13:
            return 4
        return None

    # ------------------------------------------------------------------
    # Table rendering
    # ------------------------------------------------------------------

    def _add_table(self, table_block: TableBlock):
        """Render a TableBlock as a Word table."""
        rows = table_block.rows
        if not rows:
            return

        num_cols = max(len(row) for row in rows)
        if num_cols == 0:
            return

        word_table = self.doc.add_table(rows=len(rows), cols=num_cols)
        word_table.style = "Table Grid"

        for r_idx, row in enumerate(rows):
            for c_idx, cell_text in enumerate(row):
                if c_idx >= num_cols:
                    break
                cell = word_table.cell(r_idx, c_idx)
                cell.text = cell_text or ""

                # Style the cell text
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)

        self.doc.add_paragraph()  # Space after table

    # ------------------------------------------------------------------
    # Image rendering
    # ------------------------------------------------------------------

    def _add_image(self, img_block: ImageBlock):
        """Embed an image in the document."""
        try:
            img_stream = io.BytesIO(img_block.image_bytes)

            # Convert width/height from points to inches (1 pt = 1/72 inch)
            width_in = img_block.width_pt / 72.0
            height_in = img_block.height_pt / 72.0

            # Cap to page width (6.5 inches for standard margins)
            max_width = 6.5
            if width_in > max_width:
                scale = max_width / width_in
                width_in = max_width
                height_in *= scale

            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(img_stream, width=Inches(width_in))
        except Exception:
            pass  # Skip unembeddable images

    # ------------------------------------------------------------------
    # Setup helpers
    # ------------------------------------------------------------------

    def _setup_default_styles(self):
        """Set global document defaults."""
        style = self.doc.styles["Normal"]
        font = style.font
        font.name = "Arial"
        font.size = Pt(11)

        # Narrow margins for better fidelity
        section = self.doc.sections[0]
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)

    def _clear_default_paragraph(self):
        """Remove the blank paragraph that python-docx adds by default."""
        for para in self.doc.paragraphs:
            if para.text == "":
                p = para._element
                p.getparent().remove(p)

    def _add_page_break(self):
        para = self.doc.add_paragraph()
        run = para.add_run()
        run.add_break(WD_ALIGN_PARAGRAPH.CENTER.__class__.__mro__[0])  # noqa
        # Proper page break via XML
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run._r.append(br)

    @staticmethod
    def _safe_font(font_name: str) -> str:
        """Map PDF font names to safe Word-compatible fonts."""
        name_lower = font_name.lower()
        if "times" in name_lower or "serif" in name_lower:
            return "Times New Roman"
        if "courier" in name_lower or "mono" in name_lower:
            return "Courier New"
        if "arial" in name_lower or "helvetica" in name_lower or "sans" in name_lower:
            return "Arial"
        if "calibri" in name_lower:
            return "Calibri"
        if "georgia" in name_lower:
            return "Georgia"
        return "Arial"  # Default fallback
