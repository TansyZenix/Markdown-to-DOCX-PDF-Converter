from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_display_formula_is_centered_with_single_line_spacing():
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    paragraph = generator.convert("$$x^2$$").paragraphs[0]

    assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert paragraph.paragraph_format.line_spacing_rule == WD_LINE_SPACING.SINGLE
