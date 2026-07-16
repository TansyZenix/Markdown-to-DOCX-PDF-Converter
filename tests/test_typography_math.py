from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_removes_spacing_between_chinese_and_latin_text():
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    document = generator.convert("中文 English 和 ABC。")

    assert document.paragraphs[0].text == "中文English和ABC。"


def test_display_latex_is_rendered_as_formula_image_not_source_text():
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    document = generator.convert("$$\\frac{a}{b}$$")

    assert document.inline_shapes
    assert all("frac" not in paragraph.text for paragraph in document.paragraphs)
