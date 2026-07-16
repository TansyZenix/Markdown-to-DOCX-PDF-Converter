from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_table_cells_render_bold_and_inline_code_without_markdown_markers():
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    document = generator.convert("| 字段 | 值 |\n| --- | --- |\n| **状态** | `ready` |")

    header, value = document.tables[0].rows[1].cells
    assert header.text == "状态"
    assert any(run.text == "状态" and run.bold for run in header.paragraphs[0].runs)
    assert value.text == "ready"
    assert any(run.text == "ready" and run.font.name == "Consolas" for run in value.paragraphs[0].runs)
