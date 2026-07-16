from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_cjk_adjacent_bold_and_code_do_not_leave_markdown_markers():
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    document = generator.convert("中文**强调内容**以及`代码变量`结束。")

    paragraph = document.paragraphs[0]
    assert paragraph.text == "中文强调内容以及代码变量结束。"
    assert any(run.text == "强调内容" and run.bold for run in paragraph.runs)
    assert any(run.text == "代码变量" and run.font.name == "Consolas" for run in paragraph.runs)
