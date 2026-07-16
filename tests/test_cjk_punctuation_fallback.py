from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_cjk_punctuation_text_token_renders_unparsed_bold_without_markers():
    source = "本段含**第一处强调**——随后说明**第二处强调**问题。"
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    document = generator.convert(source)

    paragraph = document.paragraphs[0]
    assert paragraph.text == "本段含第一处强调——随后说明第二处强调问题。"
    assert any(run.text == "第二处强调" and run.bold for run in paragraph.runs)
