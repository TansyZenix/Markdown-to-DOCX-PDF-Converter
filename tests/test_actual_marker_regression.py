from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.style_manager import StyleManager


ROOT = Path(__file__).resolve().parents[1]


def test_actual_source_sentence_has_no_unrendered_bold_markers():
    source = 'CatBoost（Categorical Boosting）是 Yandex 开源的梯度提升框架。其核心创新是有序提升（Ordered Boosting）——在计算每个样本的残差时，仅使用"在时间或排列顺序上位于该样本之前"的其他样本。这种机制天然地减少了传统梯度提升中的**预测偏移（Prediction Shift）**问题，即训练残差和测试残差分布不一致导致的过拟合。'
    generator = DocxGenerator(StyleManager.load(ROOT / "config" / "style.yaml"), ROOT)
    paragraph = generator.convert(source).paragraphs[0]

    assert "**" not in paragraph.text
    assert any(run.text == "预测偏移（Prediction Shift）" and run.bold for run in paragraph.runs)
