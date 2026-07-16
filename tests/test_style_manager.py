from pathlib import Path

import yaml


ROOT = Path(__file__).resolve().parents[1]


def test_style_config_declares_academic_defaults():
    config = yaml.safe_load((ROOT / "config" / "style.yaml").read_text(encoding="utf-8"))

    assert config["body"]["font_size_pt"] == 10.5
    assert config["body"]["line_spacing_pt"] == 20
    assert config["fonts"]["east_asia"] == "宋体"
    assert config["fonts"]["latin"] == "Times New Roman"
    sizes = [config["headings"][str(level)]["font_size_pt"] for level in (1, 2, 3)]
    assert sizes[0] > sizes[1] > sizes[2] > config["body"]["font_size_pt"]
