from __future__ import annotations

import tempfile
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING


class FormulaHandler:
    def add(self, paragraph, latex: str) -> None:
        import matplotlib.pyplot as plt

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        with tempfile.TemporaryDirectory(prefix="md2docx-math-") as directory:
            image_path = Path(directory) / "formula.png"
            figure = plt.figure(figsize=(0.01, 0.01))
            figure.text(0, 0, f"${latex.strip()}$", fontsize=13)
            figure.savefig(image_path, dpi=220, transparent=True, bbox_inches="tight", pad_inches=0.04)
            plt.close(figure)
            paragraph.add_run().add_picture(str(image_path))
