from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


@dataclass
class StyleManager:
    config: dict

    @classmethod
    def load(cls, path: Path) -> "StyleManager":
        with Path(path).open("r", encoding="utf-8") as handle:
            return cls(yaml.safe_load(handle))

    def apply_run_font(self, run, size_pt: float | None = None, bold: bool | None = None) -> None:
        fonts = self.config["fonts"]
        run.font.name = fonts["latin"]
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:ascii"), fonts["latin"])
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), fonts["latin"])
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), fonts["east_asia"])
        run.font.color.rgb = RGBColor.from_string(fonts["color"])
        if size_pt is not None:
            run.font.size = Pt(size_pt)
        if bold is not None:
            run.bold = bold

    def apply_document_defaults(self, document) -> None:
        page = self.config["page"]
        for section in document.sections:
            section.page_width = Cm(page["width_cm"])
            section.page_height = Cm(page["height_cm"])
            section.top_margin = Cm(page["margin_top_cm"])
            section.bottom_margin = Cm(page["margin_bottom_cm"])
            section.left_margin = Cm(page["margin_left_cm"])
            section.right_margin = Cm(page["margin_right_cm"])

        normal = document.styles["Normal"]
        self._apply_style(normal, self.config["body"])
        for level in (1, 2, 3):
            style = document.styles[f"Heading {level}"]
            self._apply_style(style, self.config["headings"][str(level)])
        if "Code Block" not in document.styles:
            document.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
        code_style = document.styles["Code Block"]
        self._apply_style(code_style, self.config["code"], font_name=self.config["fonts"]["code"])

    def _apply_style(self, style, values: dict, font_name: str | None = None) -> None:
        font_name = font_name or self.config["fonts"]["latin"]
        style.font.name = font_name
        style._element.rPr.rFonts.set(qn("w:ascii"), font_name)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), font_name)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), self.config["fonts"]["east_asia"])
        style.font.size = Pt(values["font_size_pt"])
        style.font.color.rgb = RGBColor.from_string(self.config["fonts"]["color"])
        if "bold" in values:
            style.font.bold = values["bold"]
        paragraph_format = style.paragraph_format
        if "line_spacing_pt" in values:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(values["line_spacing_pt"])
        paragraph_format.space_before = Pt(values.get("space_before_pt", 0))
        paragraph_format.space_after = Pt(values.get("space_after_pt", 0))

    def format_paragraph(self, paragraph, *, line_spacing_pt: float | None = None, before: float | None = None, after: float | None = None) -> None:
        body = self.config["body"]
        fmt = paragraph.paragraph_format
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(line_spacing_pt if line_spacing_pt is not None else body["line_spacing_pt"])
        fmt.space_before = Pt(before if before is not None else body["space_before_pt"])
        fmt.space_after = Pt(after if after is not None else body["space_after_pt"])
