from __future__ import annotations

import argparse
import sys
from pathlib import Path

from converter.docx_generator import DocxGenerator
from converter.environment import check_environment
from converter.pdf_exporter import export_pdf
from converter.style_manager import StyleManager


def resolve_output_path(input_path: Path, output: str | None) -> Path:
    return Path(output) if output else input_path.with_suffix(".docx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert Markdown to academically formatted DOCX/PDF.")
    parser.add_argument("--input", required=True, help="Markdown input file")
    parser.add_argument("--output", help="DOCX output file (default: input stem with .docx)")
    parser.add_argument("--pdf", action="store_true", help="Also export PDF via LibreOffice")
    parser.add_argument("--style", default="config/style.yaml", help="YAML style file")
    parser.add_argument("--libreoffice", help="Path to soffice.exe")
    parser.add_argument("--skip-env-check", action="store_true", help="Suppress environment warnings")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    input_path = Path(args.input).resolve()
    if not input_path.is_file():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2
    if not args.skip_env_check:
        for warning in check_environment().warnings:
            print(f"Warning: {warning}", file=sys.stderr)
    output_path = resolve_output_path(input_path, args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    styles = StyleManager.load(Path(args.style).resolve())
    document = DocxGenerator(styles, input_path.parent).convert(input_path.read_text(encoding="utf-8"))
    document.save(output_path)
    print(f"DOCX created: {output_path}")
    if args.pdf:
        print(f"PDF created: {export_pdf(output_path, output_path.parent, args.libreoffice)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
