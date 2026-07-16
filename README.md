# Markdown to DOCX/PDF Converter

A Python tool that converts Markdown documents into professionally formatted Word (DOCX) and PDF files, designed for Chinese academic papers and technical documentation.

Unlike string-replacement approaches, this tool uses **AST-level Markdown parsing** via `markdown-it-py` to produce editable Word documents. PDF output is always derived from the final DOCX, ensuring a single source of truth for document layout.

## Features

- **Full Markdown support**: headings, paragraphs, bold/italic, inline code, ordered/unordered lists, blockquotes, fenced code blocks, tables (GFM), and images
- **Chinese & Latin mixed typesetting**: automatically applies correct fonts (e.g., SimSun for CJK, Times New Roman for Latin) within the same paragraph
- **Academic heading numbering**: auto-numbers H2/H3 headings (e.g., 1.1, 1.2) upon conversion
- **Figure & table captions**: auto-numbered "Figure X" and "Table X" captions with three-line table formatting
- **LaTeX formula rendering**: converts `$$...$$` display math and `$...$` inline math to PNG images via Matplotlib
- **Style-driven formatting**: all typography settings (page size, margins, fonts, heading sizes, spacing) are configured in a single YAML file — no code changes needed
- **DOCX → PDF export**: seamless PDF export via LibreOffice (preferred) or Microsoft Word (fallback on Windows)

## Pipeline

```text
Markdown → markdown-it-py AST → python-docx DOCX → LibreOffice/Word → PDF
```

The DOCX is the single layout source — PDF is never re-parsed from the original Markdown, guaranteeing consistent pagination and formatting.

## Requirements

- Python 3.10+
- Conda environment (recommended) or venv
- [LibreOffice](https://www.libreoffice.org/) (for PDF export; optional if only DOCX is needed)
- Windows: Microsoft Word with `pywin32` works as a PDF fallback

## Installation

```bash
# Create and activate a conda environment (recommended)
conda create -n md2docx-env python=3.10
conda activate md2docx-env

# Install dependencies
pip install -r requirements.txt
```

The tool checks the active Conda environment and key dependencies on startup. If the environment name doesn't match expectations, a warning is shown — you can bypass it with `--skip-env-check`.

## Usage

### Basic conversion (Markdown → DOCX)

```bash
python main.py --input demo.md --output demo.docx
```

### Convert and export PDF

```bash
python main.py --input demo.md --output demo.docx --pdf
```

### Specify a custom style file

```bash
python main.py --input demo.md --style config/style.yaml --output demo.docx
```

### Specify LibreOffice path (if not in PATH)

```bash
python main.py --input demo.md --output demo.docx --pdf --libreoffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

### Process a sample document

```bash
python main.py --input project_delivery_document.md --output output/project_delivery_document.docx --pdf
```

### CLI Options

| Argument | Description |
|---|---|
| `--input` | Markdown input file (required) |
| `--output` | DOCX output file (default: input filename with .docx) |
| `--pdf` | Also export PDF via LibreOffice |
| `--style` | YAML style configuration file (default: `config/style.yaml`) |
| `--libreoffice` | Path to `soffice.exe` |
| `--skip-env-check` | Suppress Conda environment warnings |

## Style Configuration

All formatting parameters are in [`config/style.yaml`](config/style.yaml):

- **Page**: A4 (21.0 × 29.7 cm) with configurable margins
- **Fonts**: SimSun (east-asia), Times New Roman (latin), Consolas (code)
- **Body text**: 10.5pt, fixed 20pt line spacing
- **Headings**: Heading 1 (16pt, bold), Heading 2 (14pt, bold), Heading 3 (12pt, bold)
- **Captions**: 9pt, bold for figure/table captions
- **Tables**: Three-line table style (top/bottom 1.5pt, header underline 0.5pt)
- **Images**: Configurable max width ratio, auto-centered
- **Blockquotes**: 0.74 cm left indent

Edit the YAML file and re-run the converter — no Python code changes required.

## Project Structure

```text
md-to-docx-pdf-converter/
├── main.py                          # CLI entry point
├── requirements.txt                 # Python dependencies
├── config/
│   └── style.yaml                   # Typography configuration
├── converter/
│   ├── __init__.py                  # Package init; heading numbering & font splitting
│   ├── docx_generator.py            # DOCX generation orchestration
│   ├── environment.py               # Conda/module environment checks
│   ├── formula_handler.py           # LaTeX → PNG rendering via Matplotlib
│   ├── image_handler.py             # Image insertion with scaling & captions
│   ├── markdown_parser.py           # markdown-it-py AST parser wrapper
│   ├── pdf_exporter.py              # DOCX → PDF via LibreOffice or Word
│   ├── style_manager.py             # YAML style loader & document applier
│   └── table_handler.py             # Word table creation with three-line style
├── examples/
│   └── demo.md                      # Feature demonstration sample
├── demo.md                          # Quick-test markdown file
├── project_delivery_document.md     # Real-world technical delivery document (demo)
└── tests/                           # Unit & integration tests
    ├── test_core.py
    ├── test_style_manager.py
    ├── test_integration.py
    ├── test_formula_layout.py
    ├── test_actual_marker_regression.py
    ├── test_cjk_inline_formatting.py
    ├── test_cjk_punctuation_fallback.py
    ├── test_table_inline_formatting.py
    └── test_typography_math.py
```

## Testing

```bash
conda activate md2docx-env
python -m pytest
```

The test suite covers:
- Style loading and application (A4 page, fonts, headings)
- Environment checks and warnings
- Markdown token structure for all block types
- Heading mapping, inline formatting (bold, code)
- Native Word table creation and figure/table captions
- Formula layout (centering, line spacing)
- PDF command construction and fallback behavior
- CLI argument parsing and path resolution
- CJK punctuation rendering edge cases
- Regression tests for real-world documentation input

## Supported Markdown Syntax

| Element | Syntax | DOCX Output |
|---|---|---|
| Heading 1 | `# Title` | Word Heading 1 |
| Heading 2 | `## Section` | Word Heading 2 (auto-numbered) |
| Heading 3 | `### Subsection` | Word Heading 3 (auto-numbered as 1.1, 1.2...) |
| Bold | `**text**` | Bold run |
| Italic | `*text*` | Italic run |
| Inline code | `` `code` `` | Code run (Consolas) |
| Ordered list | `1. item` | Word numbered list |
| Unordered list | `- item` | Word bullet list |
| Blockquote | `> quote` | Indented paragraph |
| Fenced code | ```` ``` ```` | Code Block style |
| Image | `![alt](path.png)` | Scaled, centered with caption |
| Table | GFM `\|---\|` | Native Word table (three-line style) |
| Display math | `$$...$$` | Rendered as centered PNG image |
| Inline math | `$...$` | Rendered as inline PNG image |

## Key Implementation Details

- **AST-based parsing**: Uses `markdown-it-py` for structured tokenization, not fragile regex or string replacement
- **Monkey-patched font splitting**: `Paragraph.add_run()` is patched to split runs by CJK/Latin character type, applying the correct font to each segment
- **Heading numbering**: Tracks H2/H3 state to auto-generate section numbers (e.g., "1.1", "1.2") — numbers are inserted at the DOCX level, leaving the original Markdown clean
- **Academic heading offset**: Markdown `##` maps to DOCX `Heading 1`, `###` → `Heading 2`, etc., matching common academic paper structures
- **Formula rendering**: LaTeX is converted to transparent PNGs via Matplotlib, then inserted as images with proper alignment
- **PDF via DOCX**: PDF is exported from the rendered DOCX file, not re-processed from Markdown, ensuring pixel-perfect consistency

## License

This project is provided for educational and reference purposes.
