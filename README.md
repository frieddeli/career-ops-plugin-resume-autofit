# career-ops-plugin-resume-autofit

High-density DOCX and PDF resume generation engine with auto-fit page budgeting, exact 2-page target constraints, CJK typography support, and built-in visual QA auditing for [career-ops](https://github.com/career-ops-hq/career-ops).

## Features

- **Strict 2-Page Budgeting**: Automatically calculates page vertical density, target fill rate (85%–90%), and trailing margins to prevent unsightly single-line page overflows.
- **ATS-Optimized Formatting**: Generates clean OpenXML DOCX files with standardized 0.5-inch margins (36 pt), primary accent borders (`#1F4E79`), and exact tab stop alignments (`pos="10800"`).
- **Native Hyperlink Injection**: Injects raw OPC XML hyperlinks (`w:hyperlink`) ensuring clickable LinkedIn, GitHub, email, and portfolio links survive PDF conversions.
- **Bilingual & CJK Support**: Built-in support for Latin and East Asian typography (`Noto Sans CJK SC`, `PingFang SC`), solving headless LibreOffice CJK font fallback issues.
- **Headless PDF Conversion**: Bridges macOS Word AppleScript and LibreOffice headless for cross-platform PDF compilation.
- **Automated Visual Audit Gate**: Renders PDF pages to raster bitmaps via `pdftoppm` to programmatically verify page count, bottom gap percentage, link counts, and compliance checks.

## Installation

```bash
git clone https://github.com/frieddeli/career-ops-plugin-resume-autofit.git
cd career-ops-plugin-resume-autofit
pip install -r requirements.txt
```

### System Requirements

- **Python 3.9+**
- **poppler** (for `pdftoppm` and `pdftotext` auditing):
  - macOS: `brew install poppler`
  - Ubuntu/Debian: `sudo apt-get install poppler-utils`
- **PDF Renderer**:
  - macOS: Microsoft Word (native) or LibreOffice (`brew install --cask libreoffice`)
  - Linux: LibreOffice (`sudo apt-get install libreoffice`)

## Quickstart

### 1. Build Example Resume
```bash
python cli.py --build-example
```

### 2. Verify an Existing Resume PDF
```bash
python cli.py --verify ./dist/example_resume.pdf
```

### 3. Programmatic Usage in Python

```python
from engine import ResumeBuilder

builder = ResumeBuilder(lang="en", p2_space_after=4.5, p2_line_spacing=11.5)

# Contact Header
builder.build_header(
    name="Jane Doe",
    mobile="+1 (555) 012-3456",
    email="jane.doe@example.com",
    links=[("LinkedIn", "https://linkedin.com/in/janedoe"), ("GitHub", "https://github.com/janedoe")]
)

# Section
builder.add_header("Experience")
builder.add_entry_title("Acme Corp", "2023 – Present")
builder.add_entry_subtitle("Senior Cloud Engineer")
builder.add_bullet("Scalability", "Architected distributed message queues serving 50k req/s.")

# Render DOCX and PDF with automatic audit
report = builder.finalize("build/resume.docx", ["build/resume.pdf"])
print(report)
```

## License

MIT License. See [LICENSE](LICENSE) for details.
