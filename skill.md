---
name: career-ops-plugin-resume-autofit
description: Generate high-density, ATS-optimized 2-page DOCX and PDF resumes with auto-fit page budgeting and CJK typography.
license: MIT
---

# resume-autofit

High-density DOCX and PDF resume generation engine with strict 2-page target constraints, vertical density balancing, and CJK typography support for CareerOps.

## How to run it

- Generate an example resume:
  ```bash
  python cli.py --build-example
  ```
- Verify an existing resume PDF against layout and fill budget constraints:
  ```bash
  python cli.py --verify <pdf_path>
  ```
- Run export hook:
  ```bash
  node plugins.mjs run resume-autofit
  ```

## What it produces

- Clean Microsoft Word `.docx` documents formatted to exact 0.5-inch margins (36 pt) and single-page or two-page constraints.
- High-fidelity `.pdf` conversions generated via Microsoft Word (macOS) or headless LibreOffice.
- Verification audit reports measuring page count, bottom gap percentage, fill density (target 85%–90%), link count, and forbidden word compliance.
