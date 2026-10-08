"""
cli.py
------
Command-line interface to build, audit, and verify resumes with the ResumeBuilder engine.

Usage:
  python cli.py --build-example
  python cli.py --verify <pdf_path> [--banned-regex <pattern>]
"""

import sys
import os
import argparse
import subprocess
import pypdf
from PIL import Image
from engine import ResumeBuilder
from example_facts import PROFILE_EN

DEFAULT_BANNED_REGEX = r"references available upon request|curriculum vitae|responsible for"

def build_example(output_dir="./dist"):
    os.makedirs(output_dir, exist_ok=True)
    docx_path = os.path.join(output_dir, "example_resume.docx")
    pdf_path = os.path.join(output_dir, "example_resume.pdf")

    builder = ResumeBuilder(lang="en", p2_space_after=4.5, p2_line_spacing=11.5)
    
    p = PROFILE_EN
    builder.build_header(
        name=p["name"],
        mobile=p["mobile"],
        email=p["email"],
        links=p["links"],
        badge=p["badge"],
        sub_badge=p["sub_badge"]
    )
    
    builder.add_header(p["summary_title"])
    builder.add_plain_p(p["summary"])
    
    edu = p["education"]
    builder.add_header(edu["title"])
    builder.add_entry_title(edu["school"], edu["period"])
    builder.add_entry_subtitle(edu["degree"])
    for prefix, detail in edu["details"]:
        builder.add_bullet(prefix, detail)
        
    exp = p["experience"]
    builder.add_header(exp["title"])
    for role in exp["roles"]:
        builder.add_entry_title(role["company"], role["period"])
        builder.add_entry_subtitle(role["role"])
        for prefix, bullet in role["bullets"]:
            builder.add_bullet(prefix, bullet)
            
    skills = p["skills"]
    builder.add_header(skills["title"])
    for cat, items in skills["categories"]:
        builder.add_skill(cat, items)
        
    report = builder.finalize(docx_path, [pdf_path])
    print(f"[SUCCESS] Example resume built: {pdf_path}")
    print(f"Metrics: Pages={report['page_count']}, Fill={report['fill_pct']}%, Links={report['link_count']}")
    return pdf_path

def verify_pdf(pdf_path, banned_regex=DEFAULT_BANNED_REGEX):
    if not os.path.exists(pdf_path):
        print(f"[FAIL] File not found: {pdf_path}")
        return False

    reader = pypdf.PdfReader(pdf_path)
    page_count = len(reader.pages)

    # Active hyperlinks
    link_count = sum(1 for p in reader.pages for a in (p.get('/Annots') or []) if '/A' in a.get_object())

    # Measure fill percentage via rendering
    test_prefix = f"/tmp/verify_{os.path.basename(pdf_path)}"
    subprocess.run(["pdftoppm", "-png", "-r", "150", pdf_path, test_prefix], check=True)
    img_path = f"{test_prefix}-{page_count}.png" if page_count > 1 else f"{test_prefix}-1.png"
    img = Image.open(img_path).convert("L")
    w, h = img.size
    pix = img.load()
    last_y = 0
    first_y = h
    for y in range(h):
        for x in range(w):
            if pix[x, y] < 200:
                if y < first_y: first_y = y
                if y > last_y: last_y = y
    fill_pct = (last_y - first_y) / h * 100
    bottom_gap = (h - last_y) / h * 100

    # Banned words check
    banned_count = 0
    banned_matches = ""
    if banned_regex:
        txt_proc = subprocess.run(["pdftotext", pdf_path, "-"], capture_output=True, text=True)
        grep_proc = subprocess.run(["grep", "-i", "-E", banned_regex], input=txt_proc.stdout, capture_output=True, text=True)
        banned_matches = grep_proc.stdout.strip()
        banned_count = len(banned_matches.splitlines()) if banned_matches else 0

    passed = (page_count == 2) and (83.0 <= fill_pct <= 91.0) and (link_count >= 1) and (banned_count == 0)

    print("=" * 60)
    print(f"RESUME AUDIT REPORT: {os.path.basename(pdf_path)}")
    print("=" * 60)
    print(f"Pages:        {page_count} {'[PASS]' if page_count == 2 else '[FAIL (must be 2)]'}")
    print(f"P2 Fill:      {fill_pct:.1f}% {'[PASS]' if 83.0 <= fill_pct <= 91.0 else '[WARN (target 85-90%)]'}")
    print(f"P2 Gap:       {bottom_gap:.1f}%")
    print(f"Links:        {link_count} {'[PASS]' if link_count >= 1 else '[FAIL (need >= 1)]'}")
    if banned_regex:
        print(f"Banned Words: {banned_count} {'[PASS]' if banned_count == 0 else '[FAIL - violations found!]'}")
        if banned_matches:
            print("  Violations:")
            for line in banned_matches.splitlines():
                print(f"    - {line}")
    print("=" * 60)
    print(f"OVERALL STATUS: {'PASS' if passed else 'ACTION REQUIRED'}")
    print("=" * 60)
    return passed

def main():
    parser = argparse.ArgumentParser(description="Resume Auto-Fit Engine CLI")
    parser.add_argument("--build-example", action="store_true", help="Generate an example resume to test the engine")
    parser.add_argument("--verify", type=str, help="Verify a generated PDF resume against layout & content constraints")
    parser.add_argument("--banned-regex", type=str, default=DEFAULT_BANNED_REGEX, help="Regular expression of prohibited terms")
    args = parser.parse_args()

    if args.build_example:
        build_example()
    elif args.verify:
        success = verify_pdf(args.verify, banned_regex=args.banned_regex)
        sys.exit(0 if success else 1)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
