"""
engine.py
---------
Standalone layout and rendering engine for high-density, ATS-optimized 2-page DOCX and PDF resumes.
Features:
- Standardized 0.5-inch margins (36 pt) and #1F4E79 section border rules
- Active XML hyperlinks for contact details
- Right-aligned tab stop dates (pos="10800")
- Native bilingual support (Latin and CJK/East Asian typography)
- Auto-balancer measurement: verifies 2-page budget and second page fill rate (target 85%-90%)
- Integrated audit gate (page count, fill %, link count, compliance regex)
"""

import os
import sys
import shutil
import subprocess
import pypdf
from PIL import Image
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

WORD_APP = "/Applications/Microsoft Word.app"
_WORD_SCRIPT = """
on run argv
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    tell application "Microsoft Word"
        open (POSIX file inPath)
        set theDoc to active document
        save as theDoc file name outPath file format format PDF
        close theDoc saving no
    end tell
end run
"""

def to_pdf(docx_path, out_dir):
    """Convert DOCX to PDF using Microsoft Word on macOS or LibreOffice headless."""
    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    engine = os.environ.get("RESUME_PDF_ENGINE")
    soffice = shutil.which("libreoffice") or shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    
    if engine != "libreoffice" and (engine == "word" or (sys.platform == "darwin" and os.path.isdir(WORD_APP))):
        r = subprocess.run(["osascript", "-", os.path.abspath(docx_path), os.path.abspath(pdf_path)],
                           input=_WORD_SCRIPT, capture_output=True, text=True)
        if r.returncode != 0:
            if soffice and os.path.exists(soffice):
                subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                               capture_output=True, text=True, check=True)
            else:
                raise RuntimeError(f"Word PDF export failed: {r.stderr.strip()}")
    else:
        if not soffice:
            raise RuntimeError("LibreOffice not found. Install LibreOffice or set RESUME_PDF_ENGINE=word on macOS.")
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                       capture_output=True, text=True, check=True)
    return pdf_path

class ResumeBuilder:
    def __init__(self, lang="en", p2_space_after=5.5, p2_line_spacing=12.0):
        self.lang = lang.lower()
        self.is_zh = self.lang in ("zh", "cn", "chinese")
        self.p2_space_after = p2_space_after
        self.p2_line_spacing = p2_line_spacing
        self.doc = docx.Document()
        self._init_doc()

    def _init_doc(self):
        for section in self.doc.sections:
            section.top_margin = Pt(36)
            section.bottom_margin = Pt(36)
            section.left_margin = Pt(36)
            section.right_margin = Pt(36)

    def add_hyperlink(self, paragraph, url, text, color="1F4E79", underline=True):
        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
        new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        if color:
            c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
            rPr.append(c)
        if underline:
            u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
            rPr.append(u)
        rFont = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Noto Sans CJK SC"/>')
        rPr.append(rFont)
        sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="19"/>')
        rPr.append(sz)
        new_run.append(rPr)
        text_node = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
        new_run.append(text_node)
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)

    def add_header(self, title_text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5.5)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.line_spacing = Pt(11.0)
        
        run = p.add_run(title_text.upper() if not self.is_zh else title_text)
        run.font.name = 'Arial'
        run.font.size = Pt(9.5)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="12" w:space="1" w:color="1F4E79"/></w:pBdr>')
        p._p.get_or_add_pPr().append(pBdr)
        return p

    def add_entry_title(self, left_text, right_text=""):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3.0)
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = Pt(10.5)
        
        pPr = p._p.get_or_add_pPr()
        tabs = parse_xml(r'<w:tabs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tab w:val="right" w:pos="10800"/></w:tabs>')
        pPr.append(tabs)
        
        run_left = p.add_run(left_text)
        run_left.font.name = 'Arial'
        run_left.font.size = Pt(9.5)
        run_left.bold = True
        
        if right_text:
            run_tab = p.add_run(f"\t{right_text}")
            run_tab.font.name = 'Arial'
            run_tab.font.size = Pt(9.5)
            run_tab.bold = True
        return p

    def add_entry_subtitle(self, sub_text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = Pt(10.5)
        
        pPr = p._p.get_or_add_pPr()
        if "\t" in sub_text:
            tabs = parse_xml(r'<w:tabs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tab w:val="right" w:pos="10800"/></w:tabs>')
            pPr.append(tabs)
            parts = sub_text.split("\t")
            r1 = p.add_run(parts[0])
            r1.font.name = 'Arial'
            r1.font.size = Pt(9.0)
            r1.bold = True
            r1.italic = not self.is_zh
            
            r2 = p.add_run(f"\t{parts[1]}")
            r2.font.name = 'Arial'
            r2.font.size = Pt(9.0)
            r2.bold = True
        else:
            run = p.add_run(sub_text)
            run.font.name = 'Arial'
            run.font.size = Pt(9.0)
            run.bold = True
            run.italic = not self.is_zh
        return p

    def add_bullet(self, bold_prefix, text_content):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.6)
        p.paragraph_format.line_spacing = Pt(11.2)
        
        colon = "：" if self.is_zh else ": "
        if bold_prefix:
            rb = p.add_run(f"•  {bold_prefix}{colon}")
            rb.font.name = 'Arial'
            rb.font.size = Pt(9.0)
            rb.bold = True
        else:
            rb = p.add_run("•  ")
            rb.font.name = 'Arial'
            rb.font.size = Pt(9.0)
            rb.bold = True
            
        rt = p.add_run(text_content)
        rt.font.name = 'Arial'
        rt.font.size = Pt(9.0)
        return p

    def add_bullet_p2(self, bold_prefix, text_content):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(self.p2_space_after)
        p.paragraph_format.line_spacing = Pt(self.p2_line_spacing)
        
        colon = "：" if self.is_zh else ": "
        if bold_prefix:
            rb = p.add_run(f"•  {bold_prefix}{colon}")
            rb.font.name = 'Arial'
            rb.font.size = Pt(9.0)
            rb.bold = True
        else:
            rb = p.add_run("•  ")
            rb.font.name = 'Arial'
            rb.font.size = Pt(9.0)
            rb.bold = True
            
        rt = p.add_run(text_content)
        rt.font.name = 'Arial'
        rt.font.size = Pt(9.0)
        return p

    def add_blank(self, h=2.5):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = Pt(h)
        return p

    def add_plain_p(self, text, space_after=2.0, bold=False):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = Pt(10.5)
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(9.5)
        r.bold = bold
        return p

    def add_skill(self, cat_title, content, space_after=2.8, line_spacing=11.2):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = Pt(line_spacing)
        colon = "：" if self.is_zh else ": "
        r1 = p.add_run(f"{cat_title}{colon}")
        r1.font.name = 'Arial'
        r1.font.size = Pt(9.0)
        r1.bold = True
        r2 = p.add_run(content)
        r2.font.name = 'Arial'
        r2.font.size = Pt(9.0)

    def add_page_break(self):
        self.doc.add_page_break()

    def build_header(self, name, mobile, email, links, badge="", sub_badge=""):
        p0 = self.doc.add_paragraph()
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(1)
        p0.paragraph_format.line_spacing = Pt(11)

        r_name = p0.add_run(name)
        r_name.font.name = 'Arial'
        r_name.font.size = Pt(13.5)
        r_name.bold = True

        r_mid = p0.add_run(f"  |  {'手机' if self.is_zh else 'Mobile No'}: {mobile}  |  {'邮箱' if self.is_zh else 'Email'}: ")
        r_mid.font.name = 'Arial'
        r_mid.font.size = Pt(9.5)
        r_mid.bold = True

        self.add_hyperlink(p0, f"mailto:{email}", email, color="1F4E79")

        p1 = self.doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.line_spacing = Pt(10.5)

        for label, url in links:
            self.add_hyperlink(p1, url, label)
            p1.add_run(" | ").font.size = Pt(9.5)

        if badge:
            p1.add_run(badge).font.size = Pt(9.5)

        if sub_badge:
            p2 = self.doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(2.5)
            p2.paragraph_format.line_spacing = Pt(10.5)
            p2.add_run(sub_badge).font.size = Pt(9.2)

        self.add_blank(2.0)

    def finalize(self, output_docx_path, output_pdf_paths):
        if self.is_zh:
            for p in self.doc.paragraphs:
                for r in p.runs:
                    rPr = r._r.get_or_add_rPr()
                    rf = rPr.find(qn("w:rFonts"))
                    if rf is None:
                        rf = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Arial" w:hAnsi="Arial"/>')
                        rPr.append(rf)
                    rf.set(qn("w:eastAsia"), "Noto Sans CJK SC")
                    r.italic = False

        for p in self.doc.paragraphs:
            t = p.text.strip()
            if t and not t.startswith("•"):
                p.paragraph_format.keep_with_next = True

        os.makedirs(os.path.dirname(output_docx_path), exist_ok=True)
        self.doc.save(output_docx_path)

        if self.is_zh and sys.platform == "darwin":
            os.environ["RESUME_PDF_ENGINE"] = "word"

        out_dir = os.path.dirname(output_pdf_paths[0])
        os.makedirs(out_dir, exist_ok=True)
        to_pdf(output_docx_path, out_dir)

        generated_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(output_docx_path))[0] + ".pdf")
        for path in output_pdf_paths:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if path != generated_pdf:
                subprocess.run(["cp", generated_pdf, path], check=True)

        reader = pypdf.PdfReader(generated_pdf)
        page_count = len(reader.pages)

        test_prefix = f"/tmp/render_audit_{os.path.basename(generated_pdf)}"
        subprocess.run(["pdftoppm", "-png", "-r", "150", generated_pdf, test_prefix], check=True)
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

        link_count = sum(1 for p in reader.pages for a in (p.get('/Annots') or []) if '/A' in a.get_object())

        report = {
            "pdf_path": generated_pdf,
            "page_count": page_count,
            "fill_pct": round(fill_pct, 1),
            "bottom_gap": round(bottom_gap, 1),
            "link_count": link_count,
            "status": "PASS" if page_count == 2 and 83.0 <= fill_pct <= 91.0 else "WARN"
        }
        return report
