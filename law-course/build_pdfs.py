"""Build class PDFs from the Markdown notes in law-course/notes.

Usage: python3 law-course/build_pdfs.py
Each notes/*.md file becomes pdf/<same name>.pdf.

Supported markup (a small subset of Markdown):
  # Title, ## Heading, ### Sub-heading
  - bullet, "  - " sub-bullet, "1. " numbered line
  > callout box
  | table | rows |  (the |---| separator row is skipped)
  **bold**, *italic*, a line with only "===" forces a page break
"""
import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (KeepTogether, PageBreak, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

ROOT = Path(__file__).resolve().parent
# FreeSans has all four styles and the rupee sign (₹).
FONT_DIR = Path("/usr/share/fonts/truetype/freefont")
pdfmetrics.registerFont(TTFont("DV", FONT_DIR / "FreeSans.ttf"))
pdfmetrics.registerFont(TTFont("DV-B", FONT_DIR / "FreeSansBold.ttf"))
pdfmetrics.registerFont(TTFont("DV-I", FONT_DIR / "FreeSansOblique.ttf"))
pdfmetrics.registerFont(TTFont("DV-BI", FONT_DIR / "FreeSansBoldOblique.ttf"))
pdfmetrics.registerFontFamily("DV", normal="DV", bold="DV-B", italic="DV-I", boldItalic="DV-BI")

NAVY = colors.HexColor("#1F3A5F")
GOLD = colors.HexColor("#B8860B")
LIGHT = colors.HexColor("#EEF3F8")
CALLOUT = colors.HexColor("#FFF7E0")

S = {
    "title": ParagraphStyle("title", fontName="DV-B", fontSize=20, leading=25,
                            textColor=NAVY, alignment=TA_CENTER, spaceAfter=4),
    "subtitle": ParagraphStyle("subtitle", fontName="DV-I", fontSize=10.5, leading=14,
                               textColor=colors.HexColor("#555555"),
                               alignment=TA_CENTER, spaceAfter=10),
    "h2": ParagraphStyle("h2", fontName="DV-B", fontSize=13.5, leading=17, textColor=colors.white,
                         backColor=NAVY, borderPadding=(4, 6, 4, 6), spaceBefore=12, spaceAfter=8, keepWithNext=1),
    "h3": ParagraphStyle("h3", fontName="DV-B", fontSize=11.5, leading=15, textColor=NAVY,
                         spaceBefore=8, spaceAfter=4, keepWithNext=1),
    "body": ParagraphStyle("body", fontName="DV", fontSize=9.6, leading=13.4, spaceAfter=4),
    "bullet": ParagraphStyle("bullet", fontName="DV", fontSize=9.6, leading=13.2,
                             leftIndent=14, bulletIndent=4, spaceAfter=2),
    "bullet2": ParagraphStyle("bullet2", fontName="DV", fontSize=9.3, leading=12.8,
                              leftIndent=28, bulletIndent=18, spaceAfter=1),
    "callout": ParagraphStyle("callout", fontName="DV", fontSize=9.6, leading=13.4,
                              backColor=CALLOUT, borderColor=GOLD, borderWidth=0.8,
                              borderPadding=(5, 7, 5, 7), spaceBefore=4, spaceAfter=8),
    "cell": ParagraphStyle("cell", fontName="DV", fontSize=8.6, leading=11),
    "cellh": ParagraphStyle("cellh", fontName="DV-B", fontSize=8.6, leading=11,
                            textColor=colors.white),
}


def inline(text):
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", text)
    return text


def make_table(rows, width):
    rows = [r for r in rows if not re.fullmatch(r"\|?[\s:|-]+\|?", r)]
    cells = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    ncols = max(len(r) for r in cells)
    data = []
    for i, r in enumerate(cells):
        r += [""] * (ncols - len(r))
        style = S["cellh"] if i == 0 else S["cell"]
        data.append([Paragraph(inline(c), style) for c in r])
    # Give each column a share of the width proportional to its longest text.
    lens = [max(len(cells[i][j]) if j < len(cells[i]) else 0 for i in range(len(cells)))
            for j in range(ncols)]
    lens = [min(max(n, 9), 60) for n in lens]
    col_w = [width * n / sum(lens) for n in lens]
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA9BC")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def parse(md, width):
    story, lines, i = [], md.splitlines(), 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.strip() == "===":
            story.append(PageBreak())
        elif line.startswith("# "):
            story.append(Paragraph(inline(line[2:]), S["title"]))
            if i + 1 < len(lines) and lines[i + 1].startswith("_") and lines[i + 1].endswith("_"):
                i += 1
                story.append(Paragraph(inline(lines[i].strip("_")), S["subtitle"]))
        elif line.startswith("## "):
            story.append(Paragraph(inline(line[3:]), S["h2"]))
        elif line.startswith("### "):
            story.append(KeepTogether([Paragraph(inline(line[4:]), S["h3"])]))
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append(lines[i])
                i += 1
            story.append(make_table(rows, width))
            story.append(Spacer(1, 6))
            continue
        elif line.startswith(">"):
            parts = []
            while i < len(lines) and lines[i].startswith(">"):
                parts.append(inline(lines[i].lstrip("> ").rstrip()))
                i += 1
            story.append(Paragraph("<br/>".join(parts), S["callout"]))
            continue
        elif line.startswith("  - "):
            story.append(Paragraph(inline(line[4:]), S["bullet2"], bulletText="–"))
        elif line.startswith("- "):
            story.append(Paragraph(inline(line[2:]), S["bullet"], bulletText="•"))
        elif re.match(r"^\d+\. ", line):
            num, rest = line.split(" ", 1)
            story.append(Paragraph(inline(rest), S["bullet"], bulletText=num))
        else:
            story.append(Paragraph(inline(line), S["body"]))
        i += 1
    return story


def footer(title):
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(GOLD)
        canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
        canvas.setFont("DV", 7.5)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(18 * mm, 10 * mm, f"Law Refresher Course  |  {title}")
        canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
        canvas.restoreState()
    return draw


def build(md_path):
    md = md_path.read_text(encoding="utf-8")
    title = next((l[2:] for l in md.splitlines() if l.startswith("# ")), md_path.stem)
    out = ROOT / "pdf" / f"{md_path.stem}.pdf"
    doc = SimpleDocTemplate(str(out), pagesize=A4, title=title, author="Law Refresher Course",
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=20 * mm)
    doc.build(parse(md, doc.width), onFirstPage=footer(title), onLaterPages=footer(title))
    return out


if __name__ == "__main__":
    for md_file in sorted((ROOT / "notes").glob("*.md")):
        print("built", build(md_file).relative_to(ROOT.parent))
