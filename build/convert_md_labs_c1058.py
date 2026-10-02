"""C1058 Markdown LG mirror + labs (.md and ReportLab .pdf): WSQ TGS-2024042308 -> non-WSQ C1058."""
import pathlib, re, shutil
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer

OLD_T, NEW_T = "Mastering Articulate Storyline 360 for e-Learning Content Creation", "Articulate Storyline 360 Masterclass"
OLD_C, NEW_C = "TGS-2024042308", "C1058"
URL = "https://www.tertiarycourses.com.sg/articulate-storyline-360-essential-training.html"
VERSION = "v1.0 · 2 October 2026"
COMMON = [("v10 · 22 August 2026", VERSION),
          ("and formal assessment.", "and graded quiz release."),
          ("Both questions assess the mapped criteria", "Both questions test the taught decisions")]

# ---------------- Learner Guide Markdown mirror ----------------
md = pathlib.Path("source-wsq/LEARNER-GUIDE.md").read_text()
md = re.sub(r"## TSC alignment\n.*?(?=\n## )", "", md, flags=re.S)
md = re.sub(r"\*\*Alignment:\*\* (LO\d)[^\n]*", r"**Learning outcome:** \1", md)
md = md.replace("https://www.tertiarycourses.com.sg/wsq-mastering-articulate-storyline-360-for-e-learning-content-creation.html", URL)
md = "\n".join(l for l in md.split("\n")
               if not l.startswith(("- **Legacy trainer deck", "- **Drive assessment", "- **Approved assessment",
                                    "- **Approved courseware mapping")))
for a, b in COMMON + [(OLD_T, NEW_T), (OLD_C, NEW_C)]:
    md = md.replace(a, b)
pathlib.Path("LEARNER-GUIDE.md").write_text(md)

# ---------------- labs ----------------
if pathlib.Path("labs").exists():
    shutil.rmtree("labs")
shutil.copytree("source-wsq/labs", "labs")
for f in pathlib.Path("labs").rglob("*.md"):
    s = f.read_text()
    for a, b in COMMON + [(OLD_T, NEW_T), (OLD_C, NEW_C)]:
        s = s.replace(a, b)
    f.write_text(s)

# ---------------- lab PDFs: re-render with ReportLab (the originals are ReportLab output) ----------------
FD = "/System/Library/Fonts/Supplemental/"
pdfmetrics.registerFont(TTFont("Arial", FD + "Arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", FD + "Arial Bold.ttf"))
pdfmetrics.registerFontFamily("Arial", normal="Arial", bold="Arial-Bold")
BLUE, INK, GREY = HexColor("#1F6FEB"), HexColor("#1F2937"), HexColor("#6B7280")
H1 = ParagraphStyle("h1", fontName="Arial-Bold", fontSize=20, leading=24, textColor=INK, spaceAfter=6)
H2 = ParagraphStyle("h2", fontName="Arial-Bold", fontSize=13, leading=16, textColor=BLUE, spaceBefore=10, spaceAfter=4)
BODY = ParagraphStyle("b", fontName="Arial", fontSize=8.5, leading=12, textColor=INK, alignment=TA_LEFT, spaceAfter=3)
META = ParagraphStyle("m", parent=BODY, fontSize=8, textColor=INK)


def inline(t):
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", t)
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    return t.replace("`", "")


def chrome(c, doc):
    c.saveState()
    w, h = A4
    c.setFont("Arial-Bold", 7.5)
    c.setFillColor(GREY)
    c.drawString(20 * mm, h - 12 * mm, f"{NEW_C} | Activity Resource")
    c.setStrokeColor(HexColor("#E5E7EB"))
    c.line(20 * mm, h - 14 * mm, w - 20 * mm, h - 14 * mm)
    c.line(20 * mm, 14 * mm, w - 20 * mm, 14 * mm)
    c.setFont("Arial", 7)
    c.drawString(20 * mm, 10 * mm, "Tertiary Infotech Academy Pte Ltd")
    c.drawRightString(w - 20 * mm, 10 * mm, f"Page {doc.page}")
    c.restoreState()


def render(md_path):
    story, meta = [], []
    for line in md_path.read_text().split("\n"):
        s = line.strip()
        if not s:
            continue
        if s.startswith("# "):
            story += [Paragraph(inline(s[2:]).replace(" — ", " - "), H1)]
        elif s.startswith(("**Course", "**Version")):
            meta.append(inline(s))
            continue
        else:
            if meta:
                story += [Paragraph(" ".join(meta), META), HRFlowable(width="100%", color=BLUE, thickness=0.8, spaceAfter=4)]
                meta = []
            if s.startswith("## "):
                story.append(Paragraph(inline(s[3:]), H2))
            elif s.startswith("- [ ] "):
                story.append(Paragraph("[ ]&nbsp;&nbsp;" + inline(s[6:]), BODY))
            elif s.startswith("- "):
                story.append(Paragraph(inline(s[2:]), BODY, bulletText="•"))
            else:
                story.append(Paragraph(inline(s), BODY))
    out = md_path.with_suffix(".pdf")
    SimpleDocTemplate(str(out), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm,
                      bottomMargin=20 * mm, title=md_path.parent.name, author="Tertiary Infotech Academy Pte Ltd"
                      ).build(story, onFirstPage=chrome, onLaterPages=chrome)


n = 0
for f in sorted(pathlib.Path("labs").rglob("*.md")):
    if f.with_suffix(".pdf").exists():
        render(f)
        n += 1
print("rebuilt", n, "lab PDFs")
