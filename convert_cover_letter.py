# -*- coding: utf-8 -*-
import os
from fpdf import FPDF
from docx import Document

DOCX = "/Users/gintaresarmaviciute/Downloads/Builder APPlication/Oliver_Varcoe_Cover_Letter.docx"
OUT  = "/Users/gintaresarmaviciute/Downloads/Builder APPlication/Oliver_Varcoe_Cover_Letter.pdf"

REPLACEMENTS = [
    ("–", "-"), ("—", "-"), ("•", "-"),
    ("'", "'"), ("'", "'"), (""", '"'), (""", '"'),
    ("·", "*"), (" ", " "),
    ("é", "e"), ("è", "e"), ("ê", "e"),
    ("à", "a"), ("â", "a"), ("ä", "a"),
    ("î", "i"), ("ï", "i"),
    ("ô", "o"), ("ö", "o"),
    ("ù", "u"), ("û", "u"), ("ü", "u"),
    ("ç", "c"), ("æ", "ae"), ("ø", "o"),
    ("É", "E"), ("À", "A"), ("Ç", "C"),
    ("Ö", "O"), ("Ü", "U"),
]

def clean(t):
    for k, v in REPLACEMENTS:
        t = t.replace(k, v)
    # catch-all: drop anything still outside latin-1
    return t.encode("latin-1", errors="replace").decode("latin-1")


class CL(FPDF):
    def header(self): pass

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(160, 155, 150)
        self.cell(0, 5, str(self.page_no()), align="C")


def build():
    doc = Document(DOCX)
    pdf = CL("P", "mm", "A4")
    pdf.set_margins(22, 22, 22)
    pdf.set_auto_page_break(True, 16)
    pdf.add_page()

    paras = list(doc.paragraphs)
    n = len(paras)

    for i, p in enumerate(paras):
        text = p.text.strip()
        if not text:
            continue
        sname = p.style.name if p.style else ""

        # Name
        if i == 0:
            pdf.set_font("Helvetica", "B", 13)
            pdf.set_text_color(30, 25, 23)
            pdf.cell(0, 7, clean(text))
            pdf.ln(7)
            continue

        # Contact line
        if i == 1:
            pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(100, 95, 90)
            pdf.cell(0, 5, clean(text))
            pdf.ln(5)
            pdf.ln(2)
            pdf.set_draw_color(200, 195, 190)
            pdf.set_line_width(0.3)
            pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
            pdf.ln(6)
            continue

        # Date
        if i == 2:
            pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(120, 115, 110)
            pdf.cell(0, 5, clean(text))
            pdf.ln(5)
            pdf.ln(4)
            continue

        # Bullet
        if sname == "List Paragraph":
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(50, 45, 40)
            pdf.set_x(pdf.l_margin + 4)
            pdf.cell(4, 5.5, "-")
            pdf.multi_cell(0, 5.5, clean(text))
            pdf.ln(1)
            continue

        # Short sign-off (Best, / Oliver)
        if len(text) < 12 and i > n - 5:
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(50, 45, 40)
            pdf.cell(0, 5.5, clean(text))
            pdf.ln(5.5)
            continue

        # Body paragraph
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(50, 45, 40)
        pdf.multi_cell(0, 5.5, clean(text))
        pdf.ln(3)

    pdf.output(OUT)
    print("Saved: {}  ({:,} bytes)".format(OUT, os.path.getsize(OUT)))

build()
