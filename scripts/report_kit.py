"""和文 A4 レポート（reportlab + matplotlib）用の共通部品。

フォント: Yu Gothic（Medium/Bold）。数式は matplotlib の mathtext で画像化して埋め込む。
図は matplotlib で PNG 化し本文幅 172mm で貼る。表は Table（縦に長い表はページをまたいで分割）。
"""
from __future__ import annotations

import io
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, Paragraph, Spacer, Table, TableStyle

FONT_DIR = "C:/Windows/Fonts/"
pdfmetrics.registerFont(TTFont("YG", FONT_DIR + "YuGothM.ttc", subfontIndex=0))
pdfmetrics.registerFont(TTFont("YGB", FONT_DIR + "YuGothB.ttc", subfontIndex=0))
pdfmetrics.registerFontFamily("YG", normal="YG", bold="YGB", italic="YG", boldItalic="YGB")
plt.rcParams["font.family"] = ["Yu Gothic", "Meiryo", "sans-serif"]

B, O, A, Y, MUT, INK, GRID = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#52514e", "#0b0b0b", "#e6e5e0"
BODY = ParagraphStyle("b", fontName="YG", fontSize=10.5, leading=17, spaceAfter=5, wordWrap="CJK", alignment=TA_LEFT)
SMALL = ParagraphStyle("s", parent=BODY, fontSize=9, leading=14, textColor=colors.HexColor(MUT))
CAP = ParagraphStyle("c", parent=BODY, fontSize=8.5, leading=12.5, textColor=colors.HexColor(MUT), spaceAfter=8)
H1 = ParagraphStyle("h1", fontName="YGB", fontSize=15, leading=21, spaceBefore=12, spaceAfter=7, textColor=colors.HexColor(INK), wordWrap="CJK")
H2 = ParagraphStyle("h2", fontName="YGB", fontSize=12.5, leading=18, spaceBefore=9, spaceAfter=5, textColor=colors.HexColor(INK), wordWrap="CJK")
TITLE = ParagraphStyle("t", fontName="YGB", fontSize=19, leading=27, spaceAfter=6, textColor=colors.HexColor(INK), wordWrap="CJK")
SUB = ParagraphStyle("st", fontName="YG", fontSize=12, leading=18, spaceAfter=14, textColor=colors.HexColor(MUT), wordWrap="CJK")
CELL = ParagraphStyle("cell", fontName="YG", fontSize=9, leading=12.5, wordWrap="CJK")
CELLB = ParagraphStyle("cellb", parent=CELL, fontName="YGB")


def Pg(t, style=BODY):
    return Paragraph(t, style)


def sp(h=4):
    return Spacer(1, h * mm)


def fmt(v):
    return f"{v:,.0f}"


def style_ax(a, t, xl, yl):
    a.grid(color=GRID, lw=0.8); a.tick_params(colors=MUT, labelsize=8.5)
    for s in a.spines.values():
        s.set_color("#d0cfca")
    a.set_title(t, fontsize=10.5, color=INK, loc="left"); a.set_xlabel(xl, color=MUT, fontsize=9); a.set_ylabel(yl, color=MUT, fontsize=9)


def fig_to_img(fig, width_mm=172):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white"); plt.close(fig); buf.seek(0)
    w, h = PILImage.open(buf).size; buf.seek(0)
    return Image(buf, width=width_mm * mm, height=width_mm * mm * h / w)


def EQ(latex, size=14, max_w=172):
    """mathtext（LaTeX 風）を PNG にして貼る。\\bigl \\text \\arg は非対応なので使わない。"""
    fig = plt.figure(figsize=(10, 0.6)); fig.text(0.01, 0.5, f"${latex}$", fontsize=size, ha="left", va="center"); plt.axis("off")
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.04); plt.close(fig); buf.seek(0)
    w, h = PILImage.open(buf).size; buf.seek(0); wpt = min(max_w * mm, w * 72 / 300 * 0.9)
    img = Image(buf, width=wpt, height=wpt * h / w); img.hAlign = "LEFT"
    return img


def TBL(rows, widths, head=True):
    data = [[Paragraph(str(c), CELLB if (head and i == 0) else CELL) for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1 if head else 0)
    ts = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d0cfca")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    if head:
        ts.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0efe9")))
    for i in range(1 if head else 0, len(rows)):
        if i % 2 == 0:
            ts.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#fafaf7")))
    t.setStyle(TableStyle(ts))
    return t


def make_footer(text):
    def footer(canvas, doc):
        canvas.saveState(); canvas.setFont("YG", 8); canvas.setFillColor(colors.HexColor(MUT))
        canvas.drawString(16 * mm, 10 * mm, text); canvas.drawRightString(A4[0] - 16 * mm, 10 * mm, f"{doc.page}"); canvas.restoreState()
    return footer
