import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 HRFlowable, Table, TableStyle, KeepTogether)
from reportlab.lib.enums import TA_LEFT, TA_CENTER

TC = {
    "🌌 Midnight Executive": {"p": "#7c3aed", "a": "#a78bfa"},
    "🌊 Ocean Deep":         {"p": "#06b6d4", "a": "#67e8f9"},
    "🌿 Forest & Moss":      {"p": "#16a34a", "a": "#4ade80"},
    "🔥 Coral Energy":       {"p": "#ef4444", "a": "#fb923c"},
    "⬜ Clean Minimal":       {"p": "#4f46e5", "a": "#818cf8"},
    "🌸 Berry & Cream":      {"p": "#a855f7", "a": "#e879f9"},
}

def build_pdf(slides, topic, theme, watermark_text=""):
    tc = TC.get(theme, TC["🌌 Midnight Executive"])
    primary = colors.HexColor(tc["p"])
    accent  = colors.HexColor(tc["a"])
    example_bg = colors.HexColor("#1a1a2e")
    stat_bg    = colors.HexColor("#12122a")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    ss = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=ss["Normal"], **kw)

    doc_title = S("DT", fontSize=26, textColor=primary, spaceAfter=4,
                  fontName="Helvetica-Bold")
    doc_sub   = S("DS", fontSize=11, textColor=colors.HexColor("#888888"),
                  spaceAfter=20, fontName="Helvetica-Oblique")
    s_num     = S("SN", fontSize=8, textColor=colors.HexColor("#aaaaaa"),
                  spaceAfter=1, fontName="Helvetica")
    s_title   = S("ST", fontSize=15, textColor=primary, spaceBefore=14,
                  spaceAfter=3, fontName="Helvetica-Bold")
    s_body    = S("SB", fontSize=10, textColor=colors.HexColor("#333333"),
                  leading=16, spaceAfter=3, fontName="Helvetica", leftIndent=10)
    s_note    = S("SNO", fontSize=9, textColor=colors.HexColor("#555555"),
                  leading=14, spaceBefore=3, fontName="Helvetica-Oblique", leftIndent=12)
    s_example = S("SE", fontSize=10, textColor=colors.HexColor("#1a1a5e"),
                  leading=15, spaceAfter=4, fontName="Helvetica", leftIndent=12,
                  backColor=colors.HexColor("#f0f0ff"), borderPad=6)
    s_stat    = S("SS", fontSize=12, textColor=primary, leading=16, spaceAfter=4,
                  fontName="Helvetica-Bold", leftIndent=8)
    s_quote   = S("SQ", fontSize=11, textColor=colors.HexColor("#444466"),
                  leading=18, spaceAfter=4, fontName="Helvetica-Oblique",
                  leftIndent=16, rightIndent=16)
    s_img     = S("SI", fontSize=8, textColor=colors.HexColor("#9999cc"),
                  spaceAfter=2, fontName="Helvetica-Oblique", leftIndent=8)
    s_wm      = S("SW", fontSize=9, textColor=colors.HexColor("#aaaaaa"),
                  alignment=TA_CENTER)

    story = []
    story.append(Paragraph("📊 Presentation Notes", doc_title))
    story.append(Paragraph(f"Topic: {topic}", doc_sub))
    story.append(HRFlowable(width="100%", thickness=2, color=primary, spaceAfter=14))

    # TOC
    toc = [["#", "Title", "Type", "Layout", "Image"]]
    for i, s in enumerate(slides, 1):
        toc.append([str(i), s.get("title","")[:48],
                    s.get("type","content").title(),
                    s.get("layout","auto"),
                    s.get("image_keyword","")])
    tbl = Table(toc, colWidths=[0.7*cm, 9.5*cm, 2*cm, 2*cm, 3*cm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), primary),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,0), 8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f5f5ff")]),
        ("FONTSIZE",(0,1),(-1,-1), 8),
        ("FONTNAME",(0,1),(-1,-1), "Helvetica"),
        ("GRID",(0,0),(-1,-1), 0.25, colors.HexColor("#dddddd")),
        ("TOPPADDING",(0,0),(-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 18))

    for i, s in enumerate(slides, 1):
        block = []
        block.append(Paragraph(f"SLIDE {i}  ·  {s.get('type','content').upper()}", s_num))
        block.append(Paragraph(s.get("title",""), s_title))
        block.append(HRFlowable(width="100%", thickness=0.4, color=primary, spaceAfter=6))

        for line in s.get("content","").split("\n"):
            line = line.strip().lstrip("•●▪-–— ").strip()
            if line:
                block.append(Paragraph(f"▸  {line}", s_body))

        if s.get("example",""):
            block.append(Spacer(1, 4))
            block.append(Paragraph(f"📌 Example:  {s['example']}", s_example))

        if s.get("stat",""):
            block.append(Paragraph(f"📊 Key Stat:  {s['stat']}", s_stat))

        if s.get("quote",""):
            block.append(Paragraph(f'"{s["quote"]}"', s_quote))

        if s.get("notes",""):
            block.append(Paragraph(f"🎤  {s['notes']}", s_note))

        if s.get("image_keyword",""):
            block.append(Paragraph(f"🖼️  Image: {s['image_keyword']}", s_img))

        block.append(Spacer(1, 10))
        story.append(KeepTogether(block))

    if watermark_text:
        story.append(Spacer(1, 14))
        story.append(HRFlowable(width="100%", thickness=0.4,
                                 color=colors.HexColor("#cccccc")))
        story.append(Paragraph(f"🔐 Classification: {watermark_text.upper()}", s_wm))

    doc.build(story)
    buf.seek(0)
    return buf.read()
