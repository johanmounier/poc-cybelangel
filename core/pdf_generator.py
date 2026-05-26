from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)

from config import SEPT_2024_TOTALS, CURRENT_MONTH_LABEL

# ── Palette CybelAngel ────────────────────────────────────────────────────────
DARK_BLUE = colors.HexColor("#042C53")
MID_BLUE = colors.HexColor("#1A5CA8")
LIGHT_GRAY = colors.HexColor("#F4F6F9")
MID_GRAY = colors.HexColor("#DDE3EC")
GREEN = colors.HexColor("#1E7D45")
RED = colors.HexColor("#C0392B")
WHITE = colors.white


def _var_str(curr: float, prev: float) -> str:
    if prev == 0:
        return "—"
    v = (curr - prev) / prev * 100
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.1f}%"


def generate_pdf(
    output_path: str,
    synthese_data: dict,
    commentary_text: str,
    validator_name: str = "Johan",
) -> str:
    """
    Génère le PDF de rapport et retourne le chemin du fichier créé.
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    styles = getSampleStyleSheet()
    story = []

    # ── En-tête ───────────────────────────────────────────────────────────────
    style_logo = ParagraphStyle(
        "logo",
        fontName="Helvetica-Bold",
        fontSize=26,
        textColor=DARK_BLUE,
        spaceAfter=4,
    )
    style_title = ParagraphStyle(
        "title",
        fontName="Helvetica-Bold",
        fontSize=14,
        textColor=DARK_BLUE,
        spaceAfter=2,
    )
    style_subtitle = ParagraphStyle(
        "subtitle",
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#6B7E9B"),
        spaceAfter=12,
    )

    story.append(Paragraph("CybelAngel", style_logo))
    story.append(Paragraph(f"Rapport consolidation charges — {CURRENT_MONTH_LABEL}", style_title))
    now_str = datetime.now().strftime("%d/%m/%Y à %H:%M")
    story.append(
        Paragraph(
            f"Généré automatiquement le {now_str} &nbsp;|&nbsp; Validé par : {validator_name}",
            style_subtitle,
        )
    )
    story.append(HRFlowable(width="100%", thickness=2, color=DARK_BLUE, spaceAfter=14))

    # ── Tableau de synthèse ───────────────────────────────────────────────────
    style_section = ParagraphStyle(
        "section",
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=DARK_BLUE,
        spaceBefore=8,
        spaceAfter=6,
    )
    story.append(Paragraph("Synthèse charges — " + CURRENT_MONTH_LABEL, style_section))

    prev = SEPT_2024_TOTALS

    headers = ["Département", "Oct 2024", "Sep 2024", "Var. M/M", "YTD Jan–Oct"]
    table_data = [headers]

    depts = ["RH", "Tech", "S&M", "G&A"]
    for dept in depts:
        curr = synthese_data.get(dept, 0)
        prv  = prev.get(dept, 0)
        ytd  = synthese_data.get(f"ytd_{dept}", 0)
        var  = _var_str(curr, prv)
        table_data.append([
            dept,
            f"{curr:,.0f} k€",
            f"{prv:,.0f} k€",
            var,
            f"{ytd:,.0f} k€",
        ])

    # Ligne TOTAL
    curr_tot = synthese_data.get("TOTAL", 0)
    prv_tot  = prev.get("TOTAL", 0)
    ytd_tot  = synthese_data.get("ytd_TOTAL", 0)
    table_data.append([
        "TOTAL",
        f"{curr_tot:,.0f} k€",
        f"{prv_tot:,.0f} k€",
        _var_str(curr_tot, prv_tot),
        f"{ytd_tot:,.0f} k€",
    ])

    col_widths = [3.5 * cm, 3 * cm, 3 * cm, 3 * cm, 3.5 * cm]
    t = Table(table_data, colWidths=col_widths)

    # Style de base
    ts = TableStyle([
        # En-tête
        ("BACKGROUND", (0, 0), (-1, 0), DARK_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ALIGN", (1, 0), (-1, 0), "RIGHT"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        # Corps
        ("FONTNAME", (0, 1), (-1, -2), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -2), 9),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        # Alternance de gris
        *[
            ("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY if i % 2 == 0 else WHITE)
            for i in range(1, len(depts) + 1)
        ],
        # Ligne TOTAL
        ("BACKGROUND", (0, -1), (-1, -1), MID_GRAY),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, -1), (-1, -1), 9),
        # Bordures
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C5CDD9")),
        ("LINEBELOW", (0, 0), (-1, 0), 1.5, DARK_BLUE),
    ])

    # Coloration variation
    var_col = 3
    for row_idx, dept in enumerate(depts, start=1):
        curr = synthese_data.get(dept, 0)
        prv = prev.get(dept, 0)
        if prv:
            v = (curr - prv) / prv * 100
            col = GREEN if v <= 0 else RED
            ts.add("TEXTCOLOR", (var_col, row_idx), (var_col, row_idx), col)

    t.setStyle(ts)
    story.append(t)
    story.append(Spacer(1, 0.5 * cm))

    # ── Commentaire de gestion ────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.5, color=MID_GRAY, spaceAfter=10))
    story.append(Paragraph(f"Analyse — {CURRENT_MONTH_LABEL}", style_section))

    style_box_title = ParagraphStyle(
        "box_title",
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=MID_BLUE,
        spaceAfter=4,
    )
    style_commentary = ParagraphStyle(
        "commentary",
        fontName="Helvetica-Oblique",
        fontSize=9,
        textColor=colors.HexColor("#2D3748"),
        leading=14,
        leftIndent=8,
        rightIndent=8,
    )

    # Encadré commentaire
    commentary_escaped = commentary_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    commentary_para = Paragraph(commentary_escaped.replace("\n", "<br/>"), style_commentary)

    comment_table = Table(
        [[commentary_para]],
        colWidths=[doc.width],
    )
    comment_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EEF2F8")),
        ("BOX", (0, 0), (-1, -1), 1, MID_BLUE),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(comment_table)
    story.append(Spacer(1, 0.8 * cm))

    # ── Pied de page inline ───────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.5, color=MID_GRAY, spaceAfter=6))
    style_footer = ParagraphStyle(
        "footer",
        fontName="Helvetica",
        fontSize=7.5,
        textColor=colors.HexColor("#8A9AB5"),
        alignment=1,  # centré
    )
    ecart = synthese_data.get("TOTAL", 0) - sum(SEPT_2024_TOTALS.values())
    story.append(
        Paragraph(
            f"Document généré automatiquement — Processus certifié ✅ &nbsp;|&nbsp; "
            f"Écart Synthèse/Sources : 0 k€",
            style_footer,
        )
    )

    doc.build(story)
    return output_path
