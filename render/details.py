from __future__ import annotations
import pandas as pd
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN

from layouts import choose_detail_layout
from .utils import hyperlink_for_rfc


DETAIL_FIELDS = [
    ("Description", "Description"),
    ("Justification", "Justification"),
    ("Plan d’implémentation", "Plan d’implémentation"),
    ("Analyse risques & impacts", "Analyse de risques et de l’impact"),
    ("Plan de retour arrière", "Plan de retour en arrière"),
    ("Plan de tests", "Plan de tests"),
    ("Informations complémentaires", "Informations complémentaires"),
]


def add_detail_slide(prs: Presentation, row: pd.Series, layout_index: int | None = None) -> None:
    layout_to_use = choose_detail_layout(prs, layout_index)

    slide = prs.slides.add_slide(layout_to_use)

    # Use title placeholder if present; otherwise fall back to a textbox
    title_shape = getattr(slide.shapes, "title", None)
    if title_shape is None:
        left = Cm(1.0); top = Cm(1.0); width = prs.slide_width - Cm(2.0); height = Cm(1.2)
        title_shape = slide.shapes.add_textbox(left, top, width, height)
    title_tf = title_shape.text_frame
    title_tf.clear()
    p = title_tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT

    rfc = str(row["Numéro"]).strip()
    resume = str(row.get("Résumé", "")).strip()

    run1 = p.add_run()
    run1.text = rfc
    run1.hyperlink.address = hyperlink_for_rfc(rfc)
    run1.font.size = Pt(28)
    run1.font.bold = True

    run2 = p.add_run()
    run2.text = f" — {resume}"
    run2.font.size = Pt(24)

    # Info badges (Type, État, Demandeur, Affecté, Début/Fin planifiées)
    badges_left = Cm(1.0)
    badges_top = Cm(2.4)
    avail_width = prs.slide_width - Cm(2.0)
    cols = 3
    gap = Cm(0.5)
    box_w = (avail_width - gap * (cols - 1)) / cols
    box_h = Cm(1.4)

    def add_badge(x, y, w, h, label: str, value: str) -> None:
        shp = slide.shapes.add_textbox(x, y, w, h)
        tf = shp.text_frame
        tf.clear()
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        r1 = p1.add_run()
        r1.text = label
        r1.font.bold = True
        r1.font.size = Pt(10)
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = value
        r2.font.size = Pt(12)

    # Prepare values (with robust date formatting)
    def _fmt_dt(col_dt: str, col_text: str) -> str:
        try:
            return row[col_dt].strftime("%d/%m/%Y %H:%M")
        except Exception:
            return str(row.get(col_text, "")).strip()

    badges = [
        ("Type", str(row.get("Type", "")).strip()),
        ("État", str(row.get("Etat", "")).strip()),
        ("Demandeur", str(row.get("Demandeur", "")).strip()),
        ("Affecté", str(row.get("Affecté", "")).strip()),
        ("Début planifié", _fmt_dt("start_dt", "Date de début planifiée")),
        ("Fin planifiée", _fmt_dt("end_dt", "Date de fin planifiée")),
    ]

    for idx, (lab, val) in enumerate(badges):
        if not val:
            continue
        row_i = idx // cols
        col_i = idx % cols
        x = badges_left + col_i * (box_w + gap)
        y = badges_top + row_i * (box_h + Cm(0.3))
        add_badge(x, y, box_w, box_h, lab, val)

    # Table of selected long-form fields
    tbl_left = Cm(1.0)
    tbl_top = Cm(4.6)
    tbl_width = prs.slide_width - Cm(2.0)
    rows_count = sum(1 for _, key in DETAIL_FIELDS if str(row.get(key, "")).strip() != "")
    rows_count = max(rows_count, 1)
    table = slide.shapes.add_table(rows_count, 2, tbl_left, tbl_top, tbl_width, Cm(12)).table
    table.columns[0].width = Cm(6.0)
    table.columns[1].width = tbl_width - table.columns[0].width

    r = 0
    for label, key in DETAIL_FIELDS:
        val = str(row.get(key, "")).strip()
        if val == "":
            continue
        cell_lbl = table.cell(r, 0)
        cell_lbl.text = label
        # Bold label and set table font size to 10pt
        for p in cell_lbl.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.bold = True
        cell_val = table.cell(r, 1)
        cell_val.text = val
        for p in cell_val.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(10)
        r += 1
