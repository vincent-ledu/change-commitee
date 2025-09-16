from __future__ import annotations
import pandas as pd
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN

from layouts import choose_detail_layout
from .utils import hyperlink_for_rfc


DETAIL_FIELDS = [
    ("Type", "Type"),
    ("État", "Etat"),
    ("Début planifié", "Date de début planifiée"),
    ("Fin planifiée", "Date de fin planifiée"),
    ("Description", "Description"),
    ("Justification", "Justification"),
    ("Plan d’implémentation", "Plan d’implémentation"),
    ("Analyse risques & impacts", "Analyse de risques et de l’impact"),
    ("Plan de retour arrière", "Plan de retour en arrière"),
    ("Plan de tests", "Plan de tests"),
    # ("Groupe d’affectation", "Groupe d’affectation"),
    # ("Groupe gestionnaire", "Groupe gestionnaire"),
    ("Demandeur", "Demandeur"),
    ("Affecté", "Affecté"),
    # ("Element de configuration", "Element de configuration"),
    ("CAB requis", "CAB requis"),
    # ("Code de fermeture", "Code de fermeture"),
    # ("Balises", "Balises"),
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

    tbl_left = Cm(1.0); tbl_top = Cm(3.0)
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
        if cell_lbl.text_frame.paragraphs and cell_lbl.text_frame.paragraphs[0].runs:
            cell_lbl.text_frame.paragraphs[0].runs[0].font.bold = True
        cell_val = table.cell(r, 1)
        cell_val.text = val
        r += 1
