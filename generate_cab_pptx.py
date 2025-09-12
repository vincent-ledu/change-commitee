#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate a Change Advisory Board (CAB) support PowerPoint:
- Slide 1: S+1 timeline (boxes positioned from start to end date, colored by Type)
- Slides 2..N: one detail slide per change included in S+1

Colors:
  Urgent -> orange, Normal -> blue, Agile -> green
RFC number is a clickable link: https://outils.change.fr/change=rfc123

USAGE:
  python generate_cab_pptx.py \
    --data cab_changes_5_weeks.csv \
    --template "exemple_comité des changements_S+1.pptx" \
    --out Comite_changements_Splus1.pptx \
    --ref-date 2025-09-09

Dependencies:
  pip install python-pptx pandas python-dateutil openpyxl
  (optional for auto-encoding): pip install charset-normalizer
"""
from __future__ import annotations
import argparse
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta, MO
import pandas as pd
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ---------------------- Helpers ----------------------

def parse_fr_date(s) -> datetime:
    """Parse une date ou date+heure en acceptant plusieurs formats.
    Gère aussi directement les objets datetime/pandas.Timestamp.
    """
    # Déjà un datetime ?
    if isinstance(s, datetime):
        return s
    try:
        import pandas as _pd
        if isinstance(s, _pd.Timestamp):
            return s.to_pydatetime()
    except Exception:
        pass

    s = str(s or "").strip()
    # Essais explicites communs (avec et sans heure)
    fmts = (
        "%d/%m/%y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d/%m/%y %H:%M",
        "%d/%m/%Y %H:%M",
        "%d/%m/%y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%dT%H:%M:%S",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    # Dernier recours: pandas.to_datetime avec dayfirst
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt
    except Exception:
        pass

    raise ValueError(f"Unrecognized date/datetime format: {s!r}")

def week_bounds_splus1(ref_date: datetime) -> tuple[datetime, datetime, datetime]:
    """Return (monday_current, monday_next, sunday_next_end_of_day) for S+1.
    Les bornes incluent tout le dimanche (23:59:59.999999).
    """
    monday_current = (ref_date + relativedelta(weekday=MO(-1))).replace(hour=0, minute=0, second=0, microsecond=0)
    monday_next = monday_current + timedelta(weeks=1)
    # Fin de dimanche = début du lundi suivant - 1 microseconde
    sunday_next = (monday_next + timedelta(days=7)) - timedelta(microseconds=1)
    return monday_current, monday_next, sunday_next

def week_bounds_sminus1(ref_date: datetime) -> tuple[datetime, datetime]:
    """Return (monday_prev, sunday_prev_end_of_day) for S-1.
    Les bornes incluent tout le dimanche (23:59:59.999999).
    """
    monday_current = (ref_date + relativedelta(weekday=MO(-1))).replace(hour=0, minute=0, second=0, microsecond=0)
    monday_prev = monday_current - timedelta(weeks=1)
    sunday_prev = (monday_prev + timedelta(days=7)) - timedelta(microseconds=1)
    return monday_prev, sunday_prev

def hyperlink_for_rfc(rfc: str) -> str:
    return f"https://outils.change.fr/change={str(rfc).lower()}"

COLOR_MAP = {
    "urgent": RGBColor(255, 140, 0),   # orange
    "normal": RGBColor(0, 102, 204),   # blue
    "agile":  RGBColor(0, 153, 0),     # green
}
DEFAULT_COLOR = RGBColor(100, 100, 100)

# Colors for S-1 closure code pie chart
PIE_COLOR_SUCCESS = RGBColor(0, 176, 80)            # green
PIE_COLOR_SUCCESS_DIFFICULT = RGBColor(255, 192, 0) # yellow
PIE_COLOR_PARTIAL = RGBColor(255, 204, 153)         # light orange
PIE_COLOR_FAIL_NO_ROLLBACK = RGBColor(255, 140, 0)  # dark orange
PIE_COLOR_FAIL_ROLLBACK = RGBColor(192, 0, 0)       # red

def _norm_text(s: str) -> str:
    import unicodedata
    s = str(s or "").strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

def classify_closure(code: str) -> str:
    """
    Classify 'Code de fermeture' into one of the buckets:
    - succes
    - succes_difficulte
    - partiel
    - echec_sans_retour
    - echec_avec_retour
    Unknown or empty returns '' (ignored).
    """
    t = _norm_text(code)
    if not t:
        return ''
    # direct keywords
    if 'succes' in t or 'reussi' in t or 'réussi' in code.lower():
        if 'diffic' in t:
            return 'succes_difficulte'
        return 'succes'
    if 'partiel' in t or 'partial' in t:
        return 'partiel'
    if 'echec' in t or 'échec' in code.lower() or 'fail' in t:
        if 'retour' in t or 'rollback' in t:
            if 'sans' in t and ('retour' in t or 'rollback' in t):
                # e.g., "echec sans retour arriere"
                return 'echec_sans_retour'
            return 'echec_avec_retour'
        # no explicit mention of rollback -> assume without rollback
        return 'echec_sans_retour'
    # explicit rollback mention without failure keyword -> consider as with rollback
    if 'rollback' in t or 'retour arriere' in t or 'retour arrière' in code.lower():
        return 'echec_avec_retour'
    return ''

def add_sminus1_pie_slide(prs: Presentation,
                          df: pd.DataFrame,
                          monday_prev: datetime,
                          sunday_prev: datetime,
                          layout_index: int | None = None,
                          closure_col: str = 'Code de fermeture') -> None:
    """
    Add a pie chart slide summarizing S-1 by closure code categories.
    Selection criterion: rows with end date within [monday_prev, sunday_prev].
    """
    # Choose layout
    layout = _choose_detail_layout(prs, layout_index)
    slide = prs.slides.add_slide(layout)

    # Title
    left = Cm(1.0); top = Cm(1.0); width = prs.slide_width - Cm(2.0)
    title_shape = slide.shapes.add_textbox(left, top, width, Cm(1.5))
    title_tf = title_shape.text_frame
    title_tf.clear()
    p = title_tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = f"Changements S-1 par code de fermeture ({monday_prev.strftime('%d/%m/%Y')} → {sunday_prev.strftime('%d/%m/%Y')})"
    run.font.size = Pt(24)
    run.font.bold = True

    # Aggregate counts
    subset = df.copy()
    mask_end_in_week = (subset['end_dt'] >= monday_prev) & (subset['end_dt'] <= sunday_prev)
    subset = subset.loc[mask_end_in_week]

    buckets = {
        'succes': 0,
        'succes_difficulte': 0,
        'partiel': 0,
        'echec_sans_retour': 0,
        'echec_avec_retour': 0,
    }
    if closure_col in subset.columns:
        for val in subset[closure_col].fillna(''):
            key = classify_closure(val)
            if key in buckets:
                buckets[key] += 1

    # Prepare chart data in desired order
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE

    categories = [
        'Succès',
        'Succès avec difficulté',
        'Implémenté partiellement',
        'Échec sans retour arrière',
        'Échec avec retour arrière',
    ]
    keys_order = ['succes', 'succes_difficulte', 'partiel', 'echec_sans_retour', 'echec_avec_retour']
    values = [buckets[k] for k in keys_order]

    chart_data = ChartData()
    chart_data.categories = categories
    chart_data.add_series('Changements', values)

    # Add chart
    chart_left = Cm(3.0)
    chart_top = Cm(3.0)
    chart_width = Cm(20.0)
    chart_height = Cm(12.0)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, chart_left, chart_top, chart_width, chart_height, chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.legend.position = 2  # right
    chart.has_title = False
    chart.plots[0].has_data_labels = True
    dl = chart.plots[0].data_labels
    dl.show_category_name = True
    dl.show_percentage = True

    # Apply colors per slice
    colors = [
        PIE_COLOR_SUCCESS,
        PIE_COLOR_SUCCESS_DIFFICULT,
        PIE_COLOR_PARTIAL,
        PIE_COLOR_FAIL_NO_ROLLBACK,
        PIE_COLOR_FAIL_ROLLBACK,
    ]
    series = chart.series[0]
    for idx, point in enumerate(series.points):
        fill = point.format.fill
        fill.solid()
        fill.fore_color.rgb = colors[idx]

def _first_present_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def add_sminus1_non_success_slide(prs: Presentation,
                                   df: pd.DataFrame,
                                   monday_prev: datetime,
                                   sunday_prev: datetime,
                                   layout_index: int | None = None,
                                   closure_col: str = 'Code de fermeture') -> None:
    """Add a table slide listing S-1 changes that are not 'Succès'."""
    layout = _choose_detail_layout(prs, layout_index)
    slide = prs.slides.add_slide(layout)

    # Title
    left = Cm(1.0); top = Cm(1.0); width = prs.slide_width - Cm(2.0)
    title_shape = slide.shapes.add_textbox(left, top, width, Cm(1.5))
    title_tf = title_shape.text_frame
    title_tf.clear()
    p = title_tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = f"Changements S-1 non 'Succès' ({monday_prev.strftime('%d/%m/%Y')} → {sunday_prev.strftime('%d/%m/%Y')})"
    run.font.size = Pt(24)
    run.font.bold = True

    # Determine columns
    resume_col = _first_present_col(df, ['Résumé', 'Resume'])
    detail_col = _first_present_col(df, ['Détail de clôture', 'Detail de clôture', 'Détail de cloture', 'Detail de cloture'])

    subset = df.copy()
    mask_end_in_week = (subset['end_dt'] >= monday_prev) & (subset['end_dt'] <= sunday_prev)
    subset = subset.loc[mask_end_in_week]

    # Filter: not strictly 'succes'
    rows = []
    for _, r in subset.iterrows():
        code_val = r.get(closure_col, '')
        key = classify_closure(code_val)
        if key == 'succes':
            continue
        rows.append(r)

    if not rows:
        # No data: show a friendly message
        msg_shape = slide.shapes.add_textbox(Cm(1.0), Cm(3.0), width, Cm(3.0))
        msg_tf = msg_shape.text_frame
        msg_tf.text = "Aucun changement non 'Succès' pour S-1."
        return

    # Build table with header + items
    tbl_left = Cm(1.0); tbl_top = Cm(3.0)
    tbl_width = prs.slide_width - Cm(2.0)
    headers = ['Numéro', 'Résumé', 'Code de fermeture', 'Détail de clôture']
    table = slide.shapes.add_table(len(rows) + 1, len(headers), tbl_left, tbl_top, tbl_width, Cm(12)).table

    # Set column widths (approximate for readability)
    col_widths = [Cm(4.0), Cm(11.0), Cm(5.0), tbl_width - Cm(4.0 + 11.0 + 5.0)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = w

    # Header row
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
            cell.text_frame.paragraphs[0].runs[0].font.bold = True

    # Data rows
    r_idx = 1
    for r in rows:
        # Numéro with hyperlink
        cell_num = table.cell(r_idx, 0)
        tf = cell_num.text_frame
        tf.clear()
        p0 = tf.paragraphs[0]
        run0 = p0.add_run()
        rfc = str(r.get('Numéro', '')).strip()
        run0.text = rfc
        if rfc:
            run0.hyperlink.address = hyperlink_for_rfc(rfc)
            run0.font.bold = True

        # Résumé
        cell_res = table.cell(r_idx, 1)
        cell_res.text = str(r.get(resume_col, '')) if resume_col else ''

        # Code de fermeture
        cell_code = table.cell(r_idx, 2)
        cell_code.text = str(r.get(closure_col, ''))

        # Détail de clôture
        cell_det = table.cell(r_idx, 3)
        cell_det.text = str(r.get(detail_col, '')) if detail_col else ''

        r_idx += 1

# ---------------------- Robust dataset loader ----------------------

def try_guess_encoding(path: str) -> str | None:
    try:
        from charset_normalizer import from_path
        res = from_path(path).best()
        if res:
            return res.encoding
    except Exception:
        pass
    return None

def load_dataset(path: str, encoding: str | None = None, sep: str | None = None) -> tuple[pd.DataFrame, dict]:
    """
    Load CSV or Excel with robust fallbacks.
    Returns (DataFrame, meta) where meta contains chosen encoding/sep/engine/reader.
    """
    meta = {"reader": None, "encoding": encoding, "sep": sep, "engine": None}

    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
        meta.update({"reader": "excel"})
        return df, meta

    # CSV path
    encodings_to_try = [encoding] if encoding else []
    guessed = try_guess_encoding(path)
    if guessed and guessed not in encodings_to_try:
        encodings_to_try.append(guessed)
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin1", "iso-8859-1"):
        if enc not in encodings_to_try:
            encodings_to_try.append(enc)

    seps_to_try = [sep] if sep is not None else [None, ";", ",", "\t"]
    last_err = None
    for enc in encodings_to_try:
        for s in seps_to_try:
            try:
                engine = "python" if s is None else "c"  # sep=None inference needs "python"
                df = pd.read_csv(path, dtype=str, encoding=enc, sep=s, engine=engine)
                meta.update({"reader": "csv", "encoding": enc, "sep": s, "engine": engine})
                return df, meta
            except Exception as e:
                last_err = e
                continue
    raise last_err if last_err else RuntimeError("Failed to load dataset")

# ---------------------- Timeline builder ----------------------

def build_timeline_slide(prs: Presentation,
                         slide_index: int,
                         week_df: pd.DataFrame,
                         monday_next: datetime,
                         sunday_next: datetime,
                         margins_cm=(1.0, 1.0, 5.0, 2.0),
                         col_gap_cm=0.15,
                         row_height_cm=1.0,
                         row_gap_cm=0.15) -> None:
    """
    Draw colored rounded rectangles for each change over a 7-day grid (S+1).
    - slide_index: index of the slide to draw on (typically 0 for template's first slide)
    - week_df: rows intersecting S+1, with parsed start_dt / end_dt
    """
    slide = prs.slides[slide_index]
    SLIDE_WIDTH = prs.slide_width
    SLIDE_HEIGHT = prs.slide_height

    left_cm, right_cm, top_cm, bottom_cm = margins_cm
    grid_left = Cm(left_cm)
    grid_top = Cm(top_cm)
    grid_width = SLIDE_WIDTH - Cm(left_cm + right_cm)
    grid_height = SLIDE_HEIGHT - Cm(top_cm + bottom_cm)

    col_count = 7
    col_gap = Cm(col_gap_cm)
    col_width = (grid_width - col_gap * (col_count - 1)) / col_count

    row_height = Cm(row_height_cm)
    row_gap = Cm(row_gap_cm)

    rows = []  # occupancy rows over 7 columns

    def find_row(start_idx, end_idx):
        for r_idx, occ in enumerate(rows):
            if all(not occ[c] for c in range(start_idx, end_idx + 1)):
                for c in range(start_idx, end_idx + 1):
                    occ[c] = True
                return r_idx
        new_occ = [False] * col_count
        for c in range(start_idx, end_idx + 1):
            new_occ[c] = True
        rows.append(new_occ)
        return len(rows) - 1

    def day_index(dt):
        return max(0, min(6, (dt - monday_next).days))

    def clamp_to_week(dt: datetime) -> datetime:
        if dt < monday_next:
            return monday_next
        if dt > sunday_next:
            return sunday_next
        return dt

    def time_fraction_of_day(dt: datetime) -> float:
        # returns fraction in [0,1]
        seconds = dt.hour * 3600 + dt.minute * 60 + dt.second + dt.microsecond / 1_000_000.0
        return max(0.0, min(1.0, seconds / 86400.0))

    week_df = week_df.sort_values(["start_dt", "end_dt", "Numéro"], kind="stable")

    for _, row in week_df.iterrows():
        # Clamp to S+1 window for placement
        start_dt = clamp_to_week(row["start_dt"]) if isinstance(row["start_dt"], datetime) else row["start_dt"]
        end_dt = clamp_to_week(row["end_dt"]) if isinstance(row["end_dt"], datetime) else row["end_dt"]

        start_idx = day_index(start_dt)
        end_idx = day_index(end_dt)
        if end_idx < start_idx:
            end_idx = start_idx

        # Row placement still based on whole-day occupancy
        r_idx = find_row(start_idx, end_idx)

        # Sub-day positioning using start/end hour within day
        start_frac = time_fraction_of_day(start_dt)
        end_frac = time_fraction_of_day(end_dt)

        # Compute left/right edges in slide coordinates
        left = grid_left + start_idx * (col_width + col_gap) + col_width * start_frac
        right = grid_left + end_idx * (col_width + col_gap) + col_width * end_frac
        # Ensure minimum width
        if right <= left:
            right = left + Cm(0.2)
        width = right - left
        top = grid_top + r_idx * (row_height + row_gap)
        height = row_height

        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)

        typ_key = str(row.get("Type", "")).strip().lower()
        fill_color = COLOR_MAP.get(typ_key, DEFAULT_COLOR)
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
        shp.line.fill.background()

        tf = shp.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT

        rfc = str(row["Numéro"]).strip()
        resume = str(row.get("Résumé", "")).strip()

        run1 = p.add_run()
        run1.text = rfc
        run1.hyperlink.address = hyperlink_for_rfc(rfc)
        run1.font.bold = True
        run1.font.size = Pt(12)
        run1.font.color.rgb = RGBColor(255, 255, 255)

        run2 = p.add_run()
        run2.text = f" – {resume}"
        run2.font.size = Pt(11)
        run2.font.color.rgb = RGBColor(255, 255, 255)

    week_label = f"Changements S+1 ({monday_next.strftime('%d/%m/%Y')} → {sunday_next.strftime('%d/%m/%Y')})"
    for shape in slide.shapes:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            txt = (shape.text_frame.text or "").strip().lower()
            if "changement s+1" in txt or "changements s+1" in txt:
                shape.text_frame.text = week_label
                break

# ---------------------- Detail slide builder ----------------------

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

def _choose_detail_layout(prs: Presentation, layout_index: int | None) -> "pptx.slide.SlideLayout":
    total = len(prs.slide_layouts)
    if total == 0:
        raise IndexError("Template has no slide layouts")

    # If user requested an index, use it safely (clamped) and warn if needed.
    if layout_index is not None:
        idx = max(0, min(layout_index, total - 1))
        if idx != layout_index:
            print(f"[WARN] detail layout index {layout_index} out of range; using {idx} instead (0..{total-1}).")
        return prs.slide_layouts[idx]

    # Auto-pick: try to find a 'blank' or 'vide' layout by name
    name_candidates = ("blank", "vide", "title only", "titre uniquement", "titre seul")
    for i, layout in enumerate(prs.slide_layouts):
        name = (layout.name or "").strip().lower()
        if any(k in name for k in name_candidates):
            return layout

    # Otherwise, pick the layout with the fewest placeholders (closest to blank)
    def placeholder_count(layout):
        try:
            return len(layout.placeholders)
        except Exception:
            return 999

    best = min(range(total), key=lambda i: placeholder_count(prs.slide_layouts[i]))
    return prs.slide_layouts[best]


def add_detail_slide(prs: Presentation, row: pd.Series, layout_index: int | None = None) -> None:
    layout_to_use = _choose_detail_layout(prs, layout_index)

    slide = prs.slides.add_slide(layout_to_use)

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

# ---------------------- Main ----------------------

def main():
    ap = argparse.ArgumentParser(description="Generate CAB PowerPoint (S+1 timeline + details).")
    ap.add_argument("--data", required=True, help="Path to CSV/Excel with changes")
    ap.add_argument("--template", required=True, help="Path to PPTX template (timeline on slide 0)")
    ap.add_argument("--out", required=True, help="Output PPTX path")
    ap.add_argument("--ref-date", default=None, help="Reference date (YYYY-MM-DD); default: today")
    ap.add_argument("--detail-layout-index", type=int, default=None, help="Optional slide layout index to use for detail slides")
    ap.add_argument("--sminus1-pie", action="store_true", help="Add an S-1 pie chart slide by closure code")
    ap.add_argument("--sminus1-layout-index", type=int, default=None, help="Optional slide layout index for S-1 pie slide")
    ap.add_argument("--list-layouts", action="store_true", help="List slide layouts in the template and exit")
    ap.add_argument("--encoding", default=None, help="Force CSV encoding (e.g. cp1252, latin1, utf-8-sig)")
    ap.add_argument("--sep", default=None, help="Force CSV separator (e.g. ';' ',' '\\t'). If omitted, auto-try common ones.")
    args = ap.parse_args()

    df, meta = load_dataset(args.data, encoding=args.encoding, sep=args.sep)
    print(f"[INFO] Loaded dataset via {meta.get('reader')} "
          f"(encoding={meta.get('encoding')}, sep={repr(meta.get('sep'))}, engine={meta.get('engine')})")

    required_cols = ["Numéro", "Type", "Etat", "Date de début planifiée", "Date de fin planifiée"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {missing}. Present: {list(df.columns)}")

    df["start_dt"] = df["Date de début planifiée"].apply(parse_fr_date)
    df["end_dt"] = df["Date de fin planifiée"].apply(parse_fr_date)

    if args.ref_date:
        ref_date = datetime.strptime(args.ref_date, "%Y-%m-%d")
    else:
        ref_date = datetime.today()

    _, monday_next, sunday_next = week_bounds_splus1(ref_date)
    monday_prev, sunday_prev = week_bounds_sminus1(ref_date)

    mask = (df["start_dt"] <= sunday_next) & (df["end_dt"] >= monday_next)
    week_df = df.loc[mask].copy()
    print(f"[INFO] Changes in S+1: {len(week_df)}")

    prs = Presentation(args.template)

    if args.list_layouts:
        print("[INFO] Available slide layouts (index: name | placeholders)")
        for i, layout in enumerate(prs.slide_layouts):
            try:
                ph = len(layout.placeholders)
            except Exception:
                ph = "?"
            print(f"  {i}: {(layout.name or '').strip()} | placeholders={ph}")
        return
    build_timeline_slide(prs, slide_index=0, week_df=week_df,
                         monday_next=monday_next, sunday_next=sunday_next)

    # Detail slides first
    week_df = week_df.sort_values(["start_dt", "Numéro"], kind="stable")
    for _, row in week_df.iterrows():
        add_detail_slide(prs, row, layout_index=args.detail_layout_index)

    # Then S-1 statistics and non-success list
    if args.sminus1_pie:
        add_sminus1_pie_slide(prs, df=df, monday_prev=monday_prev, sunday_prev=sunday_prev,
                              layout_index=args.sminus1_layout_index)
        add_sminus1_non_success_slide(prs, df=df, monday_prev=monday_prev, sunday_prev=sunday_prev,
                                      layout_index=args.sminus1_layout_index)

    prs.save(args.out)
    print(f"[OK] Generated: {args.out}")
    print(f"[OK] S+1 week: {monday_next.strftime('%d/%m/%Y')} → {sunday_next.strftime('%d/%m/%Y')}")
    if args.sminus1_pie:
        print(f"[OK] S-1 week: {monday_prev.strftime('%d/%m/%Y')} → {sunday_prev.strftime('%d/%m/%Y')}")

if __name__ == "__main__":
    main()
