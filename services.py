from __future__ import annotations
from datetime import datetime
import re
import pandas as pd
from pptx import Presentation
from layouts import choose_detail_layout

from data_loader import load_dataset, try_guess_encoding, parse_fr_date
from periods import week_bounds_splus1, week_bounds_sminus1
from render.timeline import build_timeline_slide
from render.details import add_detail_slide


REQUIRED_COLS = [
    "Numéro", "Type", "Etat", "Date de début planifiée", "Date de fin planifiée",
]


def prepare_dataframe(data_path: str, encoding: str | None = None, sep: str | None = None) -> tuple[pd.DataFrame, dict]:
    df, meta = load_dataset(data_path, encoding=encoding, sep=sep)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {missing}. Present: {list(df.columns)}")

    # Parse dates robustly
    df["start_dt"] = df["Date de début planifiée"].apply(parse_fr_date)
    df["end_dt"] = df["Date de fin planifiée"].apply(parse_fr_date)
    return df, meta


def compute_weeks(ref_date: datetime) -> tuple[datetime, datetime, datetime, datetime]:
    _, monday_next, sunday_next = week_bounds_splus1(ref_date)
    monday_prev, sunday_prev = week_bounds_sminus1(ref_date)
    return monday_next, sunday_next, monday_prev, sunday_prev


def filter_week_df(df: pd.DataFrame, monday_next: datetime, sunday_next: datetime) -> pd.DataFrame:
    mask = (df["start_dt"] <= sunday_next) & (df["end_dt"] >= monday_next)
    return df.loc[mask].copy()


def filter_by_tags(df: pd.DataFrame, include_tags: list[str] | None, column: str = "Balises") -> pd.DataFrame:
    """Filter rows to keep only those whose `column` contains any of `include_tags`.
    - Matching is case-insensitive substring (robust to separators).
    - If include_tags is falsy or column is missing, returns df unchanged.
    """
    if not include_tags:
        return df
    if column not in df.columns:
        print(f"[WARN] Column '{column}' not found; --include-tags ignored.")
        return df
    tags = [t.strip() for t in include_tags if t and t.strip()]
    if not tags:
        return df
    pattern = "|".join(re.escape(t) for t in tags)
    mask = df[column].fillna("").astype(str).str.contains(pattern, case=False, regex=True)
    return df.loc[mask].copy()


def build_base_presentation(template_path: str,
                            week_df: pd.DataFrame,
                            monday_next: datetime,
                            sunday_next: datetime,
                            detail_layout_index: int | None = None,
                            splus1_layout_index: int | None = None) -> Presentation:
    prs = Presentation(template_path)
    # Place the S+1 timeline either on the template's first slide (default)
    # or on a newly added slide using the requested layout index.
    if splus1_layout_index is None:
        target_slide_index = 0
    else:
        layout = choose_detail_layout(prs, layout_index=splus1_layout_index)
        prs.slides.add_slide(layout)
        target_slide_index = len(prs.slides) - 1

    title = f"Changements S+1 ({monday_next.strftime('%d/%m/%Y')} → {sunday_next.strftime('%d/%m/%Y')})"
    build_timeline_slide(prs, slide_index=target_slide_index, week_df=week_df,
                         monday_next=monday_next, sunday_next=sunday_next,
                         title_text=title)
    week_df = week_df.sort_values(["start_dt", "Numéro"], kind="stable")
    for _, row in week_df.iterrows():
        add_detail_slide(prs, row, layout_index=detail_layout_index)
    return prs
