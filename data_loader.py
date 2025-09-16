from __future__ import annotations
import os
from datetime import datetime
import pandas as pd


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


def parse_fr_date(s) -> datetime:
    """Parse une date ou date+heure en acceptant plusieurs formats.
    Gère aussi directement les objets datetime/pandas.Timestamp.
    """
    if isinstance(s, datetime):
        return s
    try:
        import pandas as _pd
        if isinstance(s, _pd.Timestamp):
            return s.to_pydatetime()
    except Exception:
        pass

    text = str(s or "").strip()
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
            return datetime.strptime(text, fmt)
        except Exception:
            pass

    try:
        dt = pd.to_datetime(text, dayfirst=True, errors="raise")
        return dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt
    except Exception:
        pass

    raise ValueError(f"Unrecognized date/datetime format: {s!r}")

