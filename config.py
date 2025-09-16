from __future__ import annotations
import json
from typing import Any, Dict


def load_config(path: str) -> Dict[str, Any]:
    """Load configuration from a JSON file.
    Returns an empty dict if loading fails.
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read().strip()
        if not text:
            return {}
        # Try JSON first
        try:
            return json.loads(text)
        except Exception:
            pass
    except Exception:
        return {}
    return {}

