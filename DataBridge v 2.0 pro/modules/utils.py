"""Shared utility helpers for files, HTML, history, and safe filtering."""

from __future__ import annotations

import html
from collections import deque
from typing import Any, List, Optional, Tuple

import pandas as pd

from .settings import MAX_HISTORY, MAX_HISTORY_MEM_MB, MAX_UPLOAD_SIZE_MB

def safe_html(value: Any) -> str:
    return html.escape("" if value is None else str(value))

def validate_uploaded_file(uploaded_file) -> None:
    size_mb = uploaded_file.size / (1024 * 1024)
    if size_mb > MAX_UPLOAD_SIZE_MB:
        raise ValueError(f"File too large ({size_mb:.1f} MB). Max: {MAX_UPLOAD_SIZE_MB} MB.")

def safe_read_excel(uploaded_file) -> pd.DataFrame:
    validate_uploaded_file(uploaded_file)
    return pd.read_excel(uploaded_file)

def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")

def verify_password(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# MEMORY-EFFICIENT HISTORY MANAGER
# ────────────────────────────────────────────────────────────────
class SmartHistoryManager:
    def __init__(self, max_history: int = MAX_HISTORY, max_memory_mb: int = MAX_HISTORY_MEM_MB):
        self.history    = deque(maxlen=max_history)
        self.redo_stack: list = []
        self.max_memory_mb = max_memory_mb

    def _estimate_memory(self, df: pd.DataFrame) -> float:
        return df.memory_usage(deep=True).sum() / (1024 * 1024)

    def push(self, df: pd.DataFrame) -> bool:
        mem = self._estimate_memory(df)
        if mem > self.max_memory_mb:
            self.history.append(df.head(1000).copy())
            return False
        self.history.append(df.copy(deep=True))
        self.redo_stack.clear()
        return True

    def undo(self, current_df: pd.DataFrame) -> Optional[pd.DataFrame]:
        if not self.history:
            return None
        self.redo_stack.append(current_df.copy(deep=True))
        return self.history.pop().copy()


# ────────────────────────────────────────────────────────────────
# SECURE FILTERING
# ────────────────────────────────────────────────────────────────
def secure_multi_condition_filter(df: pd.DataFrame, conditions: List[Tuple[str, str, Any]]) -> pd.DataFrame:
    mask = pd.Series([True] * len(df), index=df.index)
    for col, op, val in conditions:
        if col not in df.columns:
            continue
        s = df[col]
        if op == '==':      mask &= (s == val)
        elif op == '!=':    mask &= (s != val)
        elif op == '>':     mask &= (s > val)
        elif op == '<':     mask &= (s < val)
        elif op == '>=':    mask &= (s >= val)
        elif op == '<=':    mask &= (s <= val)
        elif op == 'contains' and isinstance(val, str):
            mask &= s.astype(str).str.contains(val, case=False, na=False, regex=False)
        elif op == 'in' and isinstance(val, (list, tuple, set)):
            mask &= s.isin(val)
    return df[mask].copy()


