"""Data Repair Center for DataBridge v2.3.

User-approved, rule-based repairs only. This module never performs broad fillna()
or guesses sensitive fields. Every repair is proposed first and applied only when
selected by the user.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st

from .settings import T


FORBIDDEN_KEYWORDS = [
    "تاريخ الزيارة", "زيارة متابعة", "follow-up date", "follow up date", "visit date",
    "نتيجة التحليل", "نتيجة تحليل", "confirm", "تاكيدي", "تأكيدي",
    "الاحالة", "الإحالة", "referral",
    "النوع", "gender", "sex",
    "السن", "العمر", "age",
    "الكود المجمع", "beneficiary code", "code",
]
SERVICE_KEYWORDS = ["واقيات", "واقي", "condom", "مزلقات", "مزلق", "lubricant", "lube", "سرنجات", "سرنجة", "syringe", "needle"]
YES_NO_KEYWORDS = ["زهري", "دعم نفسي", "ميثادون", "syphilis", "psychological", "psycho", "methadone"]
LOCATION_KEYWORDS = ["محافظة", "governorate", "gov", "منطقة", "area", "district"]

YES_VALUES = {"نعم", "ايوه", "أيوه", "ايوا", "اه", "آه", "yes", "y", "1", "true", "صح"}
NO_VALUES = {"لا", "لأ", "لاء", "no", "n", "0", "false", "غلط"}

GOV_NORMALIZATION = {
    "القاهره": "القاهرة", "القاهرة": "القاهرة", "cairo": "القاهرة",
    "الجيزه": "الجيزة", "الجيزة": "الجيزة", "giza": "الجيزة",
    "الاسكندريه": "الإسكندرية", "الإسكندرية": "الإسكندرية", "alex": "الإسكندرية", "alexandria": "الإسكندرية",
    "الدقهليه": "الدقهلية", "الدقهلية": "الدقهلية",
    "الشرقيه": "الشرقية", "الشرقية": "الشرقية",
    "الغربيه": "الغربية", "الغربية": "الغربية",
    "المنوفيه": "المنوفية", "المنوفية": "المنوفية",
    "البحيره": "البحيرة", "البحيرة": "البحيرة",
    "الفيوم": "الفيوم", "بني سويف": "بني سويف", "السويس": "السويس",
}


def _norm(value: Any) -> str:
    s = "" if value is None else str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_key(value: Any) -> str:
    s = _norm(value).lower()
    s = re.sub(r"[إأآا]", "ا", s)
    s = s.replace("ة", "ه").replace("ى", "ي")
    return s


def _is_forbidden_col(col: str) -> bool:
    n = _norm_key(col)
    return any(_norm_key(k) in n for k in FORBIDDEN_KEYWORDS)


def _contains_any(col: str, keywords: List[str]) -> bool:
    n = _norm_key(col)
    return any(_norm_key(k) in n for k in keywords)


def _is_blank(v: Any) -> bool:
    if pd.isna(v):
        return True
    return str(v).strip().lower() in {"", "nan", "none", "null", "-", "--"}


def _date_columns(df: pd.DataFrame) -> List[str]:
    out = []
    for c in df.columns:
        n = _norm_key(c)
        if "تاريخ" in n or "date" in n or "زياره متابعه" in n:
            out.append(c)
    return out


def _service_columns(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if _contains_any(c, SERVICE_KEYWORDS) and not _is_forbidden_col(c)]


def _yes_no_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        if _is_forbidden_col(c):
            continue
        if _contains_any(c, YES_NO_KEYWORDS):
            cols.append(c)
            continue
        # content-based yes/no field detection
        s = df[c].dropna().astype(str).str.strip().head(50)
        if len(s) >= 3:
            known = sum(1 for v in s if _norm_key(v) in {_norm_key(x) for x in YES_VALUES | NO_VALUES})
            if known / max(len(s), 1) >= 0.7:
                cols.append(c)
    return list(dict.fromkeys(cols))


def _location_columns(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if _contains_any(c, LOCATION_KEYWORDS) and not _is_forbidden_col(c)]


@dataclass
class RepairSuggestion:
    id: str
    title: str
    count: int
    columns: List[str] = field(default_factory=list)
    description: str = ""


def build_error_examples(df: pd.DataFrame, max_examples: int = 5) -> Dict[str, pd.DataFrame]:
    examples: Dict[str, pd.DataFrame] = {}

    for c in _date_columns(df):
        s = df[c]
        non_blank = s[~s.apply(_is_blank)]
        parsed = pd.to_datetime(non_blank, errors="coerce", dayfirst=True)
        bad = non_blank[parsed.isna()]
        if len(bad):
            examples[f"Invalid Dates — {c}"] = pd.DataFrame({"Row": bad.index + 2, "Invalid Value": bad.astype(str)}).head(max_examples)

    for c in _service_columns(df):
        s = df[c]
        non_blank = s[~s.apply(_is_blank)]
        nums = pd.to_numeric(non_blank, errors="coerce")
        bad = non_blank[nums.isna() | (nums < 0)]
        if len(bad):
            examples[f"Invalid Quantity — {c}"] = pd.DataFrame({"Row": bad.index + 2, "Invalid Value": bad.astype(str)}).head(max_examples)

    for c in _yes_no_columns(df):
        s = df[c]
        non_blank = s[~s.apply(_is_blank)].astype(str).str.strip()
        allowed = {_norm_key(x) for x in YES_VALUES | NO_VALUES | {"نعم", "لا"}}
        bad = non_blank[~non_blank.map(lambda x: _norm_key(x) in allowed)]
        if len(bad):
            examples[f"Non-standard Yes/No — {c}"] = pd.DataFrame({"Row": bad.index + 2, "Invalid Value": bad.astype(str)}).head(max_examples)

    return examples


def build_repair_suggestions(df: pd.DataFrame) -> List[RepairSuggestion]:
    suggestions: List[RepairSuggestion] = []

    # Convert parseable date values only; no filling and no guessing.
    parseable_count = 0
    parseable_cols: List[str] = []
    for c in _date_columns(df):
        s = df[c]
        non_blank = s[~s.apply(_is_blank)]
        if non_blank.empty:
            continue
        parsed = pd.to_datetime(non_blank, errors="coerce", dayfirst=True)
        # Count text/object date cells that can be converted safely.
        count = int(parsed.notna().sum())
        if count:
            parseable_count += count
            parseable_cols.append(c)
    if parseable_count:
        suggestions.append(RepairSuggestion(
            id="convert_parseable_dates",
            title=f"Convert Parseable Dates ({parseable_count} values)",
            count=parseable_count,
            columns=parseable_cols,
            description="Converts readable dates to a consistent yyyy-mm-dd format. Does not fill or guess missing/invalid dates.",
        ))

    # Fill blanks only in allowed service quantity fields.
    for c in _service_columns(df):
        blanks = int(df[c].apply(_is_blank).sum())
        if blanks:
            safe_name = re.sub(r"\W+", "_", c)[:40]
            suggestions.append(RepairSuggestion(
                id=f"fill_blank_service__{safe_name}",
                title=f"Fill Missing {c} With 0 ({blanks} rows)",
                count=blanks,
                columns=[c],
                description="Allowed only for service quantity fields. Requires your approval.",
            ))

    # Standardize yes/no fields.
    yn_count = 0
    yn_cols: List[str] = []
    for c in _yes_no_columns(df):
        s = df[c]
        mask = s.apply(lambda v: (not _is_blank(v)) and (_norm_key(v) in {_norm_key(x) for x in YES_VALUES | NO_VALUES}) and str(v).strip() not in {"نعم", "لا"})
        count = int(mask.sum())
        if count:
            yn_count += count
            yn_cols.append(c)
    if yn_count:
        suggestions.append(RepairSuggestion(
            id="standardize_yes_no",
            title=f"Standardize Yes/No Values ({yn_count} values)",
            count=yn_count,
            columns=yn_cols,
            description="Converts ايوه/Yes/Y/1 to نعم and No/N/0 to لا.",
        ))

    # Normalize location names.
    loc_count = 0
    loc_cols: List[str] = []
    for c in _location_columns(df):
        count = int(df[c].apply(lambda v: (not _is_blank(v)) and _norm_key(v) in GOV_NORMALIZATION and str(v).strip() != GOV_NORMALIZATION[_norm_key(v)]).sum())
        if count:
            loc_count += count
            loc_cols.append(c)
    if loc_count:
        suggestions.append(RepairSuggestion(
            id="normalize_locations",
            title=f"Normalize Governorate/Area Names ({loc_count} values)",
            count=loc_count,
            columns=loc_cols,
            description="Standardizes common Arabic/English governorate variants like القاهره/Cairo → القاهرة.",
        ))

    return suggestions


def apply_repairs(df: pd.DataFrame, selected_ids: List[str], suggestions: List[RepairSuggestion]) -> Tuple[pd.DataFrame, Dict[str, int]]:
    out = df.copy()
    summary: Dict[str, int] = {}
    by_id = {s.id: s for s in suggestions}

    if "convert_parseable_dates" in selected_ids and "convert_parseable_dates" in by_id:
        fixed = 0
        for c in by_id["convert_parseable_dates"].columns:
            before = out[c].copy()
            parsed = pd.to_datetime(out[c], errors="coerce", dayfirst=True)
            mask = before.apply(lambda v: not _is_blank(v)) & parsed.notna()
            out.loc[mask, c] = parsed.loc[mask].dt.strftime("%Y-%m-%d")
            fixed += int(mask.sum())
        summary["Dates Fixed"] = fixed

    service_fixed = 0
    for sid in selected_ids:
        if sid.startswith("fill_blank_service__") and sid in by_id:
            for c in by_id[sid].columns:
                if _is_forbidden_col(c):
                    continue
                mask = out[c].apply(_is_blank)
                out.loc[mask, c] = 0
                service_fixed += int(mask.sum())
    if service_fixed:
        summary["Service Blanks Filled"] = service_fixed

    if "standardize_yes_no" in selected_ids and "standardize_yes_no" in by_id:
        fixed = 0
        for c in by_id["standardize_yes_no"].columns:
            if _is_forbidden_col(c):
                continue
            def conv(v: Any) -> Any:
                nonlocal fixed
                if _is_blank(v):
                    return v
                key = _norm_key(v)
                if key in {_norm_key(x) for x in YES_VALUES}:
                    if str(v).strip() != "نعم":
                        fixed += 1
                    return "نعم"
                if key in {_norm_key(x) for x in NO_VALUES}:
                    if str(v).strip() != "لا":
                        fixed += 1
                    return "لا"
                return v
            out[c] = out[c].apply(conv)
        summary["Yes/No Normalized"] = fixed

    if "normalize_locations" in selected_ids and "normalize_locations" in by_id:
        fixed = 0
        for c in by_id["normalize_locations"].columns:
            def conv_loc(v: Any) -> Any:
                nonlocal fixed
                if _is_blank(v):
                    return v
                key = _norm_key(v)
                if key in GOV_NORMALIZATION:
                    new = GOV_NORMALIZATION[key]
                    if str(v).strip() != new:
                        fixed += 1
                    return new
                return v
            out[c] = out[c].apply(conv_loc)
        summary["Governorates/Areas Standardized"] = fixed

    return out, summary


def render_data_repair_center(df: pd.DataFrame, lang: str = "ar") -> pd.DataFrame | None:
    t = lambda key: T[lang].get(key, key)

    st.markdown(f"### {t('repair_center_title')}")
    st.caption(t("repair_caption"))

    examples = build_error_examples(df)
    suggestions = build_repair_suggestions(df)

    if examples:
        st.markdown("#### 1) Error Examples Viewer")
        for title, ex_df in examples.items():
            with st.expander(f"🔍 Show Examples — {title}: {len(ex_df)} examples", expanded=False):
                st.dataframe(ex_df, use_container_width=True, hide_index=True)
                rows = ex_df["Row"].astype(int).tolist()
                if rows:
                    chosen = st.selectbox("📄 Show Full Row", rows, key=f"fullrow_{title}")
                    idx = max(int(chosen) - 2, 0)
                    if 0 <= idx < len(df):
                        st.dataframe(df.iloc[[idx]], use_container_width=True)
    else:
        st.success(t("repair_no_examples"))

    st.markdown("#### 2) Smart Repair Suggestions")
    if not suggestions:
        st.info(t("repair_no_suggestions"))
        return None

    selected: List[str] = []
    for s in suggestions:
        checked = st.checkbox(f"☐ {s.title}", key=f"repair_{s.id}")
        st.caption(s.description + (f" | Columns: {', '.join(s.columns[:6])}" if s.columns else ""))
        if checked:
            selected.append(s.id)

    st.warning(t("repair_forbidden_warning"))

    if st.button("🚀 Apply Suggested Fixes", use_container_width=True, disabled=not selected, key="apply_data_repairs"):
        repaired, summary = apply_repairs(df, selected, suggestions)
        st.session_state["last_repair_summary"] = summary
        st.success(t("repair_applied_success"))
        return repaired

    summary = st.session_state.get("last_repair_summary")
    if summary:
        st.markdown("#### 📊 Repair Summary")
        st.dataframe(pd.DataFrame([{"Repair": k, "Count": v} for k, v in summary.items()]), use_container_width=True, hide_index=True)
    return None
