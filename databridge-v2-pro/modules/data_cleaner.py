"""Smart Data Cleaner and Auto Data Mapper for DataBridge v2.2.

Rule-based, offline mapper for Excel files from different organizations.
It detects the data sheet/header row, normalizes columns, maps aliases/fuzzy
matches to DataBridge canonical fields, reports confidence, and produces a
standard_df that later analytics can use safely.
"""

from __future__ import annotations

import re
import json
import os
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

try:  # Optional; code works without it.
    from rapidfuzz import fuzz  # type: ignore
except Exception:  # pragma: no cover
    fuzz = None


# Canonical columns used by the current DataBridge analysis code.
CANONICAL_ALIASES: Dict[str, List[str]] = {
    "مسلسل": ["مسلسل", "serial", "serial no", "no", "number", "رقم", "رقم مسلسل"],
    "الكود المجمع": ["الكود المجمع", "كود مجمع", "code", "beneficiary code", "unique code", "id", "client id"],
    "تاريخ الزيارة": ["تاريخ الزيارة", "visit date", "date of visit", "تاريخ المقابلة", "تاريخ الوصول", "date", "interview date", "reach date"],
    "اسم الجمعية": ["اسم الجمعية", "الجمعية", "organization", "ngo", "partner", "association"],
    "اسم مندوب الوصول ثلاثي\n(الباحث الميداني)": ["مندوب الوصول", "الباحث الميداني", "outreach worker", "field worker"],
    "محافظة الوصول للمستفيد": ["محافظة الوصول", "المحافظة", "governorate", "gov", "محافظة المستفيد"],
    "منطقة الوصول للمستفيد": ["منطقة الوصول", "المنطقة", "area", "district", "location", "منطقة المستفيد"],
    "مكان المقابلة": ["مكان المقابلة", "interview place", "meeting place", "مكان الوصول"],
    "السن": ["السن", "العمر", "age", "age group", "الفئة العمرية"],
    "النوع": ["النوع", "الجنس", "gender", "sex"],
    "واقيات": ["واقيات", "واقي", "condom", "condoms", "condom qty", "عدد الواقيات"],
    "مزلقات": ["مزلقات", "مزلق", "lubricant", "lubricants", "lube", "عدد المزلقات"],
    "سرنجات": ["سرنجات", "سرنجة", "syringe", "syringes", "needle", "needles", "عدد السرنجات"],
    "زهري": ["زهري", "syphilis", "vdrl"],
    "دعم نفسي": ["دعم نفسي", "psychological", "psychosocial", "mental support", "psycho support"],
    "ميثادون": ["ميثادون", "methadone"],
    "نتيجة التحليل": ["نتيجة التحليل", "hiv result", "test result", "result", "hiv test result", "نتيجة اختبار"],
    "نتيجة التحليل التاكيدي": ["التاكيدي", "تأكيدي", "confirm", "confirmatory", "confirmation result"],
    "الاحالة علي العلاج": ["الاحالة", "الإحالة", "referral", "referred", "treatment referral", "referral to treatment"],
}

# Follow-up canonical patterns; the mapper will preserve numbered follow-up columns.
FOLLOWUP_ALIASES = {
    "زيارة متابعة": ["زيارة متابعة", "متابعة", "follow up date", "followup date", "follow-up date", "fu date"],
    "واقيات متابعة": ["واقيات متابعة", "follow up condoms", "fu condoms", "condoms follow"],
    "مزلقات متابعة": ["مزلقات متابعة", "follow up lube", "fu lube", "lubricants follow"],
    "سرنجات متابعة": ["سرنجات متابعة", "follow up syringes", "fu syringes", "needles follow"],
    "زهري متابعة": ["زهري متابعة", "follow up syphilis", "fu syphilis"],
    "دعم نفسي متابعة": ["دعم نفسي متابعة", "follow up psychological", "fu psycho"],
    "نتيجة تحليل متابعة": ["نتيجة تحليل متابعة", "follow up test result", "fu result"],
    "ميثادون متابعة": ["ميثادون متابعة", "follow up methadone", "fu methadone"],
}

REQUIRED_CORE = ["تاريخ الزيارة", "السن", "النوع"]
IMPORTANT_OPTIONAL = ["الكود المجمع", "نتيجة التحليل", "واقيات", "مزلقات", "سرنجات"]
NUMERIC_KEYWORDS = ["واقيات", "مزلقات", "سرنجات"]
YES_NO_KEYWORDS = ["زهري", "دعم نفسي", "ميثادون"]


@dataclass
class MappingDecision:
    original: str
    canonical: str
    confidence: int
    method: str
    action: str  # accepted | verify | suspicious | unknown | merged | skipped
    note: str = ""


@dataclass
class CleanReport:
    status: str = "ok"
    original_rows: int = 0
    original_cols: int = 0
    final_rows: int = 0
    final_cols: int = 0
    memory_mb: float = 0.0
    detected_sheet: Optional[str] = None
    detected_header_row: Optional[int] = None
    sheet_scores: Dict[str, int] = field(default_factory=dict)
    organization_profile: str = "global"
    profile_issues: List[Dict[str, Any]] = field(default_factory=list)
    renamed_columns: Dict[str, str] = field(default_factory=dict)
    mapping_decisions: List[MappingDecision] = field(default_factory=list)
    missing_core_columns: List[str] = field(default_factory=list)
    missing_optional_columns: List[str] = field(default_factory=list)
    invalid_dates: Dict[str, int] = field(default_factory=dict)
    negative_quantities: Dict[str, int] = field(default_factory=dict)
    dropped_empty_rows: int = 0
    standard_rows: int = 0
    standard_cols: int = 0
    notes: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def to_rows(self, lang: str = "ar") -> List[Dict[str, Any]]:
        from .settings import T
        t = lambda key: T[lang].get(key, key)
        c_check, c_result, c_type = t("rpt_col_check"), t("rpt_col_result"), t("rpt_col_type")
        rows: List[Dict[str, Any]] = [
            {c_check: t("rpt_check_status"), c_result: self.status, c_type: "info"},
            {c_check: t("rpt_detected_sheet"), c_result: self.detected_sheet or t("rpt_first_or_selected"), c_type: "info"},
            {c_check: t("rpt_rows_before"), c_result: self.original_rows, c_type: "info"},
            {c_check: t("rpt_cols_before"), c_result: self.original_cols, c_type: "info"},
            {c_check: t("rpt_rows_after"), c_result: self.final_rows, c_type: "info"},
            {c_check: t("rpt_cols_after"), c_result: self.final_cols, c_type: "info"},
            {c_check: t("rpt_memory_usage"), c_result: round(self.memory_mb, 2), c_type: "info"},
        ]
        if self.detected_header_row is not None:
            rows.append({c_check: t("rpt_detected_header_row"), c_result: self.detected_header_row + 1, c_type: "info"})
        if self.sheet_scores:
            scores = "; ".join(f"{k}: {v}" for k, v in sorted(self.sheet_scores.items(), key=lambda x: x[1], reverse=True)[:5])
            rows.append({c_check: t("rpt_sheet_ranking"), c_result: scores, c_type: "info"})
        rows.append({c_check: "Mapping Profile", c_result: self.organization_profile or "global", c_type: "info"})
        accepted = sum(1 for d in self.mapping_decisions if d.action in {"accepted", "merged"})
        verify = sum(1 for d in self.mapping_decisions if d.action in {"verify", "review"})
        suspicious = sum(1 for d in self.mapping_decisions if d.action == "suspicious")
        unknown = sum(1 for d in self.mapping_decisions if d.action == "unknown")
        rows.append({c_check: t("rpt_mapping_result"), c_result: f"Auto Accepted: {accepted} | Verify: {verify} | Suspicious: {suspicious} | Unknown: {unknown}", c_type: "warning" if suspicious or unknown else "ok"})
        if self.missing_core_columns:
            rows.append({c_check: t("rpt_missing_core_cols"), c_result: ", ".join(self.missing_core_columns), c_type: "error"})
        if self.missing_optional_columns:
            rows.append({c_check: t("rpt_missing_optional_cols"), c_result: ", ".join(self.missing_optional_columns), c_type: "warning"})
        for col, count in self.invalid_dates.items():
            if count:
                rows.append({c_check: t("rpt_invalid_dates_in").format(col=col), c_result: count, c_type: "warning"})
        for col, count in self.negative_quantities.items():
            if count:
                rows.append({c_check: t("rpt_negative_qty_in").format(col=col), c_result: count, c_type: "warning"})
        for issue in self.profile_issues:
            rows.append({c_check: "Data Profile Validation", c_result: issue.get("message", ""), c_type: issue.get("type", "warning")})
        if self.dropped_empty_rows:
            rows.append({c_check: t("rpt_empty_rows_dropped"), c_result: self.dropped_empty_rows, c_type: "ok"})
        if self.standard_rows:
            rows.append({c_check: "standard_df", c_result: t("rpt_standard_df_summary").format(rows=self.standard_rows, cols=self.standard_cols), c_type: "ok"})
        for msg in self.notes:
            rows.append({c_check: t("rpt_note"), c_result: t(msg), c_type: "info"})
        for msg in self.warnings:
            rows.append({c_check: t("dq_severity_warning"), c_result: t(msg), c_type: "warning"})
        for msg in self.errors:
            rows.append({c_check: t("dq_severity_error"), c_result: t(msg), c_type: "error"})
        return rows


def _norm_text(value: Any) -> str:
    s = "" if value is None else str(value)
    s = s.replace("\n", " ").replace("\r", " ").strip().lower()
    s = re.sub(r"[إأآا]", "ا", s)
    s = s.replace("ة", "ه").replace("ى", "ي")
    s = s.replace("ؤ", "و").replace("ئ", "ي")
    s = re.sub(r"[^\w\u0600-\u06FF]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _clean_col_name(c: Any, idx: int) -> str:
    s = "" if c is None else str(c).strip()
    s = re.sub(r"\s+", " ", s.replace("\n", " ").replace("\r", " ")).strip()
    if not s or s.lower().startswith("unnamed") or s.lower() == "nan":
        return f"عمود_{idx + 1}"
    return s


def _extract_followup_number(text: str) -> int:
    m = re.search(r"(?:متابعة|متابعه|follow\s*up|fu)\D*(\d)", text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # If there is no number but it clearly is follow-up, assume 1.
    if any(x in _norm_text(text) for x in ["متابعه", "follow up", "followup", "fu"]):
        return 1
    return 0


def _similarity(a: str, b: str) -> int:
    a2, b2 = _norm_text(a), _norm_text(b)
    if not a2 or not b2:
        return 0
    if fuzz is not None:
        return int(fuzz.token_sort_ratio(a2, b2))
    return int(SequenceMatcher(None, a2, b2).ratio() * 100)


def _alias_score(col_name: str, aliases: Iterable[str]) -> Tuple[int, str]:
    n = _norm_text(col_name)
    best_score, method = 0, "unknown"
    for alias in aliases:
        a = _norm_text(alias)
        if not a:
            continue
        score = 0
        this_method = "unknown"
        if n == a:
            score, this_method = 100, "exact"
        elif a in n:
            score, this_method = min(96, 78 + min(len(a), 18)), "contains"
        elif all(part in n for part in a.split() if len(part) > 1):
            score, this_method = 82, "tokens"
        else:
            sim = _similarity(n, a)
            if sim >= 80:
                score, this_method = sim, "fuzzy"
        if score > best_score:
            best_score, method = score, this_method
    return best_score, method


def _method_explanation(method: str) -> str:
    explanations = {
        "exact": "Reason: exact column name match",
        "contains": "Reason: matched synonym dictionary",
        "tokens": "Reason: matched important words from dictionary",
        "fuzzy": "Reason: matched using fuzzy similarity",
        "rule": "Reason: detected by follow-up/service naming rule",
        "memory": "Reason: reused saved mapping memory",
        "profile": "Reason: detected from data values pattern",
        "unknown": "Reason: no safe match found",
    }
    return explanations.get(method, f"Reason: {method}")




def _action_from_confidence(score: int) -> str:
    """Classify mapping confidence for a calmer, more useful review UX."""
    if score >= 95:
        return "accepted"
    if score >= 85:
        return "verify"
    if score >= 70:
        return "suspicious"
    return "unknown"


def _action_label(action: str) -> str:
    labels = {
        "accepted": "🟢 Auto Accepted",
        "merged": "🟢 Auto Accepted",
        "verify": "🟡 Verify",
        "suspicious": "🟠 Suspicious",
        "review": "🟡 Verify",  # backward compatibility for old saved reports
        "unknown": "🔴 Unknown",
        "skipped": "⚪ Skipped",
    }
    return labels.get(action, action)

def _mapping_memory_path() -> str:
    """Save mapping memory in a writable user directory, not beside app.py."""
    try:
        from modules.settings import get_appdata_dir
        return os.path.join(get_appdata_dir(), "mapping_memory.json")
    except Exception:
        base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".databridge")
        path = os.path.join(base, "DataBridge") if os.environ.get("APPDATA") else base
        os.makedirs(path, exist_ok=True)
        return os.path.join(path, "mapping_memory.json")


def _load_mapping_memory() -> Dict[str, Dict[str, str]]:
    path = _mapping_memory_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_mapping_memory(data: Dict[str, Dict[str, str]]) -> None:
    path = _mapping_memory_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _detect_org_profile(df: pd.DataFrame) -> str:
    for col in df.columns:
        if any(x in _norm_text(col) for x in ["اسم الجمعيه", "الجمعيه", "organization", "ngo", "partner", "association"]):
            vals = df[col].dropna().astype(str).str.strip()
            vals = vals[vals != ""]
            if not vals.empty:
                return vals.iloc[0][:80]
    return "global"


def get_mapping_summary(report: Optional[CleanReport]) -> Dict[str, Any]:
    if report is None:
        return {"accepted": 0, "verify": 0, "suspicious": 0, "unknown": 0, "review": 0, "overall_confidence": 0}
    decisions = [d for d in report.mapping_decisions if d.action not in {"skipped"}]
    accepted = sum(1 for d in decisions if d.action in {"accepted", "merged"})
    verify = sum(1 for d in decisions if d.action in {"verify", "review"})
    suspicious = sum(1 for d in decisions if d.action == "suspicious")
    unknown = sum(1 for d in decisions if d.action == "unknown")
    if decisions:
        overall = round(sum(max(0, min(100, d.confidence)) for d in decisions) / len(decisions))
    else:
        overall = 0
    # review is kept for backward compatibility with app code or older reports.
    return {"accepted": accepted, "verify": verify, "suspicious": suspicious, "unknown": unknown, "review": verify + suspicious, "overall_confidence": overall}


def save_mapping_memory_from_report(report: Optional[CleanReport], include_review: bool = True) -> int:
    if report is None:
        return 0
    data = _load_mapping_memory()
    profiles = {"global", report.organization_profile or "global"}
    saved = 0
    for profile in profiles:
        data.setdefault(profile, {})
        for d in report.mapping_decisions:
            if d.action in {"accepted", "merged", "verify"} or (include_review and d.action in {"review", "suspicious"}):
                if d.canonical and d.canonical != "غير معروف" and d.original:
                    data[profile][_norm_text(d.original)] = d.canonical
                    saved += 1
    _save_mapping_memory(data)
    return saved


def _value_ratio(series: pd.Series, predicate) -> float:
    non_empty = series.dropna()
    non_empty = non_empty[non_empty.astype(str).str.strip() != ""]
    if len(non_empty) == 0:
        return 0.0
    return float(non_empty.map(predicate).sum() / len(non_empty))


def _looks_gender(v: Any) -> bool:
    s = _norm_text(v)
    return s in {"ذكر", "انثي", "انثى", "male", "female", "m", "f", "رجل", "سيده", "امراه"}


def _looks_age(v: Any) -> bool:
    s = _norm_text(v)
    if re.search(r"\b(1[5-9]|[2-6][0-9]|70)\b", s):
        return True
    return any(x in s for x in ["15 19", "20 24", "25 29", "30 34", "35 39", "40 44", "45 49", "50 او اكثر", "اقل من 15"])


def _looks_test_result(v: Any) -> bool:
    s = _norm_text(v)
    return any(x in s for x in ["ايجابي", "سلبي", "positive", "negative", "pos", "neg", "رفض"])


def validate_data_profile(df: pd.DataFrame, report: CleanReport) -> None:
    """Validate mapped columns by values, not names only."""
    issues: List[Dict[str, Any]] = []

    def add_issue(canonical: str, message: str) -> None:
        issues.append({"canonical": canonical, "message": message, "type": "warning"})
        for d in report.mapping_decisions:
            if d.canonical == canonical and d.action in {"accepted", "merged", "verify"}:
                d.action = "suspicious"
                d.confidence = min(d.confidence, 79)
                extra = "Mapping Suspicious: values do not fit expected profile"
                d.note = f"{d.note} | {extra}" if d.note else extra

    if "السن" in df.columns:
        age_ratio = _value_ratio(df["السن"], _looks_age)
        gender_ratio = _value_ratio(df["السن"], _looks_gender)
        if gender_ratio >= 0.55 and age_ratio < 0.35:
            add_issue("السن", "⚠️ Mapping Suspicious: Column mapped as Age, but values look like Gender")

    if "النوع" in df.columns:
        gender_ratio = _value_ratio(df["النوع"], _looks_gender)
        age_ratio = _value_ratio(df["النوع"], _looks_age)
        if age_ratio >= 0.55 and gender_ratio < 0.35:
            add_issue("النوع", "⚠️ Mapping Suspicious: Column mapped as Gender, but values look like Age")
        elif gender_ratio < 0.35:
            sample = df["النوع"].dropna().astype(str).head(3).tolist()
            if sample:
                add_issue("النوع", f"⚠️ Mapping Suspicious: Gender column has unusual values: {sample}")

    for col in ["تاريخ الزيارة"] + [f"زيارة متابعة {i}" for i in range(1, 6)]:
        if col in df.columns:
            parsed, invalid = _parse_maybe_date(df[col])
            non_empty = int(df[col].notna().sum())
            if non_empty >= 5 and invalid / max(non_empty, 1) > 0.5:
                add_issue(col, f"⚠️ Mapping Suspicious: {col} was mapped as Date, but most values are not dates")

    for col in ["نتيجة التحليل"] + [f"نتيجة تحليل متابعة {i}" for i in range(1, 6)]:
        if col in df.columns:
            ratio = _value_ratio(df[col], _looks_test_result)
            non_empty = int(df[col].notna().sum())
            if non_empty >= 5 and ratio < 0.35:
                add_issue(col, f"⚠️ Mapping Suspicious: {col} was mapped as test result, but values do not look like Positive/Negative")

    for col in ["واقيات", "مزلقات", "سرنجات"] + [f"واقيات متابعة {i}" for i in range(1,6)] + [f"مزلقات متابعة {i}" for i in range(1,6)] + [f"سرنجات متابعة {i}" for i in range(1,6)]:
        if col in df.columns:
            numeric = pd.to_numeric(df[col], errors="coerce")
            non_empty = int(df[col].notna().sum())
            if non_empty >= 5 and numeric.notna().sum() / max(non_empty, 1) < 0.5:
                add_issue(col, f"⚠️ Mapping Suspicious: {col} was mapped as quantity, but many values are not numeric")

    report.profile_issues.extend(issues)


def _followup_mapping(col_name: str) -> Optional[Tuple[str, int, str]]:
    fu_no = _extract_followup_number(col_name)
    norm = _norm_text(col_name)
    if any(x in norm for x in ["ملاحظات", "ملاحظه", "notes", "note", "comments"]):
        return None
    if not fu_no:
        return None
    best: Tuple[str, int, str] = ("", 0, "unknown")
    for canonical_prefix, aliases in FOLLOWUP_ALIASES.items():
        score, method = _alias_score(col_name, aliases)
        # Boost service words inside a follow-up column.
        if "واقي" in norm or "condom" in norm:
            if canonical_prefix == "واقيات متابعة": score = max(score, 94); method = "rule"
        if "مزلق" in norm or "lube" in norm or "lubric" in norm:
            if canonical_prefix == "مزلقات متابعة": score = max(score, 94); method = "rule"
        if "سرنج" in norm or "syringe" in norm or "needle" in norm:
            if canonical_prefix == "سرنجات متابعة": score = max(score, 94); method = "rule"
        if "نتيجه" in norm or "result" in norm:
            if canonical_prefix == "نتيجة تحليل متابعة": score = max(score, 92); method = "rule"
        if "زياره" in norm or "date" in norm:
            if canonical_prefix == "زيارة متابعة": score = max(score, 92); method = "rule"
        if score > best[1]:
            best = (canonical_prefix, score, method)
    if best[1] >= 70:
        return (f"{best[0]} {fu_no}", best[1], best[2])
    return None


def detect_column_mapping(columns: Iterable[Any], memory: Optional[Dict[str, str]] = None) -> Tuple[Dict[str, str], List[MappingDecision]]:
    """Return original_col -> canonical_col and confidence decisions.

    Rules:
    - >=95 Auto Accepted
    - 85-94 Verify
    - 70-84 Suspicious
    - <70 Unknown
    """
    mapping: Dict[str, str] = {}
    decisions: List[MappingDecision] = []
    used_canonicals: set[str] = set()
    candidates: List[Tuple[int, str, str, str]] = []
    memory = memory or {}

    for col in columns:
        col_str = str(col)
        mem_target = memory.get(_norm_text(col_str))
        if mem_target:
            candidates.append((99, col_str, mem_target, "memory"))
            continue

    for col in columns:
        col_str = str(col)
        if col_str.startswith("عمود_"):
            continue
        fu_match = _followup_mapping(col_str)
        if fu_match:
            canonical, score, method = fu_match
            candidates.append((score, col_str, canonical, method))
            continue
        for canonical, aliases in CANONICAL_ALIASES.items():
            score, method = _alias_score(col_str, aliases)
            norm_col = _norm_text(col_str)
            # Avoid mapping follow-up columns to base fields.
            if _extract_followup_number(col_str) and canonical in ["واقيات", "مزلقات", "سرنجات", "زهري", "دعم نفسي", "ميثادون", "نتيجة التحليل"]:
                score = min(score, 45)
            # Avoid mapping confirmatory/rapid-specific columns to the generic base result when a better specific field exists.
            if canonical == "نتيجة التحليل" and any(x in norm_col for x in ["تاكيدي", "تأكيدي", "confirm"]):
                score = min(score, 45)
            if score >= 70:
                candidates.append((score, col_str, canonical, method))

    for score, original, canonical, method in sorted(candidates, key=lambda x: x[0], reverse=True):
        if original in mapping:
            continue
        if canonical in used_canonicals and original != canonical:
            decisions.append(MappingDecision(original, canonical, score, method, "skipped", "العمود القياسي مستخدم بالفعل"))
            continue
        action = _action_from_confidence(score)
        if original.strip() != canonical:
            mapping[original] = canonical
        used_canonicals.add(canonical)
        decisions.append(MappingDecision(original, canonical, score, method, action, _method_explanation(method)))

    # Unknown useful-looking columns for review screen.
    mapped_originals = {d.original for d in decisions}
    for col in columns:
        col_str = str(col)
        if col_str not in mapped_originals and not col_str.startswith("عمود_"):
            norm = _norm_text(col_str)
            if any(x in norm for x in ["date", "تاريخ", "age", "سن", "gender", "نوع", "result", "نتيجه", "referral", "احاله", "واقي", "مزلق", "سرنج", "follow", "متابعه"]):
                decisions.append(MappingDecision(col_str, "غير معروف", 0, "unknown", "unknown", _method_explanation("unknown") + " — يحتاج ربط يدوي لاحقًا"))
    return mapping, decisions


def _header_row_score(row: pd.Series) -> int:
    values = [v for v in row.tolist() if pd.notna(v) and str(v).strip()]
    if not values:
        return 0
    text_hits = 0
    alias_hits = 0
    for v in values:
        s = str(v)
        if not re.fullmatch(r"\d+(\.0)?", s.strip()):
            text_hits += 1
        for aliases in CANONICAL_ALIASES.values():
            score, _ = _alias_score(s, aliases)
            if score >= 80:
                alias_hits += 1
                break
        if _followup_mapping(s):
            alias_hits += 1
    return text_hits + alias_hits * 5


def _score_sheet(preview: pd.DataFrame) -> int:
    if preview.empty:
        return 0
    row_scores = preview.apply(_header_row_score, axis=1).tolist()
    best_header = max(row_scores) if row_scores else 0
    non_empty_rows = int(preview.dropna(how="all").shape[0])
    non_empty_cols = int(preview.dropna(axis=1, how="all").shape[1])
    return int(best_header * 3 + min(non_empty_rows, 50) + min(non_empty_cols, 80))


def _detect_sheet(uploaded_file, scan_rows: int = 20) -> Tuple[str | int, Dict[str, int]]:
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    previews = pd.read_excel(uploaded_file, sheet_name=None, header=None, nrows=scan_rows)
    scores = {str(name): _score_sheet(df) for name, df in previews.items()}
    if not scores:
        return 0, {}
    best_name = max(scores, key=scores.get)
    return best_name, scores


def smart_read_excel(uploaded_file, sheet_name: int | str | None = None, scan_rows: int = 20) -> Tuple[pd.DataFrame, CleanReport]:
    """Read Excel with sheet/header detection, then clean and map it."""
    detected_sheet: int | str
    sheet_scores: Dict[str, int] = {}
    if sheet_name is None:
        detected_sheet, sheet_scores = _detect_sheet(uploaded_file, scan_rows=scan_rows)
    else:
        detected_sheet = sheet_name

    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    preview = pd.read_excel(uploaded_file, sheet_name=detected_sheet, header=None, nrows=scan_rows)
    row_scores = preview.apply(_header_row_score, axis=1).tolist() if len(preview) else [0]
    best_row = int(max(range(len(row_scores)), key=lambda i: row_scores[i])) if row_scores else 0
    best_score = row_scores[best_row] if row_scores else 0
    header_row = best_row if best_score >= 7 else 0

    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    df = pd.read_excel(uploaded_file, sheet_name=detected_sheet, header=header_row)
    cleaned, report = clean_dataframe(df)
    report.detected_sheet = str(detected_sheet)
    report.sheet_scores = sheet_scores
    report.detected_header_row = header_row
    std = generate_standard_df(cleaned)
    report.standard_rows = len(std)
    report.standard_cols = len(std.columns)
    return cleaned, report


def _parse_maybe_date(series: pd.Series) -> Tuple[pd.Series, int]:
    """Parse dates without corrupting already-normalized YYYY-MM-DD values.

    Previous versions used ``pd.to_datetime(..., dayfirst=True)`` directly.
    That makes pandas interpret strings like ``2026-01-05`` as 2026-05-01,
    which created fake months and reduced May from 213 records to 25.
    This parser preserves real datetime values, parses ISO YYYY-MM-DD as
    year-month-day, then only uses day-first parsing for non-ISO text dates.
    """
    original_non_empty = series.notna() & (series.astype(str).str.strip() != "")

    # Preserve datetime columns/cells exactly.
    if pd.api.types.is_datetime64_any_dtype(series):
        parsed = pd.to_datetime(series, errors="coerce")
    else:
        text = series.astype("string").str.strip()
        parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

        # Excel serial numbers.
        numeric = pd.to_numeric(series, errors="coerce")
        serial_mask = numeric.between(20000, 60000)
        if serial_mask.any():
            parsed.loc[serial_mask] = pd.to_datetime(
                numeric.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
            )

        # ISO dates generated by the source converter: YYYY-MM-DD.
        iso_mask = parsed.isna() & text.str.match(r"^\d{4}-\d{1,2}-\d{1,2}$", na=False)
        if iso_mask.any():
            parsed.loc[iso_mask] = pd.to_datetime(text.loc[iso_mask], errors="coerce", format="%Y-%m-%d")

        # Common slash/dash dates written as DD/MM/YYYY or DD-MM-YYYY.
        dmy_mask = parsed.isna() & text.str.match(r"^\d{1,2}[/-]\d{1,2}[/-]\d{4}$", na=False)
        if dmy_mask.any():
            parsed.loc[dmy_mask] = pd.to_datetime(text.loc[dmy_mask], errors="coerce", dayfirst=True)

        # Fallback for any other parseable date text.
        rest_mask = parsed.isna() & original_non_empty
        if rest_mask.any():
            parsed.loc[rest_mask] = pd.to_datetime(text.loc[rest_mask], errors="coerce", dayfirst=True)

    invalid = int((original_non_empty & parsed.isna()).sum())
    return parsed, invalid


def _clean_yes_no_value(v: Any) -> Any:
    if pd.isna(v):
        return None
    s = _norm_text(v)
    if s in {"نعم", "اه", "ايوه", "yes", "y", "true", "1", "done", "تم"}:
        return "نعم"
    if s in {"لا", "لاء", "no", "n", "false", "0", "none", "not done"}:
        return "لا"
    return v


def _clean_test_result(v: Any) -> Any:
    if pd.isna(v):
        return None
    s = _norm_text(v)
    if any(x in s for x in ["ايجابي", "positive", "pos", "+"]):
        return "ايجابي"
    if any(x in s for x in ["سلبي", "negative", "neg", "-"]):
        return "سلبي"
    if any(x in s for x in ["رفض", "refuse", "refused"]):
        return None
    return v


def clean_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, CleanReport]:
    report = CleanReport(original_rows=len(df), original_cols=len(df.columns))
    out = df.copy()

    new_cols: List[str] = []
    seen: Dict[str, int] = {}
    for i, c in enumerate(out.columns):
        base = _clean_col_name(c, i)
        if base in seen:
            seen[base] += 1
            base = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        new_cols.append(base)
    out.columns = new_cols

    before_drop = len(out)
    out = out.dropna(how="all").reset_index(drop=True)
    report.dropped_empty_rows = before_drop - len(out)

    memory_data = _load_mapping_memory()
    memory = {}
    memory.update(memory_data.get("global", {}))
    mapping, decisions = detect_column_mapping(out.columns, memory=memory)
    report.mapping_decisions = decisions
    for original, canonical in mapping.items():
        if canonical not in out.columns and original in out.columns:
            out = out.rename(columns={original: canonical})
            report.renamed_columns[original] = canonical
        elif canonical in out.columns and original in out.columns and original != canonical:
            out[canonical] = out[canonical].combine_first(out[original])
            report.renamed_columns[original] = f"{canonical} (دمج)"
            for d in report.mapping_decisions:
                if d.original == original and d.canonical == canonical:
                    d.action = "merged"

    for col in out.select_dtypes(include=["object"]).columns:
        out[col] = out[col].map(lambda x: x.strip() if isinstance(x, str) else x)
        out[col] = out[col].replace({"": None, "nan": None, "NaN": None, "None": None})

    for col in list(out.columns):
        norm = _norm_text(col)
        looks_date = ("تاريخ" in col and "نتيجة" not in col) or re.search(r"\bdate\b", norm)
        looks_followup_date = ("زياره متابعه" in norm or "زيارة متابعة" in col) and not any(k in col for k in ["نتيجة", "واقي", "مزلق", "سرنج", "زهري", "دعم", "ميثادون"])
        if looks_date or looks_followup_date:
            parsed, invalid = _parse_maybe_date(out[col])
            out[col] = parsed
            if invalid:
                report.invalid_dates[col] = invalid

    for col in list(out.columns):
        norm_col = _norm_text(col)
        is_need_question = any(x in norm_col for x in ["هل تحتاج", "تحتاج الي", "تحتاج الى", "هل تم تقديم", "تم تقديم", "need", "needs", "required", "provided"])
        is_qty_col = any(k in col for k in NUMERIC_KEYWORDS) and not is_need_question
        if is_qty_col:
            before_nonempty = out[col].notna()
            converted = pd.to_numeric(out[col], errors="coerce")
            bad_numeric = int((before_nonempty & converted.isna()).sum())
            negative = int((converted < 0).sum())
            if bad_numeric or negative:
                report.negative_quantities[col] = bad_numeric + negative
            out[col] = converted

    for col in list(out.columns):
        norm_col = _norm_text(col)
        is_need_question = any(x in norm_col for x in ["هل تحتاج", "تحتاج الي", "تحتاج الى", "هل تم تقديم", "تم تقديم", "need", "needs", "required", "provided"])
        if any(k in col for k in YES_NO_KEYWORDS) or is_need_question:
            out[col] = out[col].map(_clean_yes_no_value)

    for col in list(out.columns):
        if "نتيجة" in col and "تحليل" in col:
            out[col] = out[col].map(_clean_test_result)

    # Remove Excel formatted tail rows that only contain auto-filled serial/org values.
    # A valid DataBridge row should have at least one real business field.
    meaningful_cols = [c for c in ["تاريخ الزيارة", "الكود المجمع", "السن", "النوع", "نتيجة التحليل"] if c in out.columns]
    if meaningful_cols:
        before_meaningful = len(out)
        meaningful_mask = out[meaningful_cols].notna().any(axis=1)
        out = out[meaningful_mask].reset_index(drop=True)
        report.dropped_empty_rows += before_meaningful - len(out)

    report.organization_profile = _detect_org_profile(out)
    # Re-run profile-specific memory only as a note for now; saved mappings are reused globally and by profile after approval.
    validate_data_profile(out, report)

    report.final_rows = len(out)
    report.final_cols = len(out.columns)
    report.memory_mb = float(out.memory_usage(deep=True).sum() / (1024 * 1024))
    report.missing_core_columns = [c for c in REQUIRED_CORE if c not in out.columns]
    report.missing_optional_columns = [c for c in IMPORTANT_OPTIONAL if c not in out.columns]

    review_count = sum(1 for d in report.mapping_decisions if d.action == "review")
    unknown_count = sum(1 for d in report.mapping_decisions if d.action == "unknown")

    if report.memory_mb > 250:
        report.warnings.append("msg_high_memory")
    if report.missing_core_columns:
        report.status = "needs_review"
        report.errors.append("msg_incomplete_analysis")
    elif report.invalid_dates or report.negative_quantities or report.profile_issues or review_count or unknown_count:
        report.status = "warning"
        if report.invalid_dates:
            report.warnings.append("msg_dates_converted_blank")
        if report.negative_quantities:
            report.warnings.append("msg_qty_converted_blank")
        if report.profile_issues:
            report.warnings.append("msg_profile_validation_suspicious")
        if review_count or unknown_count:
            report.warnings.append("msg_mapping_needs_review")
    else:
        report.notes.append("msg_cleaned_high_confidence")

    return out, report


def generate_standard_df(df: pd.DataFrame) -> pd.DataFrame:
    """Create a unified visit-level standard DataFrame.

    Base visit = one row. Each follow-up date = one extra row, carrying the same
    beneficiary attributes and the follow-up service columns for that visit.
    """
    rows: List[Dict[str, Any]] = []
    base_fields = {
        "beneficiary_id": "الكود المجمع",
        "serial": "مسلسل",
        "visit_date": "تاريخ الزيارة",
        "age": "السن",
        "gender": "النوع",
        "governorate": "محافظة الوصول للمستفيد",
        "area": "منطقة الوصول للمستفيد",
        "hiv_test_result": "نتيجة التحليل",
        "referral": "الاحالة علي العلاج",
        "condoms": "واقيات",
        "lubricants": "مزلقات",
        "syringes": "سرنجات",
        "syphilis": "زهري",
        "psychological_support": "دعم نفسي",
        "methadone": "ميثادون",
    }
    for idx, row in df.iterrows():
        base = {std: row.get(src) if src in df.columns else None for std, src in base_fields.items()}
        base["source_row"] = idx + 2
        base["visit_type"] = "basic"
        rows.append(base.copy())
        for n in range(1, 6):
            date_col = f"زيارة متابعة {n}"
            if date_col in df.columns and pd.notna(row.get(date_col)):
                fu = base.copy()
                fu["visit_type"] = f"followup_{n}"
                fu["visit_date"] = row.get(date_col)
                fu["condoms"] = row.get(f"واقيات متابعة {n}") if f"واقيات متابعة {n}" in df.columns else None
                fu["lubricants"] = row.get(f"مزلقات متابعة {n}") if f"مزلقات متابعة {n}" in df.columns else None
                fu["syringes"] = row.get(f"سرنجات متابعة {n}") if f"سرنجات متابعة {n}" in df.columns else None
                fu["syphilis"] = row.get(f"زهري متابعة {n}") if f"زهري متابعة {n}" in df.columns else None
                fu["psychological_support"] = row.get(f"دعم نفسي متابعة {n}") if f"دعم نفسي متابعة {n}" in df.columns else None
                fu["hiv_test_result"] = row.get(f"نتيجة تحليل متابعة {n}") if f"نتيجة تحليل متابعة {n}" in df.columns else None
                fu["methadone"] = row.get(f"ميثادون متابعة {n}") if f"ميثادون متابعة {n}" in df.columns else None
                rows.append(fu)
    return pd.DataFrame(rows)


def report_to_dataframe(report: Optional[CleanReport], lang: str = "ar") -> pd.DataFrame:
    from .settings import T
    t = lambda key: T[lang].get(key, key)
    if report is None:
        return pd.DataFrame(columns=[t("rpt_col_check"), t("rpt_col_result"), t("rpt_col_type")])
    return pd.DataFrame(report.to_rows(lang))


def mapping_report_to_dataframe(report: Optional[CleanReport], lang: str = "ar") -> pd.DataFrame:
    from .settings import T
    t = lambda key: T[lang].get(key, key)
    cols = [t("rpt_col_original"), t("rpt_col_canonical"), t("rpt_col_confidence"), t("rpt_col_method"), t("gap_status"), t("rpt_note")]
    if report is None:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame([
        {
            cols[0]: d.original,
            cols[1]: d.canonical,
            cols[2]: d.confidence,
            cols[3]: d.method,
            cols[4]: _action_label(d.action),
            cols[5]: d.note,
        }
        for d in report.mapping_decisions
    ])
