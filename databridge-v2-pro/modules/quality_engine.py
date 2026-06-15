"""Data quality validation engine and scoring."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

VALID_AGE_GROUPS = [
    '15 - 19', '20 - 24', '25 - 29', '30 - 34',
    '35 - 39', '40 - 44', '45 - 49', '50 او اكثر'
]
AGE_UNDER_15 = ['أقل من 15', 'less than 15', '< 15', 'اقل من 15']

# Governorate abbreviation map (first Arabic letter)
GOV_MAP = {
    'ز': 'الجيزة', 'ق': 'القاهرة', 'ا': 'الاسكندرية',
    'س': 'السويس', 'ش': 'الشرقية', 'غ': 'الغربية',
    'م': 'المنوفية', 'ب': 'البحيرة', 'ك': 'الكفراوية',
    'د': 'الدقهلية', 'ف': 'الفيوم', 'ن': 'بني سويف',
    'ج': 'جنوب سيناء', 'ر': 'الرياض',
}

def _parse_composite_code(code: str) -> dict:
    """
    Parse composite code: DD MM YY GOV_LETTER C1 C2 C3 C4 C5
    Returns dict with parsed fields or empty dict on failure.
    """
    if not isinstance(code, str):
        return {}
    parts = code.strip().split()
    # Minimum: 3 date parts + 1 gov + 5 name chars = 9 parts
    if len(parts) < 9:
        return {}
    try:
        day  = int(parts[0])
        mon  = int(parts[1])
        year = int(parts[2])
        gov  = parts[3]
        name_chars = parts[4:]  # Should be 5 chars
        if not (1 <= day <= 31 and 1 <= mon <= 12):
            return {}
        if len(name_chars) != 5:
            return {}
        return {
            "day": day, "month": mon, "year": year,
            "gov": gov, "name_chars": name_chars, "valid": True
        }
    except (ValueError, IndexError):
        return {}


def run_quality_engine(hdf: pd.DataFrame) -> List[Dict]:
    """
    Run all quality checks. Returns list of issue dicts.
    Each dict: {row_idx, serial, field, value, rule, severity}
    Severity: 'error' | 'warning'
    """
    issues: List[Dict] = []

    # ── Column lookup helpers ──
    def find_col(keyword: str, exclude: List[str] = []) -> Optional[str]:
        for c in hdf.columns:
            if keyword in c and not any(ex in c for ex in exclude):
                return c
        return None

    col_serial   = find_col('مسلسل')
    col_code     = find_col('الكود المجمع')
    col_age      = find_col('السن')
    col_gov      = find_col('محافظة')
    col_test     = find_col('نتيجة التحليل', ['متابعه', 'تاكيدي'])
    col_confirm  = find_col('التاكيدي')
    col_referral = find_col('الاحالة')
    col_fu_date  = find_col('زيارة متابعة')
    col_condoms  = find_col('واقيات', ['متابعه'])
    col_lube     = find_col('مزلقات', ['متابعه'])
    col_syringes = find_col('سرنجات', ['متابعه'])

    # ── Track duplicates ──
    serial_counts = hdf[col_serial].value_counts() if col_serial else pd.Series(dtype=int)
    code_counts   = hdf[col_code].value_counts()   if col_code   else pd.Series(dtype=int)
    seen_serials: set = set()
    seen_codes:   set = set()

    for idx, row in hdf.iterrows():
        serial_val = row.get(col_serial, idx) if col_serial else idx

        def add_issue(field: str, value: Any, rule_key: str, severity: str = 'error'):
            issues.append({
                "row_idx":  idx,
                "serial":   serial_val,
                "field":    field,
                "value":    str(value)[:60],
                "rule_key": rule_key,
                "severity": severity,
                "notes":    "",
            })

        # ── Rule 1: Duplicate serial ──
        if col_serial:
            s_val = row[col_serial]
            if pd.notna(s_val):
                if s_val in seen_serials:
                    add_issue(col_serial, s_val, 'dq_err_duplicate_serial')
                else:
                    seen_serials.add(s_val)

        # ── Rule 2: Duplicate composite code ──
        if col_code:
            c_val = row[col_code]
            if pd.notna(c_val):
                if c_val in seen_codes:
                    add_issue(col_code, c_val, 'dq_err_duplicate_code')
                else:
                    seen_codes.add(c_val)

        # ── Rule 3: Composite code format + cross-check ──
        if col_code:
            raw_code = str(row[col_code]).strip() if pd.notna(row[col_code]) else ''
            if raw_code:
                parsed = _parse_composite_code(raw_code)
                if not parsed:
                    add_issue(col_code, raw_code, 'dq_err_invalid_code_format')
                else:
                    # Cross-check year vs age group
                    if col_age and pd.notna(row[col_age]):
                        age_str = str(row[col_age]).strip()
                        # Extract min age from group string e.g. "20 - 24" → 20
                        try:
                            min_age = int(age_str.split('-')[0].split('او')[0].strip())
                            birth_year_full = 1900 + parsed['year'] if parsed['year'] >= 26 else 2000 + parsed['year']
                            visit_year = 2026
                            expected_age = visit_year - birth_year_full
                            # Allow ±5 year tolerance for age group ranges
                            if not (min_age - 5 <= expected_age <= min_age + 10):
                                add_issue(col_code, raw_code, 'dq_err_code_mismatch')
                        except (ValueError, AttributeError):
                            pass

                    # Note: gov letter in composite code = birth governorate (not in sheet)
                    # Cross-check not possible — skipped by design

        # ── Rule 4: Age validation ──
        if col_age and pd.notna(row[col_age]):
            age_str = str(row[col_age]).strip()
            is_under_15 = any(u in age_str for u in AGE_UNDER_15)
            is_valid_group = any(g in age_str for g in VALID_AGE_GROUPS)

            if is_under_15:
                add_issue(col_age, age_str, 'dq_err_age_u15')
            elif not is_valid_group:
                add_issue(col_age, age_str, 'dq_err_age')

        # ── Rule 5: Positive result → must have referral ──
        if col_test and col_referral:
            test_val    = str(row[col_test]).strip()    if pd.notna(row[col_test])    else ''
            referral_val= str(row[col_referral]).strip() if pd.notna(row[col_referral]) else ''
            if test_val == 'ايجابي' and not referral_val:
                add_issue(col_referral, 'فراغ', 'dq_err_positive_referral')

        # ── Rule 6: Positive → confirmatory test (warning if missing) ──
        if col_test and col_confirm:
            test_val    = str(row[col_test]).strip()   if pd.notna(row[col_test])   else ''
            confirm_val = str(row[col_confirm]).strip() if pd.notna(row[col_confirm]) else ''
            if test_val == 'ايجابي' and not confirm_val:
                add_issue(col_confirm, 'فراغ', 'dq_warn_positive_noconfirm', severity='warning')

        # ── Rule 7: Follow-up data exists → must have follow-up date ──
        if col_fu_date:
            fu_date_val = row[col_fu_date]
            # Check if any follow-up field has data
            fu_cols = [c for c in hdf.columns if 'متابعه' in c]
            has_fu_data = any(pd.notna(row[c]) for c in fu_cols if c in hdf.columns)
            if has_fu_data and pd.isna(fu_date_val):
                add_issue(col_fu_date, 'فراغ', 'dq_err_followup_nodate')

        # ── Rule 8: Negative quantities ──
        for qty_col in [col_condoms, col_lube, col_syringes]:
            if qty_col and pd.notna(row[qty_col]):
                try:
                    val_num = float(row[qty_col])
                    if val_num < 0:
                        add_issue(qty_col, val_num, 'dq_err_negative_qty')
                except (ValueError, TypeError):
                    pass

    return issues


# ────────────────────────────────────────────────────────────────


def compute_quality_score(issues: list, total_records: int) -> dict:
    """
    Critical × 5 + Major × 2 + Minor × 0.5
    Penalty Rate = sum / total × 100
    Score = 100 - Penalty Rate (min 0)
    """
    SEVERITY_MAP = {
        'dq_err_age_u15':               ('critical', 5),
        'dq_err_positive_referral':     ('critical', 5),
        'dq_warn_positive_noconfirm':   ('critical', 5),
        'dq_err_followup_nodate':       ('critical', 5),
        'dq_err_code_mismatch':         ('major',    2),
        'dq_err_age':                   ('major',    2),
        'dq_err_duplicate_code':        ('minor',    0.5),
        'dq_err_duplicate_serial':      ('info',     0),
        'dq_err_invalid_code_format':   ('info',     0),
        'dq_err_negative_qty':          ('info',     0),
    }
    critical = major = minor = info = 0
    penalty = 0.0
    for issue in issues:
        cat, weight = SEVERITY_MAP.get(issue['rule_key'], ('info', 0))
        if cat == 'critical': critical += 1
        elif cat == 'major':  major    += 1
        elif cat == 'minor':  minor    += 1
        else:                 info     += 1
        penalty += weight
    if total_records > 0:
        penalty_rate = (penalty / total_records) * 100
    else:
        penalty_rate = 0
    score = max(0.0, 100.0 - penalty_rate)
    if score >= 95:   status, color = "🟢 Excellent", "#6bff8e"
    elif score >= 90: status, color = "🟢 Good",      "#6bff8e"
    elif score >= 80: status, color = "🟡 Acceptable", "#ffb74d"
    elif score >= 70: status, color = "🟠 Needs Review","#ff9800"
    else:             status, color = "🔴 Poor",       "#ff6b6b"
    valid = total_records - len({i['row_idx'] for i in issues if SEVERITY_MAP.get(i['rule_key'],('info',0))[0] != 'info'})
    return {
        "score": round(score, 1), "status": status, "color": color,
        "critical": critical, "major": major, "minor": minor, "info": info,
        "valid": max(valid, 0), "penalty_rate": round(penalty_rate, 2)
    }

