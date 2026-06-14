"""
DataBridge — Chart Helpers & Gender Analytics
==============================================
دوال مساعدة للرسوم البيانية وحساب مستفيدي الخدمات حسب الجنس.

يحل محل:
- show_bar_values / show_pie_values / show_line_values في app.py
- _recipients_by_gender (كانت معطّلة وترجع 0 دائماً)
- دوال _gender_is_male / _gender_is_female
"""

from __future__ import annotations

import pandas as pd
from typing import Any, Optional, List

import plotly.graph_objects as go


# ══════════════════════════════════════════════════════════════════
#  CHART LABEL HELPERS
# ══════════════════════════════════════════════════════════════════

def show_bar_values(fig: go.Figure) -> go.Figure:
    """إظهار القيم الرقمية فوق الأعمدة لتبقى ظاهرة عند التصدير."""
    fig.update_traces(texttemplate='%{y}', textposition='outside', cliponaxis=False)
    fig.update_layout(uniformtext_minsize=10, uniformtext_mode='show')
    return fig


def show_pie_values(fig: go.Figure) -> go.Figure:
    """إظهار التسمية + القيمة + النسبة على الدوائر."""
    fig.update_traces(textinfo='label+value+percent', textposition='auto')
    return fig


def show_line_values(fig: go.Figure) -> go.Figure:
    """إظهار القيم الرقمية على نقاط الخط البياني."""
    fig.update_traces(mode='lines+markers+text', texttemplate='%{y}', textposition='top center')
    return fig


# ══════════════════════════════════════════════════════════════════
#  GENDER DETECTION
# ══════════════════════════════════════════════════════════════════

def _is_male(value: Any) -> bool:
    s = str(value).strip().replace(' ', '')
    return s in {'ذكر', 'male', 'Male', 'M'}


def _is_female(value: Any) -> bool:
    s = str(value).strip().replace(' ', '')
    return s in {'أنثى', 'انثى', 'أنثي', 'انثي', 'female', 'Female', 'F'}


# ══════════════════════════════════════════════════════════════════
#  GENDER RECIPIENTS — الإصلاح الرئيسي
# ══════════════════════════════════════════════════════════════════

def count_recipients_by_gender(
    base_df: pd.DataFrame,
    full_df: pd.DataFrame,
    tool_col_base: Optional[str],
    tool_name: str,
    want_male: bool,
    month_filter_kind: str,
    selected_month: Optional[str],
    from_month: Optional[str],
    to_month: Optional[str],
) -> int:
    """
    حساب عدد المستفيدين (ذكور أو إناث) الذين استلموا خدمة معينة،
    مع مراعاة زيارات المتابعة 1-5.

    كانت الدالة القديمة (_recipients_by_gender) ترجع 0 دائماً.
    هذه الدالة تحل المشكلة بالحساب الفعلي من:
      1. الزيارات الأساسية (base_df)
      2. كل زيارات المتابعة 1-5 (full_df)

    المعاملات:
        base_df         : DataFrame مفلتر بالشهر/النطاق (الزيارات الأساسية فقط)
        full_df         : DataFrame الكامل (لاستخراج متابعات بتاريخها الخاص)
        tool_col_base   : اسم عمود الخدمة في الزيارة الأساسية (None → تخطّ)
        tool_name       : اسم الخدمة عربياً لاستخراج أعمدة المتابعة (مثل 'سرنجات')
        want_male       : True للذكور، False للإناث
        month_filter_*  : معاملات فلتر الشهر لتطبيقها على تواريخ المتابعة
    """
    gender_fn = _is_male if want_male else _is_female

    # 1) الزيارات الأساسية
    base_count = 0
    gender_col = _find_gender_col(base_df)
    if tool_col_base and gender_col:
        has_service = pd.to_numeric(base_df[tool_col_base], errors='coerce').fillna(0) > 0
        gender_match = base_df[gender_col].apply(gender_fn)
        base_count = int((has_service & gender_match).sum())

    # 2) زيارات المتابعة 1-5
    fu_count = 0
    gender_col_full = _find_gender_col(full_df)
    if gender_col_full:
        norm = _norm_ar
        for n in range(1, 6):
            fu_date_col = _find_followup_date_col(full_df.columns, n)
            fu_tool_col = _find_service_col(list(full_df.columns), tool_name, n)
            if not fu_date_col or not fu_tool_col:
                continue
            period_mask = _date_in_period(
                full_df[fu_date_col], month_filter_kind,
                selected_month, from_month, to_month
            )
            has_service = pd.to_numeric(full_df[fu_tool_col], errors='coerce').fillna(0) > 0
            gender_match = full_df[gender_col_full].apply(gender_fn)
            fu_count += int((period_mask & has_service & gender_match).sum())

    return base_count + fu_count


def qty_by_gender(
    base_df: pd.DataFrame,
    full_df: pd.DataFrame,
    tool_col_base: Optional[str],
    tool_name: str,
    want_male: bool,
    month_filter_kind: str,
    selected_month: Optional[str],
    from_month: Optional[str],
    to_month: Optional[str],
) -> int:
    """
    حساب إجمالي الكميات (ليس عدد المستفيدين) حسب الجنس،
    للزيارات الأساسية + المتابعات 1-5.
    """
    gender_fn = _is_male if want_male else _is_female

    # 1) الأساسي
    base_qty = 0
    gender_col = _find_gender_col(base_df)
    if tool_col_base and gender_col:
        gender_mask = base_df[gender_col].apply(gender_fn)
        base_qty = int(pd.to_numeric(base_df.loc[gender_mask, tool_col_base], errors='coerce').fillna(0).sum())

    # 2) المتابعات
    fu_qty = 0
    gender_col_full = _find_gender_col(full_df)
    if gender_col_full:
        for n in range(1, 6):
            fu_date_col = _find_followup_date_col(full_df.columns, n)
            fu_tool_col = _find_service_col(list(full_df.columns), tool_name, n)
            if not fu_date_col or not fu_tool_col:
                continue
            period_mask = _date_in_period(
                full_df[fu_date_col], month_filter_kind,
                selected_month, from_month, to_month
            )
            gender_mask = full_df[gender_col_full].apply(gender_fn)
            combined = period_mask & gender_mask
            fu_qty += int(pd.to_numeric(full_df.loc[combined, fu_tool_col], errors='coerce').fillna(0).sum())

    return base_qty + fu_qty


# ══════════════════════════════════════════════════════════════════
#  INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════

def _norm_ar(text: Any) -> str:
    return str(text).replace('ة', 'ه').strip()


def _find_gender_col(df: pd.DataFrame) -> Optional[str]:
    for c in df.columns:
        if 'النوع' in c:
            return c
    return None


def _find_followup_date_col(columns, n: int) -> Optional[str]:
    exclude = {'واقيات', 'مزلقات', 'زهري', 'دعم', 'سرنجات', 'ميثادون', 'نتيجة'}
    for c in columns:
        nc = _norm_ar(c)
        if _norm_ar('زيارة متابعة') in nc and str(n) in nc:
            if not any(_norm_ar(x) in nc for x in exclude):
                return c
    return None


def _find_service_col(
    cols: List[str],
    keyword: str,
    followup_no: Optional[int] = None,
) -> Optional[str]:
    norm_kw = _norm_ar(keyword)
    for c in cols:
        nc = _norm_ar(c)
        if norm_kw not in nc:
            continue
        if followup_no is None:
            if 'متابع' not in nc:
                return c
        else:
            if 'متابع' in nc and str(followup_no) in nc:
                return c
    return None


def _date_in_period(
    dt_series: pd.Series,
    month_filter_kind: str,
    selected_month: Optional[str],
    from_month: Optional[str],
    to_month: Optional[str],
) -> pd.Series:
    dt = pd.to_datetime(dt_series, errors='coerce')
    if month_filter_kind == "single" and selected_month:
        return dt.dt.to_period('M').astype(str).eq(selected_month)
    if month_filter_kind == "range" and from_month and to_month:
        ps = dt.dt.to_period('M').astype(str)
        return ps.ge(from_month) & ps.le(to_month)
    return dt.notna()
