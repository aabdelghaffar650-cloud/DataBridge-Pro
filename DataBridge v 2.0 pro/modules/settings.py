"""Translations and app settings."""

APP_VERSION = "v2.3.6 Pro — AI Settings Manager"
MAX_HISTORY = 10
MAX_UPLOAD_SIZE_MB = 100
MAX_HISTORY_MEM_MB = 500
USERS_FILE = "users.json"


def get_appdata_dir() -> str:
    """Writable per-user directory for DataBridge runtime files."""
    import os
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".databridge")
    path = os.path.join(base, "DataBridge") if os.environ.get("APPDATA") else base
    os.makedirs(path, exist_ok=True)
    return path


def get_config_path() -> str:
    import os
    return os.path.join(get_appdata_dir(), "config.json")


def load_config() -> dict:
    import json, os
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                data.setdefault("ai_provider", "gemini")
                data.setdefault("gemini_model", "gemini-2.5-flash")
                data.setdefault("gemini_api_key", "")
                return data
            return {"ai_provider": "gemini", "gemini_model": "gemini-2.5-flash", "gemini_api_key": ""}
        except Exception:
            pass
    return {"ai_provider": "gemini", "gemini_model": "gemini-2.5-flash", "gemini_api_key": ""}


def save_config(cfg: dict) -> None:
    import json
    path = get_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


T = {
    "ar": {
        "app_name": "DataBridge",
        "app_sub": "M&E Hub for Health Programs",
        "app_slogan": "منصة المتابعة والتقييم المتكاملة",
        "login_title": "تسجيل الدخول",
        "username": "اسم المستخدم",
        "password": "كلمة المرور",
        "login_btn": "دخول",
        "login_error": "❌ اسم المستخدم أو كلمة المرور غلط",
        "logout": "تسجيل خروج",
        "upload_label": "ارفع ملف IDUs Database (Excel)",
        "upload_hint": "📋 ارفع ملف IDUs Database وهيتحلل تلقائياً",
        "upload_hint2": "✅ البرنامج عارف هيكل الشيت وهيحلل الفراغات المنطقية تلقائياً",
        "tab_data": "📋 البيانات",
        "tab_gaps": "🔍 تحليل الفراغات",
        "tab_stats": "📊 الإحصائيات",
        "tab_export": "💾 التصدير",
        "total": "إجمالي المستفيدين",
        "has_followup": "لديهم متابعة",
        "no_followup": "بدون متابعة",
        "positive": "نتيجة إيجابية",
        "refused": "رفضوا التحليل",
        "beneficiary": "مستفيد",
        "positive_case": "حالة إيجابية",
        "from_total": "من الإجمالي",
        "show_label": "عرض:",
        "all_data": "كل البيانات",
        "without_followup": "بدون متابعة",
        "with_followup": "مع متابعة",
        "positive_cases": "الحالات الإيجابية",
        "rows_count": "عدد الصفوف",
        "gaps_info": "الفراغات دي مش غلط — كل فراغ ليه سبب منطقي موضح أدناه",
        "gaps_report": "تقرير الفراغات",
        "no_gaps": "✅ لا توجد فراغات!",
        "col_name": "العمود",
        "gap_count": "عدد الفراغات",
        "gap_pct": "نسبة %",
        "gap_reason": "سبب الفراغ المنطقي",
        "gap_status": "الحالة",
        "justified": "✅ مبرر منطقياً",
        "needs_review": "⚠️ يحتاج مراجعة",
        "export_csv": "📥 تصدير CSV",
        "export_excel": "📥 تصدير Excel مع التقرير",
        "save_excel_title": "💾 حفظ في شيت Excel",
        "upload_db": "ارفع ملف IDUs Database Excel",
        "save_all_btn": "✅ أضف كل السجلات للشيت",
        "save_success": "تم الحفظ بنجاح!",
        "download_updated": "⬇️ حمّل الشيت المحدّث",
        "chart_test": "نتائج التحليل",
        "chart_followup": "توزيع المتابعة",
        "chart_services": "الخدمات المقدمة (من تلقوا الخدمة فعلاً)",
        "chart_age": "توزيع الفئات العمرية",
        "chart_area": "التوزيع الجغرافي — منطقة الوصول",
        "chart_gov": "توزيع المحافظات",
        "chart_age_prot": "أدوات الوقاية حسب الفئة العمرية",
        "chart_monthly": "الزيارات الشهرية",
        "with_followup_label": "مع متابعة",
        "without_followup_label": "بدون متابعة",
        "refused_test": "رفض التحليل",
        "file_error": "خطأ في تحميل الملف",
        "lang_btn": "English",
        "section_title": "IDUs Database — Be Frienders 2026",
        "col_reason_confirm": "التحليل سلبي أو لم يُجرَ تحليل",
        "col_reason_referral": "التحليل سلبي أو لم يُجرَ تحليل",
        "col_reason_test_fu": "لا توجد متابعة أو لم يُجرَ تحليل",
        "col_reason_condoms_fu": "لا توجد متابعة",
        "col_reason_lube_fu": "لا توجد متابعة",
        "col_reason_syph_fu": "رفض المستفيد أو لا توجد متابعة",
        "col_reason_psycho_fu": "رفض المستفيد أو لا توجد متابعة",
        "col_reason_syr_fu": "لا توجد متابعة",
        "col_reason_meth_fu": "رفض المستفيد أو لا توجد متابعة",
        "col_reason_default": "رفض المستفيد",
        # ── Phase 2 ──
        "tab_quality": "🔎 جودة البيانات",
        "dq_title": "محرك جودة البيانات",
        "dq_subtitle": "فحص شامل لكل سجل وفق القواعد المنطقية المعتمدة",
        "dq_errors": "خطأ جودة بيانات",
        "dq_warnings": "تحذير",
        "dq_ok": "سجل صحيح",
        "dq_col_row": "رقم السجل",
        "dq_col_field": "الحقل",
        "dq_col_value": "القيمة",
        "dq_col_rule": "القاعدة المنتهكة",
        "dq_col_severity": "الخطورة",
        "dq_col_notes": "ملاحظات",
        "dq_save_notes": "💾 حفظ الملاحظات",
        "dq_notes_saved": "✅ تم حفظ الملاحظات",
        "dq_export": "📥 تصدير تقرير الجودة",
        "dq_filter_label": "فلتر حسب الخطورة:",
        "dq_all": "الكل",
        "dq_err_age": "العمر خارج النطاق المسموح (15-70 أو فئات محددة)",
        "dq_err_age_u15": "عمر أقل من 15 — يلزم ملاحظة ولي الأمر",
        "dq_err_positive_referral": "نتيجة إيجابية بدون إحالة للعلاج",
        "dq_err_followup_nodate": "توجد بيانات متابعة بدون تاريخ متابعة",
        "dq_err_negative_qty": "كمية سالبة غير مقبولة",
        "dq_err_duplicate_serial": "رقم مسلسل مكرر",
        "dq_err_duplicate_code": "كود مجمع مكرر",
        "dq_err_code_mismatch": "الكود المجمع لا يطابق بيانات تاريخ الميلاد أو المحافظة",
        "dq_warn_positive_noconfirm": "نتيجة إيجابية بدون تحليل تأكيدي (نادر — قد يكون رفض المستفيد)",
        "dq_err_invalid_code_format": "تنسيق الكود المجمع غير صحيح",
        "dq_notes_placeholder": "أضف ملاحظة...",
        "dq_no_issues": "✅ لا توجد مشاكل في جودة البيانات!",
        "dq_run_btn": "🔍 تشغيل فحص الجودة",
        "dq_show_errors_only": "🔴 الأخطاء فقط",
        "dq_show_warnings_only": "🟡 التحذيرات فقط",
        "dq_show_all_issues": "كل المشاكل",
        "dq_records_with_issues": "سجلات بها مشاكل",
    },
    "en": {
        "app_name": "DataBridge",
        "app_sub": "M&E Hub for Health Programs",
        "app_slogan": "Complete Monitoring & Evaluation Platform",
        "login_title": "Login",
        "username": "Username",
        "password": "Password",
        "login_btn": "Sign In",
        "login_error": "❌ Incorrect username or password",
        "logout": "Sign Out",
        "upload_label": "Upload IDUs Database (Excel)",
        "upload_hint": "📋 Upload your IDUs Database file — auto-analyzed on upload",
        "upload_hint2": "✅ The system recognizes this sheet structure and validates logical gaps automatically",
        "tab_data": "📋 Data",
        "tab_gaps": "🔍 Gap Analysis",
        "tab_stats": "📊 Statistics",
        "tab_export": "💾 Export",
        "total": "Total Beneficiaries",
        "has_followup": "With Follow-up",
        "no_followup": "No Follow-up",
        "positive": "Positive Result",
        "refused": "Refused Testing",
        "beneficiary": "beneficiary",
        "positive_case": "positive case",
        "from_total": "of total",
        "show_label": "View:",
        "all_data": "All Data",
        "without_followup": "No Follow-up",
        "with_followup": "With Follow-up",
        "positive_cases": "Positive Cases",
        "rows_count": "Row count",
        "gaps_info": "These gaps are not errors — each has a documented logical reason",
        "gaps_report": "Gap Report",
        "no_gaps": "✅ No gaps found!",
        "col_name": "Column",
        "gap_count": "Gap Count",
        "gap_pct": "Percentage %",
        "gap_reason": "Logical Reason",
        "gap_status": "Status",
        "justified": "✅ Logically Justified",
        "needs_review": "⚠️ Needs Review",
        "export_csv": "📥 Export CSV",
        "export_excel": "📥 Export Excel with Report",
        "save_excel_title": "💾 Save to Excel Sheet",
        "upload_db": "Upload IDUs Database Excel",
        "save_all_btn": "✅ Add All Records to Sheet",
        "save_success": "Saved successfully!",
        "download_updated": "⬇️ Download Updated Sheet",
        "chart_test": "Test Results",
        "chart_followup": "Follow-up Distribution",
        "chart_services": "Services Delivered",
        "chart_age": "Age Group Distribution",
        "chart_area": "Geographic Distribution — Area",
        "chart_gov": "Governorate Distribution",
        "chart_age_prot": "Protective Tools by Age Group",
        "chart_monthly": "Monthly Visits",
        "with_followup_label": "With Follow-up",
        "without_followup_label": "No Follow-up",
        "refused_test": "Refused Testing",
        "file_error": "Error loading file",
        "lang_btn": "العربية",
        "section_title": "IDUs Database — Be Frienders 2026",
        "col_reason_confirm": "Negative test or no test performed",
        "col_reason_referral": "Negative test or no test performed",
        "col_reason_test_fu": "No follow-up or no test performed",
        "col_reason_condoms_fu": "No follow-up visit",
        "col_reason_lube_fu": "No follow-up visit",
        "col_reason_syph_fu": "Refused or no follow-up",
        "col_reason_psycho_fu": "Refused or no follow-up",
        "col_reason_syr_fu": "No follow-up visit",
        "col_reason_meth_fu": "Refused or no follow-up",
        "col_reason_default": "Beneficiary refused",
        # ── Phase 2 ──
        "tab_quality": "🔎 Data Quality",
        "dq_title": "Data Quality Engine",
        "dq_subtitle": "Comprehensive record-level validation against approved business rules",
        "dq_errors": "Data Quality Errors",
        "dq_warnings": "Warnings",
        "dq_ok": "Valid Records",
        "dq_col_row": "Record #",
        "dq_col_field": "Field",
        "dq_col_value": "Value",
        "dq_col_rule": "Rule Violated",
        "dq_col_severity": "Severity",
        "dq_col_notes": "Notes",
        "dq_save_notes": "💾 Save Notes",
        "dq_notes_saved": "✅ Notes saved",
        "dq_export": "📥 Export Quality Report",
        "dq_filter_label": "Filter by severity:",
        "dq_all": "All",
        "dq_err_age": "Age outside allowed range (15-70 or defined categories)",
        "dq_err_age_u15": "Age under 15 — guardian note required",
        "dq_err_positive_referral": "Positive result without treatment referral",
        "dq_err_followup_nodate": "Follow-up data exists but no follow-up date",
        "dq_err_negative_qty": "Negative quantity not acceptable",
        "dq_err_duplicate_serial": "Duplicate serial number",
        "dq_err_duplicate_code": "Duplicate composite code",
        "dq_err_code_mismatch": "Composite code does not match birth date or governorate data",
        "dq_warn_positive_noconfirm": "Positive result without confirmatory test (rare — may be refusal)",
        "dq_err_invalid_code_format": "Composite code format is invalid",
        "dq_notes_placeholder": "Add a note...",
        "dq_no_issues": "✅ No data quality issues found!",
        "dq_run_btn": "🔍 Run Quality Check",
        "dq_show_errors_only": "🔴 Errors Only",
        "dq_show_warnings_only": "🟡 Warnings Only",
        "dq_show_all_issues": "All Issues",
        "dq_records_with_issues": "Records with Issues",
    }
}


