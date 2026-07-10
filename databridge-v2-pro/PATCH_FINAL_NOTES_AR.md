# DataBridge Pro — Final Tauri Ready Patch

آخر نسخة قبل PyInstaller/Tauri.

## الملفات التي يجب استبدالها
- app.py
- modules/reports.py
- databridge-backend.spec

## آخر إصلاحات مؤكدة
- إزالة Q2 من تقارير Year-to-Date / Annual / Selected Period مع بقائها فقط عندما يكون النطاق Q2 فعلاً.
- إصلاح عبارة: `against an annual target`.
- تنظيف Metadata الخاصة بملفات Word حتى لا تحتفظ بعنوان القالب القديم.
- إزالة الـ double bullets في قسم Data Quality وConclusion.
- الحفاظ على Template Fill للتقرير الشهري الرسمي والربع سنوي الرسمي.
- الحفاظ على شكل Summary Excel الشهري و Q2 كما هو.
- إصلاح 905 rows.
- ظهور Targets داخل M&E Report.
