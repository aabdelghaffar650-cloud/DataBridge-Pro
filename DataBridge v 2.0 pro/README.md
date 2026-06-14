# DataBridge v2.2 Pro — Auto Data Mapper

تشغيل البرنامج:

```powershell
streamlit run app.py --server.port 8503
```

## الجديد في v2.2

- Smart Sheet Detection: اختيار الشيت الأقرب للبيانات تلقائيًا.
- Smart Header Detection: اكتشاف صف العناوين الحقيقي حتى لو فوقه عناوين/صفوف فاضية.
- Column Name Normalizer: تنظيف أسماء الأعمدة عربي/إنجليزي.
- Dictionary Mapping Engine: ربط أسماء الأعمدة المختلفة بالأسماء القياسية.
- Fuzzy Matching: تحمل أخطاء الكتابة وتشابه الأسماء بدون الاعتماد على AI.
- Confidence Report: عرض درجة ثقة لكل عمود.
- Standard Format Generator: إنتاج `standard_df` داخلي بنظام زيارة أساسية + زيارات متابعة.

## ملاحظات

- AI Mapper لم يتم تفعيله بعد، وسيظل مرحلة مستقبلية عند انخفاض الثقة فقط.
- راجع تقرير Auto Data Mapper قبل الاعتماد على الأرقام لو ظهرت أعمدة Review أو Unknown.

## v2.2.2 Auto Mapper Engine additions
- Confidence Dashboard before approval: Auto Mapped / Need Review / Unknown / Overall Mapping Confidence.
- Mapping Explanation: each mapped column shows why it was mapped (dictionary, fuzzy, rule, memory, etc.).
- Data Profile Validation: checks column values after mapping to catch suspicious mappings such as Age columns containing Gender values.
- Learning Memory: approved mappings are saved to `mapping_memory.json` and reused in future uploads by global/profile mapping.

## v2.2.3 Mapper UX Update
- Mapping states are now separated into: Auto Accepted (95%+), Verify (85-94%), Suspicious (70-84%), and Unknown (<70%).
- Synonym dictionary matches in the 85-94% range are shown as Verify instead of generic Review.
- Data Profile Validation issues are marked Suspicious.
