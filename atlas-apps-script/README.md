# Atlas Dashboard — Google Apps Script

نسخة أونلاين من الداش بورد تعمل على Google Apps Script، بنفس الخطوط والألوان، مع **الوضع الليلي** وإمكانية **تحميل البيانات من Google Sheets** مباشرة.

## خطوات التشغيل

1. افتح [script.google.com](https://script.google.com) ← **New project**
   - أو من داخل Google Sheet: **Extensions ← Apps Script** (ليقرأ الشيت تلقائيًا بدون رابط).
2. الصق محتوى `Code.gs` مكان الكود الافتراضي.
3. اضغط **+** بجوار Files ← **HTML** ← سمِّه `Index` بالضبط، والصق محتوى `Index.html`.
4. **Deploy ← New deployment** ← النوع **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone with Google account** أو **Anyone** حسب الحاجة.
5. وافق على الصلاحيات، وانسخ رابط الـ Web app.

## المزايا

- رفع ملفات CSV / TSV / XLSX / XLS (كما في النسخة الأصلية).
- تحميل Google Sheet: الصق رابط الشيت (يُقرأ التبويب الموجود في الرابط)، أو اتركه فارغًا لقراءة الشيت المرتبط بالسكريبت.
- زر **Night / Day** في الشريط العلوي، ويُحفظ اختيارك تلقائيًا.
- التصدير إلى Excel وPDF (الطباعة تخرج دائمًا بالوضع الفاتح).

> عند أي تعديل على الكود: **Deploy ← Manage deployments ← Edit ← New version** ليظهر التحديث على نفس الرابط.
