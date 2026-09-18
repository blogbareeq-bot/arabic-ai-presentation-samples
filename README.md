# نماذج عروض تقديمية عربية بالذكاء الاصطناعي — Arabic AI Presentation Samples

مستودع عام لتحميل ملفات PowerPoint المعتمدة. كل ملف يُفحص آليًا قبل الدمج: سلامة بنية
OOXML، تطابق البصمات، وجاهزية العربية (اتجاه RTL وخط النصوص المركّبة).

Public download repository for the approved PowerPoint sample files. Every deck is
machine-verified before merge: OOXML integrity, checksum fidelity, and Arabic
readiness (RTL direction and complex-script font).

## المحتويات / Layout

```
├── BUNDLE-MANIFEST.json      # بيان موقّع: الحجم، SHA-256، عدد الشرائح لكل ملف
├── SHA256SUMS.txt            # بصمات جاهزة لـ sha256sum -c
├── tools/pptx-bundle.mjs     # أداة الفحص والبيان (بلا أي اعتماديات)
└── .github/workflows/        # بوابة تحقق تعمل على كل دفعة وطلب دمج
```

## التحقق / Verify

```bash
npm test                # اختبارات ذاتية للأدوات (تبني ملفات تالفة وتتحقق من رفضها)
npm run verify          # فحص البنية والبصمات
npm run verify:strict   # نفس الفحص مع إلزام جاهزية العربية
npm run manifest        # إعادة توليد البيان والبصمات بعد إضافة ملفات
npm run manifest:check  # إثبات أن البيان مطابق للملفات الحالية
```

للتحقق اليدوي من بصمة ملف واحد:

```bash
sha256sum -c SHA256SUMS.txt
```

### ما يفحصه `verify`

| الفحص | النوع |
| --- | --- |
| قراءة ZIP وسلامة CRC لكل جزء | خطأ |
| وجود `[Content_Types].xml` و`_rels/.rels` و`ppt/presentation.xml` | خطأ |
| تطابق عدد الشرائح المعلَن مع ملفات الشرائح، وربط كل شريحة بعلاقة | خطأ |
| تطابق SHA-256 والحجم مع `BUNDLE-MANIFEST.json` و`SHA256SUMS.txt` | خطأ |
| وجود ملف في المستودع وغير مُدرج في البيان (أو العكس) | خطأ |
| شريحة بها نص عربي بدون خط نصوص مركّبة `<a:cs>` | تنبيه (خطأ في `--strict`) |
| شريحة بلا نص عربي، أو فقرة بلا `rtl="1"`، أو غياب صفحات الملاحظات | تنبيه |

رموز الخروج: `0` سليم، `1` يوجد أخطاء، `2` خطأ استخدام.

## إضافة نموذج جديد / Adding a deck

الطريقة الأولى — دمج حزمة كاملة:

```bash
npm run import -- حزمة-النماذج.zip --into decks
npm run import -- /path/to/folder --into decks --dry-run   # تجربة بلا كتابة
```

أداة الدمج تستخرج ملفات `.pptx` فقط، تتجاهل `__MACOSX` والملفات المؤقتة، تمنع
الهروب من مجلد المستودع (zip-slip)، ترفض أي ملف تالف قبل الكتابة، وتكشف
المكرر بالبصمة فلا تُضاف نسخة ثانية، ثم تعيد توليد البيان والبصمات تلقائيًا.

الطريقة الثانية — ملفًا بملف:

1. ضع ملف `.pptx` المعتمد داخل المستودع (أي مجلد؛ تُفحص كل الملفات).
2. شغّل `npm run manifest` لتحديث البيان والبصمات.
3. ادفع التغيير — بوابة GitHub Actions لن تسمح بالدمج قبل نجاح الفحص.

> الأداة بلا اعتماديات خارجية وتعمل على Node.js 20 فأحدث.
> لا تُدرج ملفات الأقفال المؤقتة (`~$*.pptx`) — يتجاهلها الفحص تلقائيًا.
