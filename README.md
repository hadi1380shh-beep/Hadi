# رساله علمی سطح ۳ — هادی شبستانی

**عنوان:** قلمرو اختیارات ولی فقیه در تعلیق و نقض یک‌جانبه معاهدات بین‌المللی (مطالعه فقهی ـ حقوقی)

## فایل اصلی
- `Resaleh_Sath3_Hadi_Shabestani.docx` — نسخه نهایی رساله (Word)

## ساختار
- `thesis_src/00_front.md` — تقدیم، سپاس، چکیده فارسی، کلیدواژه‌ها، اختصارات، چکیده عربی
- `thesis_src/01_moqaddame.md` — مقدمه (تبیین مسئله، پیشینه، فرضیه، روش، ساختار)
- `thesis_src/02_bakhsh1_fasl1.md` — بخش اول، فصل اول: مفهوم‌شناسی
- `thesis_src/03_bakhsh1_fasl2.md` — بخش اول، فصل دوم: کلیات و مبانی
- `thesis_src/04_bakhsh2_fasl3.md` — بخش دوم، فصل سوم: ادله وجوب وفای به عهد
- `thesis_src/05_bakhsh2_fasl4.md` — بخش دوم، فصل چهارم: اختیار حاکم و ضوابط عدول
- `thesis_src/06_bakhsh3_fasl5.md` — بخش سوم، فصل پنجم: مسئولیت بین‌المللی و حقوق اساسی
- `thesis_src/07_natije.md` — نتیجه‌گیری
- `thesis_src/08_manabe.md` — فهرست منابع
- `thesis_src/09_payvast.md` — پیوست‌ها (جدول مواد، واژه‌نامه سه‌زبانه)
- `thesis_src/build_thesis.py` — سازنده فایل Word (قالب شیوه‌نامه حوزوی)

## بازسازی فایل Word
```bash
pip install python-docx
python3 thesis_src/build_thesis.py
```

## نکته
پس از باز کردن فایل در Word، روی «فهرست مطالب» کلیک راست کرده و Update Field ← Update entire table را بزنید.
