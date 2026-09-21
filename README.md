# رساله علمی سطح ۳ — هادی شبستانی

**عنوان:** عدول ولی فقیه از تعهدات بین‌المللی دولت اسلامی از منظر فقه امامیه و حقوق بین‌الملل

## فایل اصلی

- `Resaleh_Sath3_Hadi_Shabestani_Final2.docx` — نسخه نهایی رساله (Word)؛ حدود ۱۱۰ صفحه، ۳۶ هزار کلمه، ۱۴۴ پانوشت، بدون جدول
- `Resaleh_Sath3_Hadi_Shabestani_Final2.pdf` — همان رساله به‌صورت PDF (۱۲۲ صفحه؛ فهرست و شماره صفحات خودسازگار)
- `Resaleh_Sath3_Hadi_Shabestani.docx` — نسخه اول (بایگانی)

## ساختار سورس

- `thesis_src/00_front.md` — جلد داخلی، بسمله، تقدیم، سپاس، چکیده فارسی، پیشگفتار، چکیده انگلیسی
- `thesis_src/01_moqaddame.md` — مقدمه: کلیات تحقیق (مسئله، ضرورت، اهداف، سؤالات، فرضیه، پیشینه، روش، ساماندهی)
- `thesis_src/02_bakhsh1_fasl1.md` — بخش اول، فصل اول: مفهوم‌شناسی و چارچوب نظری (۵ گفتار)
- `thesis_src/03_bakhsh1_fasl2.md` — بخش اول، فصل دوم: کلیات حقوق معاهدات و قانون اساسی (۵ گفتار)
- `thesis_src/04_bakhsh2_fasl3.md` — بخش دوم، فصل سوم: ادله وجوب وفای به عهد (۸ گفتار)
- `thesis_src/05_bakhsh2_fasl4.md` — بخش دوم، فصل چهارم: مبانی جواز عدول و ضوابط آن (۶ گفتار)
- `thesis_src/06_bakhsh3_fasl5.md` — بخش سوم، فصل پنجم: مسئولیت بین‌المللی و جایگاه حقوق اساسی (۵ گفتار)
- `thesis_src/07_natije.md` — نتیجه‌گیری (۵ دستاورد، پاسخ سؤال‌های فرعی، نوآوری‌ها و پیشنهادها)
- `thesis_src/08_manabe.md` — فهرست منابع (۸ دسته + یادداشت‌های کتاب‌شناختی)
- `thesis_src/09_payvast.md` — پیوست‌ها (چک‌لیست تصمیم، اصول قانون اساسی، مواد وین، گاه‌شمار، عهدنامه مالک، واژه‌نامه)
- `thesis_src/build_thesis.py` — سازنده فایل Word (قالب شیوه‌نامه حوزوی: B Lotus / B Titr / B Badr)
- `thesis_src/build_pdf.py` — سازنده فایل PDF از روی Word نهایی (فونت‌های B + DejaVu، صفحه‌بندی مستقل، وارسی خودکار فهرست)

## بازسازی فایل Word

```bash
pip install --break-system-packages python-docx
python3 thesis_src/build_thesis.py
```

## ساخت فایل PDF

```bash
pip install --break-system-packages reportlab arabic_reshaper python-bidi pypdf pymupdf matplotlib fonttools
python3 thesis_src/build_pdf.py
```

## نکته

پس از باز کردن فایل در Word، روی «فهرست مطالب» کلیک راست کرده و Update Field ← Update entire table را بزنید.
