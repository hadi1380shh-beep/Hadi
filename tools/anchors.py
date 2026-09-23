import sys, zipfile
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)
W, LH, HS = 1.05, 1.30, 0.835
eng = L.Engine(w_scale=W, lh_scale=LH, h_scale=HS)
res, pages, fnu = L.layout(paras, st, fns, eng)
print("total physical pages:", pages)

# physical page where section 3 starts = page of first paragraph after para 239
sec3_start = res[240][2]
print("section 3 starts at physical page:", sec3_start)
offset = sec3_start - 1
print("displayed = physical -", offset)

targets = [
    "چکیده", "پیشگفتار", "Abstract", "فهرست مطالب", "مقدمه: کلیات تحقیق",
    "گفتار چهارم: اسباب رافع، آثار مسئولیت و راهبردهای کاهش آن",
    "بند چهارم: تحلیل اصل ۱۱۰؛ نقد ادعای حصر و تفکیک انواع اختیار",
    "نتیجه‌گیری", "فهرست منابع", "پیوست‌ها",
]
for i, p in enumerate(paras):
    t = p["text"].strip()
    if t in targets:
        phys = res[i][2]
        disp = phys - offset if i > 239 else phys
        print(f"  phys={phys:4d} disp={disp:4d}  [{i}] {t[:60]}")
