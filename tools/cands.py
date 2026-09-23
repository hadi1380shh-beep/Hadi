import sys, zipfile, re, html
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)
eng = L.Engine(w_scale=1.05, lh_scale=1.30, h_scale=0.835)
res, total, _ = L.layout(paras, st, fns, eng)
off = res[240][2] - 1

# ---- marker-implied page numbers (all breaks counted, dups removed)
impl = 1
marker_page = {}
prev_was_break = False
for i, p in enumerate(paras):
    evts = sorted([(m, "M") for m in p["marks"]] + [(b, "B") for b in p["breaks"]])
    if not evts:
        if p["text"].strip():
            marker_page.setdefault(i, impl)
    for pos, kind in evts:
        if kind == "M" and prev_was_break:
            prev_was_break = False
            marker_page.setdefault(i, impl)
            continue
        impl += 1
        prev_was_break = True
        marker_page.setdefault(i, impl)
    if p["text"].strip():
        marker_page.setdefault(i, impl)
    if p["sect"]:
        impl += 1
        prev_was_break = True
marker_off = marker_page[240] - 1

# ---- current cached TOC numbers
body = z.read("word/document.xml").decode("utf-8")
toc = {}
for m in re.finditer(r'PAGEREF (_Toc\d+) \\h .*?<w:fldChar w:fldCharType="separate"/></w:r>'
                     r'<w:r>(?:(?!</w:r>).)*?<w:t(?:\s[^>]*)?>([^<]*)</w:t>', body, re.S):
    toc[m.group(1)] = m.group(2)
spans = [(mm.start(), mm.end()) for mm in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S)]
bm = {}
for mm in re.finditer(r'<w:bookmarkStart[^>]*w:name="(_Toc\d+)"[^>]*/>', body):
    for k, (a, b) in enumerate(spans):
        if a <= mm.start() < b:
            bm[k] = mm.group(1)
            break

targets = ["چکیده", "پیشگفتار", "Abstract", "فهرست مطالب", "نتیجه‌گیری", "فهرست منابع",
           "فصل پنجم: مسئولیت بین‌المللی دولت و جایگاه حقوق اساسی تصمیم",
           "بخش دوم: تحلیل فقهی مسئله", "گفتار چهارم: اسباب رافع، آثار مسئولیت و راهبردهای کاهش آن",
           "بند چهارم: تحلیل اصل ۱۱۰؛ نقد ادعای حصر و تفکیک انواع اختیار"]
print(f"{'heading':62s} {'sim':>4s} {'mark':>5s} {'TOC':>5s}")
for i, p in enumerate(paras):
    t = p["text"].strip()
    if t in targets:
        sim = res[i][2] - (off if i > 239 else 0)
        mk = marker_page.get(i, 0) - (marker_off if i > 239 else 0)
        print(f"{t[:60]:62s} {sim:4d} {mk:5d} {toc.get(bm.get(i,''),''):>5s}")
