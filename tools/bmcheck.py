import sys, zipfile, re, html
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile("/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani (2) (1).docx")
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)
eng = L.Engine(w_scale=1.05, lh_scale=1.30, h_scale=0.835)
res, total, _ = L.layout(paras, st, fns, eng)
off = res[240][2] - 1

body = z.read("word/document.xml").decode("utf-8")
spans = [(m.start(), m.end()) for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S)]
bm_para = {}
for m in re.finditer(r'<w:bookmarkStart[^>]*w:name="(_Toc\d+)"[^>]*/>', body):
    for k, (a, b) in enumerate(spans):
        if a <= m.start() < b:
            bm_para[m.group(1)] = k
            break

# every TOC line: its text, its bookmark, the bookmarked paragraph text, computed page
bad = 0
print(f"{'TOC entry':58s} {'bm':12s} {'para':>4s} {'pg':>3s}  bookmarked paragraph text")
for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S):
    blk = m.group(0)
    if 'w:pStyle w:val="TOC' not in blk:
        continue
    txt = html.unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", blk, re.S)))
    pb = re.search(r"PAGEREF (_Toc\d+)", blk)
    if not pb:
        print(f"{txt[:56]:58s} {'NO FIELD':12s} {'-':>4s} {'-':>3s}  -")
        continue
    name = pb.group(1)
    k = bm_para.get(name)
    if k is None:
        print(f"{txt[:56]:58s} {name:12s} {'MISSING':>4s} {'-':>3s}  bookmark not in body!")
        bad += 1
        continue
    ptxt = paras[k]["text"].strip()
    pg = res[k][2] - (off if k > 239 else 0)
    entry = re.sub(r"[۰-۹]+$", "", txt).strip()
    match = "OK " if entry[:18] == ptxt[:18] else "??"
    if match == "??":
        bad += 1
        print(f"{match} {txt[:52]:56s} {name:12s} {k:>4d} {pg:>3d}  {ptxt[:44]}")
print("\nmismatched / missing:", bad)
for t in ["پیوست‌ها", "پیوست اول: چک‌لیست عملی تصمیم‌گیری درباره عدول از تعهد بین‌المللی",
          "پیوست دوم: گزیده اصول قانون اساسی مرتبط", "پیوست سوم: ترجمه و خلاصه مواد کلیدی کنوانسیون وین",
          "پیوست ششم: واژه‌نامه فارسی ـ عربی ـ انگلیسی", "فهرست منابع", "نتیجه‌گیری"]:
    for i, p in enumerate(paras):
        if p["text"].strip() == t:
            print(f"  body heading [{i}] {t[:46]:48s} -> displayed {res[i][2] - (off if i > 239 else 0)}")
            break
