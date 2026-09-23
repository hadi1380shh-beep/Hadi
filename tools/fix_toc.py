"""Fix the page numbers in فهرست مطالب -- and nothing else.

Three kinds of repair, all inside the TOC block:
  1. the 176 PAGEREF fields  -> cached result replaced with the computed page
  2. one hard-typed number ("پیوست اول ... ۱۰۳") -> corrected
  3. two TOC lines that had no number at all -> number appended
"""
import sys, zipfile, re, os
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

SRC = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani (2) (1).docx"
DST = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_TOCfixed.docx"
FA = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")

def fa(n):
    return str(n).translate(FA)

# ------------------------------------------------------------------ paginate
z = zipfile.ZipFile(SRC)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)
eng = L.Engine(w_scale=1.05, lh_scale=1.30, h_scale=0.835)
res, total_pages, _ = L.layout(paras, st, fns, eng)
off = res[240][2] - 1

def disp(k):
    return res[k][2] - (off if k > 239 else 0)

body = z.read("word/document.xml").decode("utf-8")
spans = [(m.start(), m.end()) for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S)]
bm_para = {}
for m in re.finditer(r'<w:bookmarkStart[^>]*w:name="(_Toc\d+)"[^>]*/>', body):
    for k, (a, b) in enumerate(spans):
        if a <= m.start() < b:
            bm_para[m.group(1)] = k
            break
page_of = {n: disp(k) for n, k in bm_para.items()}

def body_heading(text, after=239):
    for i in range(after, len(paras)):
        if paras[i]["text"].strip() == text:
            return i, disp(i)
    raise KeyError(text)

# --------------------------------------------------- 1. the PAGEREF fields
FIELD = re.compile(
    r'(<w:instrText xml:space="preserve"> PAGEREF (_Toc\d+) \\h </w:instrText>.*?'
    r'<w:fldChar w:fldCharType="separate"/></w:r>)(<w:r>(?:(?!</w:r>).)*?)'
    r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)', re.S)

# the orphan entry has no bookmark; its section is the renamed heading below
ORPHAN_TEXT = "تطبیق و روش‌شناسی کتاب‌شناختی منابع"
page_of["_Toc000171"] = disp(body_heading(ORPHAN_TEXT)[0])

n_fields = 0
def repl(m):
    global n_fields
    head, name, run_open, t_open, old, t_close = m.groups()
    n_fields += 1
    return head + run_open + t_open + fa(page_of[name]) + t_close

body = FIELD.sub(repl, body)

# ------------------------------------- 2. the hard-typed number (پیوست اول)
pg_p1 = disp(body_heading("پیوست اول: چک‌لیست عملی تصمیم‌گیری درباره عدول از تعهد بین‌المللی")[0])
old_run = ('<w:r w:rsidR="00266FCE"><w:rPr><w:rFonts w:hint="cs"/><w:noProof/>'
           '<w:color w:val="202020"/></w:rPr><w:t>۱۰۳</w:t></w:r>')
new_run = old_run.replace("<w:t>۱۰۳</w:t>", f"<w:t>{fa(pg_p1)}</w:t>")
assert body.count(old_run) == 1, body.count(old_run)
body = body.replace(old_run, new_run)

# ------------------------- 3. the two TOC lines that carry no number at all
pg_pay = disp(body_heading("پیوست‌ها")[0])
pg_p3 = disp(body_heading("پیوست سوم: ترجمه و خلاصه مواد کلیدی کنوانسیون وین")[0])

def add_number(toc_text, colour, page):
    global body
    anchor = (f'<w:t>{toc_text}</w:t></w:r><w:r><w:rPr><w:noProof/>'
              f'<w:color w:val="{colour}"/></w:rPr><w:tab/></w:r></w:p>')
    assert body.count(anchor) == 1, (toc_text, body.count(anchor))
    ins = (f'<w:t>{toc_text}</w:t></w:r><w:r><w:rPr><w:noProof/>'
           f'<w:color w:val="{colour}"/></w:rPr><w:tab/></w:r>'
           f'<w:r><w:rPr><w:noProof/><w:color w:val="{colour}"/></w:rPr>'
           f'<w:t>{fa(page)}</w:t></w:r></w:p>')
    body = body.replace(anchor, ins)

add_number("پیوست‌ها", "26384A", pg_pay)
add_number("پیوست سوم: ترجمه و خلاصه مواد کلیدی کنوانسیون وین", "202020", pg_p3)

# --------------------------------------------------------------- write file
with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = new_body = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = body.encode("utf-8")
        zout.writestr(item, data)

print(f"simulated pages      : {total_pages}")
print(f"PAGEREF fields fixed : {n_fields}")
print(f"hard-typed number    : ۱۰۳ -> {fa(pg_p1)}  (پیوست اول)")
print(f"numbers added        : پیوست‌ها -> {fa(pg_pay)}, پیوست سوم -> {fa(pg_p3)}")
print(f"written              : {DST} ({os.path.getsize(DST)} bytes)")
