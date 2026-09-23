import zipfile, re, html

A = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani (2) (1).docx"
B = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_TOCfixed.docx"
a = zipfile.ZipFile(A).read("word/document.xml").decode("utf-8")
b = zipfile.ZipFile(B).read("word/document.xml").decode("utf-8")
print("length old/new:", len(a), len(b))

DIG = re.compile(r"[۰-۹]+")
ma, mb = DIG.sub("#", a), DIG.sub("#", b)
print("document.xml identical after masking every Persian-digit run:", ma == mb)

ta = [t for t in DIG.split(a)]
tb = [t for t in DIG.split(b)]
da = DIG.findall(a)
db = DIG.findall(b)
print("digit runs:", len(da), len(db), "| non-digit segments equal:", ta == tb)
diffs = [(x, y) for x, y in zip(da, db) if x != y]
print("digit runs that changed:", len(diffs))
print("first 5 changes:", diffs[:5])


def toc_lines(x):
    out = []
    for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", x, re.S):
        blk = m.group(0)
        if 'w:pStyle w:val="TOC' not in blk:
            continue
        lvl = int(re.search(r'w:pStyle w:val="TOC(\d)"', blk).group(1))
        txt = html.unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", blk, re.S)))
        out.append((lvl, txt))
    return out

old, new = toc_lines(a), toc_lines(b)
print(f"\n{'heading':66s} {'old':>5s} {'new':>5s}")
print("-" * 80)
NUM = re.compile(r"([۰-۹]+)$")
for (l1, t1), (l2, t2) in zip(old, new):
    m1, m2 = NUM.search(t1), NUM.search(t2)
    name = NUM.sub("", t1).strip()
    o = m1.group(1) if m1 else "—"
    n = m2.group(1) if m2 else "—"
    flag = "" if o == n else "   <="
    print(f"{'  ' * (l1 - 1) + name[:64]:66s} {o:>5s} {n:>5s}{flag}")
