import zipfile, re, html, sys

A = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani (2) (1).docx"
B = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_TOCfixed.docx"
za, zb = zipfile.ZipFile(A), zipfile.ZipFile(B)

print("== package parts ==")
others = [n for n in za.namelist() if n != "word/document.xml"]
bad = [n for n in others if za.getinfo(n).CRC != zb.getinfo(n).CRC]
print(f"  {len(others)} parts other than document.xml, differing: {bad or 'none'}")
print("  same part list/order:", za.namelist() == zb.namelist())

a = za.read("word/document.xml").decode("utf-8")
b = zb.read("word/document.xml").decode("utf-8")

TOCP = re.compile(r'<w:p [^>]*>(?:(?!</w:p>).)*?w:pStyle w:val="TOC\d".*?</w:p>', re.S)
pa, pb = list(TOCP.finditer(a)), list(TOCP.finditer(b))
print("\n== location of the change ==")
print(f"  TOC paragraphs: old {len(pa)}, new {len(pb)}")
print("  bytes before the first TOC paragraph identical:",
      a[:pa[0].start()] == b[:pb[0].start()])
print("  bytes after the last TOC paragraph identical:",
      a[pa[-1].end():] == b[pb[-1].end():])
print(f"  untouched text outside the TOC: {pa[0].start() + len(a) - pa[-1].end():,} of {len(a):,} chars")

DIG = re.compile(r"[۰-۹]+")
NUMRUN = re.compile(r'<w:r><w:rPr><w:noProof/><w:color w:val="[0-9A-F]+"/></w:rPr><w:t>[#۰-۹]+</w:t></w:r>')
blk_a, blk_b = a[pa[0].start():pa[-1].end()], b[pb[0].start():pb[-1].end()]
ma = DIG.sub("#", blk_a); mb = DIG.sub("#", blk_b)
na, nb = len(NUMRUN.findall(ma)), len(NUMRUN.findall(mb))
print(f"  number-bearing runs in the TOC: old {na}, new {nb} (2 lines gained a number)")
print("  every other byte of the TOC block unchanged:",
      NUMRUN.sub("", ma) == NUMRUN.sub("", mb))

print("\n== the corrected table of contents ==")
def lines(x):
    out = []
    for m in TOCP.finditer(x):
        blk = m.group(0)
        lvl = int(re.search(r'w:pStyle w:val="TOC(\d)"', blk).group(1))
        runs = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", blk, re.S)
        txt = html.unescape("".join(runs))
        num = runs[-1] if runs and DIG.fullmatch(html.unescape(runs[-1])) else None
        name = txt[: len(txt) - len(num)] if num else txt
        out.append((lvl, name, num))
    return out

INV = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
prev, mono = -1, True
print(f"  {'':2s}{'heading':66s} {'old':>5s} {'new':>5s}")
old_lines = lines(a)
for (lvl, name, num), (_, _, onum) in zip(lines(b), old_lines):
    v = int(num.translate(INV)) if num else None
    if v is not None and v < prev and not name.startswith("مقدمه"):
        mono = False
    if v is not None:
        prev = v
    print(f"  {'  ' * (lvl - 1)}{name.strip()[:64]:66s} {onum or '—':>5s} {num or '—':>5s}")
print("\n  page numbers non-decreasing down the table:", mono)
