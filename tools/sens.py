import sys, zipfile, re, html
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)
body = z.read("word/document.xml").decode("utf-8")
body = body[body.find("<w:body>"):]

# ---- bookmark order: which bookmark sits at which heading paragraph
bm_order = []
for m in re.finditer(r'<w:bookmarkStart[^>]*w:name="(_Toc\d+)"[^>]*/>', body):
    bm_order.append((m.start(), m.group(1)))
para_spans = [(m.start(), m.end()) for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S)]
bm_para = {}
for pos, name in bm_order:
    for k, (a, b) in enumerate(para_spans):
        if a <= pos < b:
            bm_para[name] = k
            break
print("bookmarks:", len(bm_order), "mapped to paragraphs:", len(bm_para))


def pages_for(w, lh, hs):
    eng = L.Engine(w_scale=w, lh_scale=lh, h_scale=hs)
    res, total, _ = L.layout(paras, st, fns, eng)
    off = res[240][2] - 1
    out = {}
    for name, k in bm_para.items():
        phys = res[k][2]
        out[name] = (phys - off) if k > 239 else phys
    return out, total, off


base, tot, off = pages_for(1.05, 1.30, 0.835)
print("baseline total pages:", tot, "offset:", off)
variants = [(1.00, 1.30, 0.835), (1.10, 1.30, 0.835),
            (1.05, 1.30, 0.82), (1.05, 1.30, 0.85),
            (1.05, 1.25, 0.87), (1.05, 1.35, 0.80),
            (1.02, 1.30, 0.845), (1.08, 1.30, 0.825)]
for w, lh, hs in variants:
    o, t, f = pages_for(w, lh, hs)
    diffs = [abs(o[k] - base[k]) for k in base]
    same = sum(1 for d in diffs if d == 0)
    print(f"  w={w:.2f} lh={lh:.2f} hs={hs:.3f} -> pages={t:3d} offset={f:3d} "
          f"| identical={same}/{len(base)} maxdiff={max(diffs)} meandiff={sum(diffs)/len(diffs):.2f}")
