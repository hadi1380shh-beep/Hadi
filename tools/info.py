import sys, zipfile, re
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
paras = L.parse(z)
print("paragraphs with explicit page breaks:")
for i, p in enumerate(paras):
    if p["breaks"]:
        m = re.search(r'<w:pStyle w:val="([^"]+)"', p["ppr"])
        print(f"  [{i}] textlen={len(p['text'])} breaks={p['breaks']} style={m.group(1) if m else '-'} :: {p['text'][:40]!r}")
print("pageBreakBefore:", [i for i, p in enumerate(paras) if "<w:pageBreakBefore" in p["ppr"]])
print("sectPr paragraphs:", [i for i, p in enumerate(paras) if p["sect"]])
print("keepNext direct:", sum(1 for p in paras if "<w:keepNext" in p["ppr"]))
print("widowControl off:", sum(1 for p in paras if '<w:widowControl w:val="0"' in p["ppr"]))
