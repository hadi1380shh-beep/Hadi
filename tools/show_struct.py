import sys, zipfile, re
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
paras = L.parse(z)
for i in range(280, 300):
    p = paras[i]
    m = re.search(r'<w:pStyle w:val="([^"]+)"', p["ppr"])
    fl = []
    if p["breaks"]: fl.append("BRK")
    if p["marks"]: fl.append("LRPB" + str(p["marks"]))
    if p["sect"]: fl.append("SECT")
    print(f"[{i:3d}] {m.group(1) if m else '-':11s} {'|'.join(fl):16s} len={len(p['text']):4d} :: {p['text'][:52]}")
print()
for i in range(470, 485):
    p = paras[i]
    m = re.search(r'<w:pStyle w:val="([^"]+)"', p["ppr"])
    fl = []
    if p["breaks"]: fl.append("BRK")
    if p["marks"]: fl.append("LRPB" + str(p["marks"]))
    print(f"[{i:3d}] {m.group(1) if m else '-':11s} {'|'.join(fl):16s} len={len(p['text']):4d} :: {p['text'][:52]}")
