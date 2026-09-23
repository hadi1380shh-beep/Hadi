import sys, zipfile, re
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
paras = L.parse(z)
st = L.Styles(z)
eng = L.Engine(w_scale=1.05, lh_scale=1.30, h_scale=0.835)
for i in list(range(14, 60)):
    p = paras[i]
    pm = eng.para_metrics(p, st)
    blocks = eng.wrap(pm["runs"], L.TEXT_W_PT - pm["left"] - pm["right"], pm["first"]) if pm["runs"] else []
    h = sum(eng.line_height(max([s for _, s in c], default=pm["sz"]) if c else pm["sz"], pm["line"], pm["rule"]) for a, b, c in blocks) + pm["before"] + pm["after"]
    print(f"[{i:3d}] sz={pm['sz']:5.1f} line={pm['line']:4d} rule={pm['rule']:6s} b={pm['before']:5.1f} a={pm['after']:5.1f} lines={len(blocks):3d} h={h:6.1f} :: {p['text'][:40]}")
