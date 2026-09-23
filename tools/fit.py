import sys, zipfile, statistics
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L
import math

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)

W, LH = 1.05, 1.30
eng = L.Engine(w_scale=W, lh_scale=LH)

def para_height(p):
    pm = eng.para_metrics(p, st)
    if not pm["runs"]:
        return eng.line_height(pm["sz"], pm["line"], pm["rule"]) + pm["before"] + pm["after"]
    width = L.TEXT_W_PT - pm["left"] - pm["right"]
    lines = eng.wrap(pm["runs"], width, pm["first"])
    h = 0.0
    for a, b, chunk in lines:
        s = max([sz for _, sz in chunk], default=pm["sz"]) if chunk else pm["sz"]
        h += eng.line_height(s, pm["line"], pm["rule"])
    return h + pm["before"] + pm["after"]

def fn_height(fid):
    f = fns.get(fid)
    if not f: return 20.0
    txt = "".join(t for t, _ in f["runs"])
    size = max([s for _, s in f["runs"]], default=11.0)
    n = max(1, math.ceil(eng.text_width(txt, size) / L.TEXT_W_PT)) if txt.strip() else 1
    return n * eng.line_height(size, f["line"], "auto") + (f["before"] + f["after"]) / 20.0 + 14.0

# cumulative height including footnote areas
cum = 0.0
ratios = []
impl = 1
for i, p in enumerate(paras):
    h = para_height(p)
    for pos, fid in p["fns"]:
        h += fn_height(fid)
    evts = sorted([(m, "M") for m in p["marks"]] + [(b, "B") for b in p["breaks"]])
    for pos, kind in evts:
        if kind == "M":
            impl += 1
            ratios.append(((impl - 1) * L.TEXT_H_PT) / cum if cum > 0 else None)
    cum += h
    if p["sect"]:
        impl += 1

ratios = [r for r in ratios if r]
print("n ratios:", len(ratios))
print("median:", round(statistics.median(ratios), 4), " mean:", round(statistics.mean(ratios), 4))
print("min:", round(min(ratios), 4), "max:", round(max(ratios), 4))
qs = statistics.quantiles(ratios, n=10)
print("deciles:", [round(q, 3) for q in qs])
print("\nratios in order:")
print(" ".join(f"{r:.3f}" for r in ratios))
