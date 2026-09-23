import sys, zipfile
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)

W, LH, HS = 1.05, 1.30, 0.835
eng = L.Engine(w_scale=W, lh_scale=LH, h_scale=HS)
res, pages, fn_used = L.layout(paras, st, fns, eng)

# marker-implied page = 1 + number of page breaks before it (lrpb + explicit + section)
impl = 1
rows = []
for i, p in enumerate(paras):
    evts = sorted([(m, "M") for m in p["marks"]] + [(b, "B") for b in p["breaks"]])
    for pos, kind in evts:
        if kind == "M":
            impl += 1
            sim = None
            lines = res[i][1]
            for k, (a, b, pg) in enumerate(lines):
                if a <= pos < b or (pos >= len(p["text"]) and k == len(lines) - 1):
                    sim = pg; break
            rows.append((i, pos, impl, sim, p["text"][:34]))
        else:
            impl += 1
    if p["sect"]:
        impl += 1
print(f"total pages simulated={pages}")
print(f"{'para':>5} {'marker_pg':>9} {'sim_pg':>6} {'diff':>5}  text")
for i, pos, impl_pg, sim, txt in rows:
    d = (sim - impl_pg) if sim else None
    print(f"{i:5d} {impl_pg:9d} {str(sim):>6} {str(d):>5}  {txt}")
