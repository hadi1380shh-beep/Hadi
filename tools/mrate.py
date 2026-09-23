import sys, zipfile
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)


def marker_report(w, lh, hs):
    eng = L.Engine(w_scale=w, lh_scale=lh, h_scale=hs)
    res, total, _ = L.layout(paras, st, fns, eng)
    ok = tot = 0
    misses = []
    for i, p in enumerate(paras):
        if not p["marks"]:
            continue
        lines = res[i][1]
        prev = res[i - 1][1][-1][2] if i and res[i - 1][1] else None
        for m in p["marks"]:
            tot += 1
            hit = False
            for k, (a, b, pg) in enumerate(lines):
                if a <= m < b or (m >= len(p["text"]) and k == len(lines) - 1):
                    hit = (k == 0 and (prev is None or pg != prev)) or (k > 0 and pg != lines[k - 1][2])
                    break
            if hit:
                ok += 1
            else:
                misses.append(i)
    return ok, tot, total, misses


for hs in [0.820, 0.825, 0.830, 0.835, 0.840, 0.845, 0.850]:
    ok, tot, total, misses = marker_report(1.05, 1.30, hs)
    print(f"hs={hs:.3f} -> pages={total:3d}  markers matched {ok}/{tot}")
ok, tot, total, misses = marker_report(1.05, 1.30, 0.835)
print("\nmissed marker paragraphs (hs=0.835):", misses)
for i in misses:
    print(f"   [{i}] simpg={0} :: {paras[i]['text'][:60]}")
