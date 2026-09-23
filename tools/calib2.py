import sys, zipfile, time
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z)
fns = L.load_footnotes(z)
paras = L.parse(z)
marks_total = sum(len(p["marks"]) for p in paras)


def score(res):
    ok = 0
    for i, p in enumerate(paras):
        if not p["marks"]:
            continue
        lines = res[i][1]
        prev = res[i - 1][1][-1][2] if i and res[i - 1][1] else None
        for m in p["marks"]:
            for k, (a, b, pg) in enumerate(lines):
                if a <= m < b or (m >= len(p["text"]) and k == len(lines) - 1):
                    if k == 0:
                        if prev is None or pg != prev:
                            ok += 1
                    else:
                        if pg != lines[k - 1][2]:
                            ok += 1
                    break
    return ok


def run(w, lh, pad=0.0):
    eng = L.Engine(w_scale=w, lh_scale=lh, para_pad=pad)
    res, pages, _ = L.layout(paras, st, fns, eng)
    return res, pages, score(res)


best = []
t0 = time.time()
for lh in [x / 100 for x in range(100, 191, 5)]:
    for w in [x / 100 for x in range(80, 131, 5)]:
        res, pages, sc = run(w, lh)
        best.append((sc, -abs(pages - 109), w, lh, pages))
best.sort(reverse=True)
print(f"{time.time()-t0:.0f}s  markers={marks_total}")
for b in best[:15]:
    print(f"  match={b[0]:3d}  pages={b[4]:3d}  w={b[2]:.2f} lh={b[3]:.2f}")
