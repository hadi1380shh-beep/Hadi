import sys, zipfile, time
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z); fns = L.load_footnotes(z); paras = L.parse(z)


def evaluate(w, lh, hs):
    eng = L.Engine(w_scale=w, lh_scale=lh, h_scale=hs)
    res, total, _ = L.layout(paras, st, fns, eng)
    ok = tot = 0
    for i, p in enumerate(paras):
        if not p["marks"]:
            continue
        lines = res[i][1]
        prev = res[i - 1][1][-1][2] if i and res[i - 1][1] else None
        for m in p["marks"]:
            tot += 1
            for k, (a, b, pg) in enumerate(lines):
                if a <= m < b or (m >= len(p["text"]) and k == len(lines) - 1):
                    if (k == 0 and (prev is None or pg != prev)) or (k > 0 and pg != lines[k - 1][2]):
                        ok += 1
                    break
    return ok, tot, total, res


best = []
t0 = time.time()
for lh in [x / 100 for x in range(100, 176, 5)]:
    for w in [x / 100 for x in range(90, 121, 2)]:
        ok, tot, total, res = evaluate(w, lh, 1.0)
        best.append((ok, -abs(total - 109), w, lh, total))
best.sort(reverse=True)
print(f"{time.time()-t0:.0f}s")
for b in best[:20]:
    print(f"  match={b[0]:3d}/86 pages={b[4]:3d}  w={b[2]:.2f} lh={b[3]:.2f}")
