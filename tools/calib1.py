import sys, zipfile, time, itertools
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
st = L.Styles(z)
fns = L.load_footnotes(z)
paras = L.parse(z)
print("paragraphs:", len(paras), "footnotes:", len(fns))
marks = [(i, m) for i, p in enumerate(paras) for m in p["marks"]]
print("markers:", len(marks))


def score(res):
    ok = 0
    for i, p in enumerate(paras):
        if not p["marks"]:
            continue
        lines = res[i][1]
        prev_page = res[i - 1][1][-1][2] if i and res[i - 1][1] else None
        for m in p["marks"]:
            hit = False
            for k, (a, b, pg) in enumerate(lines):
                if a <= m < b or (m == len(p["text"]) and k == len(lines) - 1):
                    if k == 0:
                        hit = (prev_page is None or pg != prev_page)
                    else:
                        hit = (pg != lines[k - 1][2])
                    break
            if hit:
                ok += 1
    return ok


def run(w, lh, pad=0.0):
    eng = L.Engine(w_scale=w, lh_scale=lh, para_pad=pad)
    res, pages, _ = L.layout(paras, st, fns, eng)
    return res, pages, score(res)


t0 = time.time()
res, pages, sc = run(1.0, 1.6)
print(f"baseline w=1.00 lh=1.60 -> pages={pages} markers matched={sc}/{len(marks)}  ({time.time()-t0:.1f}s)")
for i, p in enumerate(paras):
    if p["text"].strip() in ("چکیده", "پیشگفتار", "فهرست مطالب", "مقدمه: کلیات تحقیق") \
       or p["text"].startswith("بند چهارم: تحلیل اصل ۱۱۰") \
       or p["text"].startswith("گفتار چهارم: اسباب رافع، آثار"):
        print(f"   pg {res[i][2]:3d}  {p['text'][:50]}")
