import sys, zipfile
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
paras = L.parse(z)
print("explicit-break paragraphs and whether the NEXT paragraph starts with an lrpb:")
n_dup = 0
for i, p in enumerate(paras):
    if not p["breaks"]:
        continue
    nxt = paras[i + 1] if i + 1 < len(paras) else None
    dup = bool(nxt and nxt["marks"] and min(nxt["marks"]) == 0)
    n_dup += dup
    print(f"  [{i}] -> next [{i+1}] marks={nxt['marks'][:3] if nxt else None} dup={dup} :: {nxt['text'][:34] if nxt else ''}")
print("duplicates (explicit break + lrpb at same boundary):", n_dup)
print()
print("lrpb at offset 0 total:", sum(1 for p in paras for m in p["marks"] if m == 0))
print("lrpb mid-paragraph total:", sum(1 for p in paras for m in p["marks"] if m > 0))
print("all lrpb:", sum(len(p["marks"]) for p in paras))
print("section breaks:", sum(1 for p in paras if p["sect"]))
print("=> page count if lrpb are all real and explicit breaks are duplicates:",
      1 + sum(len(p["marks"]) for p in paras) + sum(1 for p in paras if p["sect"]))
