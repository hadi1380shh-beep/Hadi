import sys, zipfile
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L
z=zipfile.ZipFile(L.DOCX); st=L.Styles(z); fns=L.load_footnotes(z); paras=L.parse(z)
W,LH,HS = 1.05,1.30,0.835
eng=L.Engine(w_scale=W,lh_scale=LH,h_scale=HS)
res,pages,fnu=L.layout(paras,st,fns,eng)
print("pages:",pages)
# marker match
tot=ok=0
for i,p in enumerate(paras):
    if not p["marks"]: continue
    lines=res[i][1]; prev=res[i-1][1][-1][2] if i and res[i-1][1] else None
    for m in p["marks"]:
        tot+=1
        for k,(a,b,pg) in enumerate(lines):
            if a<=m<b or (m>=len(p["text"]) and k==len(lines)-1):
                if (k==0 and (prev is None or pg!=prev)) or (k>0 and pg!=lines[k-1][2]): ok+=1
                break
print(f"marker match: {ok}/{tot}")
print("\nfirst 40 paragraphs:")
for i in range(0,40):
    p=paras[i]
    print(f"[{i:3d}] pg={res[i][2]:3d} sect={int(p['sect'])} brk={len(p['breaks'])} lrpb={len(p['marks'])} :: {p['text'][:48]}")
