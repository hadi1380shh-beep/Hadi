import sys, zipfile, re
sys.path.insert(0, "/home/user/Hadi/tools")
import layout as L

z = zipfile.ZipFile(L.DOCX)
body = z.read("word/document.xml").decode("utf-8")
body = body[body.find("<w:body>"):]
spans = [(m.start(), m.end()) for m in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", body, re.S)]
for i in list(range(14, 32)):
    a, b = spans[i]
    blk = body[a:b]
    blk = re.sub(r"<w:proofErr[^>]*/>", "", blk)
    txt = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", blk))
    szs = re.findall(r'<w:sz w:val="(\d+)"', blk)
    sp = re.search(r"<w:spacing[^>]*/>", blk)
    jc = re.search(r'<w:jc w:val="(\w+)"', blk)
    stl = re.search(r'<w:pStyle w:val="([^"]+)"', blk)
    print(f"[{i}] sz={sorted(set(szs))} style={stl.group(1) if stl else '-'} "
          f"{sp.group(0) if sp else ''} jc={jc.group(1) if jc else '-'} :: {txt[:38]}")
