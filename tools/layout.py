"""Layout simulator for the thesis DOCX.

Word's `w:lastRenderedPageBreak` markers in the file record where page breaks
fell the last time Word repaginated it.  Content was added after that save, so
some markers are stale -- but the great majority are still valid, which makes
them usable as ground truth to calibrate a simple text-layout model of this
document (single column, RTL, B Nazanin, footnotes at the bottom of the page).

The model is calibrated by grid-searching two parameters (character advance
scale and line-height scale) so that simulated page breaks coincide with the
recorded markers as often as possible.
"""
import re
import html
import zipfile
import math
import unicodedata
from collections import defaultdict

DOCX = "/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani (2) (1).docx"

# ---------------------------------------------------------------- geometry
PAGE_W, PAGE_H = 11906, 16838          # twips (A4)
M_TOP, M_BOT, M_LEFT, M_RIGHT = 1417, 1417, 1134, 1701
TEXT_W_PT = (PAGE_W - M_LEFT - M_RIGHT) / 20.0
TEXT_H_PT = (PAGE_H - M_TOP - M_BOT) / 20.0

# ------------------------------------------------- per-character advance (em)
NARROW = set("اآأإٱ۰٭|دذرزوؤةطلظ١٢٣٤٥٦٧٨٩")
WIDE = set("سشیصضطظمـ")
MID = set("بپتثجچحخعغفقکكگنهءيىئ")


def char_w(ch):
    if ch == "\u200c":          # ZWNJ
        return 0.0
    if ch == " ":
        return 0.24
    if ch in NARROW:
        return 0.30
    if ch in WIDE:
        return 0.62
    if ch in MID:
        return 0.50
    if ch in "،؛:,.-–—!?()«»[]{}\"'٪×+=/\\@#&*<>":
        return 0.28
    if ch.isdigit():
        return 0.50
    if ch.isascii() and ch.isalpha():
        return 0.52
    cat = unicodedata.category(ch)
    if cat in ("Mn", "Cf"):     # combining marks
        return 0.0
    return 0.50


# ------------------------------------------------------------ style helpers
def load_styles(z):
    st = z.read("word/styles.xml").decode("utf-8")
    styles = {}
    dd = re.search(r"<w:docDefaults>.*?</w:docDefaults>", st, re.S).group(0)
    defaults = {
        "sz": int(re.search(r'<w:sz w:val="(\d+)"', dd).group(1)),
        "after": 200, "before": 0, "line": 276, "rule": "auto",
        "keepNext": False, "keepLines": False, "firstLine": 0,
        "left": 0, "right": 0, "jc": None,
    }
    for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', st, re.S):
        sid, blk = m.group(1), m.group(2)
        d = {}
        g = lambda pat: (re.search(pat, blk).group(1) if re.search(pat, blk) else None)
        d["based"] = g(r'<w:basedOn w:val="([^"]+)"')
        d["sz"] = g(r'<w:sz w:val="(\d+)"')
        d["before"] = g(r'<w:spacing[^>]*w:before="(\d+)"')
        d["after"] = g(r'<w:spacing[^>]*w:after="(\d+)"')
        d["line"] = g(r'<w:spacing[^>]*w:line="(\d+)"')
        d["rule"] = g(r'<w:spacing[^>]*w:lineRule="(\w+)"')
        d["keepNext"] = "<w:keepNext" in blk
        d["keepLines"] = "<w:keepLines" in blk
        d["firstLine"] = g(r'<w:ind[^>]*w:firstLine="(-?\d+)"')
        d["left"] = g(r'<w:ind[^>]*w:left="(-?\d+)"')
        d["jc"] = g(r'<w:jc w:val="(\w+)"')
        styles[sid] = d
    return styles, defaults


class Styles:
    def __init__(self, z):
        self.styles, self.defaults = load_styles(z)

    def get(self, sid, key):
        seen = 0
        while sid and sid in self.styles and seen < 10:
            v = self.styles[sid].get(key)
            if v not in (None, False, ""):
                return v
            sid = self.styles[sid].get("based")
            seen += 1
        return self.defaults.get(key)


# ------------------------------------------------------------- footnote info
def load_footnotes(z):
    fn = z.read("word/footnotes.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r'<w:footnote (?:[^>]*?)w:id="(-?\d+)"[^>]*>(.*?)</w:footnote>', fn, re.S):
        fid = int(m.group(1))
        if fid < 1:
            continue
        runs = []
        for r in re.finditer(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", m.group(2), re.S):
            blk = r.group(1)
            txt = "".join(html.unescape(t) for t in
                          re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", blk, re.S))
            if "<w:footnoteRef/>" in blk:
                txt += " "
            if not txt.strip():
                continue
            szm = re.search(r'<w:sz w:val="(\d+)"', blk)
            runs.append((txt, int(szm.group(1)) / 2.0 if szm else 11.0))
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", m.group(2), re.S)
        p = ppr.group(0) if ppr else ""
        before = int(re.search(r'<w:spacing[^>]*w:before="(\d+)"', p).group(1)) if re.search(r'<w:spacing[^>]*w:before="(\d+)"', p) else 0
        after = int(re.search(r'<w:spacing[^>]*w:after="(\d+)"', p).group(1)) if re.search(r'<w:spacing[^>]*w:after="(\d+)"', p) else 0
        line = int(re.search(r'<w:spacing[^>]*w:line="(\d+)"', p).group(1)) if re.search(r'<w:spacing[^>]*w:line="(\d+)"', p) else 240
        out[fid] = dict(runs=runs, before=before, after=after, line=line)
    return out


# ------------------------------------------------------------- body parsing
RUN_TOK = re.compile(
    r"<w:lastRenderedPageBreak\s*/>"
    r"|<w:br[^>]*w:type=\"page\"\s*/>"
    r"|<w:t(?:\s[^>]*)?>.*?</w:t>"
    r"|<w:footnoteReference[^>]*/>"
    r"|<w:tab\s*/>", re.S)

TOK = re.compile(
    r"(?P<ps><w:p(?:\s[^>]*)?>)"
    r"|(?P<pe></w:p>)"
    r"|(?P<pr><w:pPr>.*?</w:pPr>)"
    r"|(?P<rr><w:r(?:\s[^>]*)?>.*?</w:r>)"
    r"|(?P<sect><w:sectPr\b)", re.S)


def parse(z):
    body = z.read("word/document.xml").decode("utf-8")
    body = body[body.find("<w:body>"):]
    paras = []
    stack = []
    for m in TOK.finditer(body):
        k, s = m.lastgroup, m.group(0)
        if k == "ps":
            stack.append(dict(runs=[], fns=[], breaks=[], marks=[], ppr="", sect=False))
        elif k == "pe":
            d = stack.pop()
            d["text"] = "".join(t for t, _ in d["runs"])
            paras.append(d)
        elif k == "pr":
            if stack:
                stack[-1]["ppr"] = s
                if re.search(r"<w:sectPr\b", s):
                    stack[-1]["sect"] = True
        elif k == "rr":
            if not stack:
                continue
            cur = stack[-1]
            szm = re.search(r'<w:sz w:val="(\d+)"', s)
            szc = re.search(r'<w:szCs w:val="(\d+)"', s)
            bold = "<w:b/>" in s or "<w:bCs/>" in s
            sz = int(szc.group(1) if szc else (szm.group(1) if szm else 0)) / 2.0
            pos = len("".join(t for t, _ in cur["runs"]))
            for t in RUN_TOK.finditer(s):
                tk = t.group(0)
                if tk.startswith("<w:lastRenderedPageBreak"):
                    cur["marks"].append(pos)
                elif tk.startswith("<w:br"):
                    cur["breaks"].append(pos)
                elif tk.startswith("<w:footnoteReference"):
                    cur["fns"].append((pos, int(re.search(r'w:id="(-?\d+)"', tk).group(1))))
                elif tk.startswith("<w:tab"):
                    cur["runs"].append(("\t", sz))
                else:
                    cur["runs"].append(
                        (html.unescape(re.sub(r"^<w:t(?:\s[^>]*)?>|</w:t>$", "", tk, flags=re.S)), sz))
        elif k == "sect":
            if stack:
                stack[-1]["sect"] = True
    return paras


# ------------------------------------------------------------------- layout
class Engine:
    def __init__(self, w_scale=1.0, lh_scale=1.0, para_pad=0.0, h_scale=1.0):
        self.w_scale = w_scale
        self.lh_scale = lh_scale
        self.para_pad = para_pad
        self.h_scale = h_scale

    def line_height(self, size, line, rule):
        if rule and rule != "auto":
            return self.h_scale * line / 20.0
        return self.h_scale * self.lh_scale * size * (line / 240.0)

    def text_width(self, txt, size):
        return sum(char_w(c) for c in txt) * size * self.w_scale

    def wrap(self, runs, width_pt, first_indent_pt=0.0):
        """Greedy wrap -> list of lines, each a list of (text,size) chunks with
        the running character offset of the line start."""
        # flatten into (char, size) stream
        chars = []
        for txt, sz in runs:
            if not sz:
                sz = 14.0
            for ch in txt:
                chars.append((ch, sz))
        lines = []
        cur = []
        cur_w = 0.0
        avail = width_pt - first_indent_pt
        start = 0
        i = 0
        n = len(chars)
        while i < n:
            ch, sz = chars[i]
            if ch == "\n":
                lines.append((start, i, cur))
                cur = []
                cur_w = 0.0
                avail = width_pt
                i += 1
                start = i
                continue
            w = char_w(ch) * sz * self.w_scale
            if ch == " " and cur and cur_w + w > avail:
                # line break at space: drop the trailing space
                lines.append((start, i, cur))
                cur = []
                cur_w = 0.0
                avail = width_pt
                i += 1
                start = i
                continue
            if cur_w + w > avail and cur:
                lines.append((start, i, cur))
                cur = []
                cur_w = 0.0
                avail = width_pt
                start = i
            cur.append((ch, sz))
            cur_w += w
            i += 1
        if cur or not lines:
            lines.append((start, n, cur))
        return lines

    def para_metrics(self, p, st):
        ppr = p["ppr"]

        def g(pat, d=None):
            m = re.search(pat, ppr)
            return m.group(1) if m else d

        style = g(r'<w:pStyle w:val="([^"]+)"')
        sz_default = float(st.get(style, "sz") or 28) / 2.0
        before = int(g(r'<w:spacing[^>]*w:before="(\d+)"') or st.get(style, "before") or 0) / 20.0 * self.h_scale
        after = int(g(r'<w:spacing[^>]*w:after="(\d+)"') or st.get(style, "after") or 0) / 20.0 * self.h_scale
        line = int(g(r'<w:spacing[^>]*w:line="(\d+)"') or st.get(style, "line") or 276)
        rule = g(r'<w:spacing[^>]*w:lineRule="(\w+)"') or st.get(style, "rule") or "auto"
        keep_next = "<w:keepNext" in ppr or bool(st.get(style, "keepNext"))
        keep_lines = "<w:keepLines" in ppr or bool(st.get(style, "keepLines"))
        pbb = "<w:pageBreakBefore" in ppr
        fl = g(r'<w:ind[^>]*w:firstLine="(-?\d+)"')
        if fl is None:
            fl = st.get(style, "firstLine") or 0
        first = int(fl) / 20.0
        left = int(g(r'<w:ind[^>]*w:left="(-?\d+)"') or st.get(style, "left") or 0) / 20.0
        right = int(g(r'<w:ind[^>]*w:right="(-?\d+)"') or 0) / 20.0
        runs = [(t, s if s else sz_default) for t, s in p["runs"]]
        return dict(style=style, sz=sz_default, before=before, after=after, line=line,
                    rule=rule, keep_next=keep_next, keep_lines=keep_lines, pbb=pbb,
                    first=first, left=left, right=right, runs=runs)


def layout(paras, st, footnotes, eng, widow=True):
    """Simulate Word pagination.

    Returns (result, total_pages, footnote_heights) where result[i] =
    (para_index, [(line_start, line_end, page), ...], start_page).
    """
    state = {"page": 1, "used": 0.0}
    fn_used = defaultdict(float)
    FN_SEP = 14.0 * eng.h_scale
    result = []

    def page_room():
        fnh = (FN_SEP + fn_used[state["page"]]) if fn_used[state["page"]] else 0.0
        return TEXT_H_PT - fnh - state["used"]

    def new_page():
        state["page"] += 1
        state["used"] = 0.0

    def footnote_height(fid):
        f = footnotes.get(fid)
        if not f:
            return 20.0 * eng.h_scale
        txt = "".join(t for t, _ in f["runs"])
        size = max([s for _, s in f["runs"]], default=11.0)
        n = max(1, math.ceil(eng.text_width(txt, size) / TEXT_W_PT)) if txt.strip() else 1
        return (n * eng.line_height(size, f["line"], "auto")
                + (f["before"] + f["after"]) / 20.0 * eng.h_scale)

    def block_h(chunk, pm):
        s = max([sz for _, sz in chunk], default=pm["sz"]) if chunk else pm["sz"]
        return eng.line_height(s, pm["line"], pm["rule"])

    # pre-compute metrics/blocks per paragraph
    info = []
    for p in paras:
        pm = eng.para_metrics(p, st)
        width = TEXT_W_PT - pm["left"] - pm["right"]
        if pm["runs"]:
            blocks = eng.wrap(pm["runs"], width, pm["first"])
        else:
            blocks = [(0, 0, [])]
        info.append((pm, blocks))

    def first_line_cost(idx):
        """height needed to start paragraph idx on a page (before + 1 line)"""
        if idx >= len(paras):
            return 0.0
        pm, blocks = info[idx]
        return pm["before"] + block_h(blocks[0][2], pm)

    for idx, p in enumerate(paras):
        pm, blocks = info[idx]
        fns_by_pos = {}
        for pos, fid in p["fns"]:
            fns_by_pos.setdefault(pos, []).append(fid)

        # ---- forced page break paragraph (empty paragraph holding a break)
        if p["breaks"] and not pm["runs"]:
            h = block_h([], pm) + pm["before"] + pm["after"]
            if h > page_room():
                new_page()
            state["used"] += h
            result.append((idx, [(0, 0, state["page"])], state["page"]))
            for _ in set(p["breaks"]):
                new_page()
            if p["sect"]:
                new_page()
            continue

        placed = []
        li = 0
        n = len(blocks)
        first_on_page = True
        retried = False
        while li < n:
            room = page_room()
            if first_on_page:
                room -= pm["before"]
            fit = 0
            hh = 0.0
            j = li
            while j < n:
                bh = block_h(blocks[j][2], pm)
                if hh + bh > room + 1e-6:
                    break
                hh += bh
                fit += 1
                j += 1
            if fit == 0:
                new_page()
                first_on_page = True
                continue
            # keepLines: whole paragraph must stay together
            if pm["keep_lines"] and fit < n - li:
                new_page()
                first_on_page = True
                continue
            # widow/orphan control: never leave a single line alone
            if widow and n > 1 and not retried:
                if fit == 1 and li == 0:
                    retried = True
                    new_page()
                    first_on_page = True
                    continue
                if li > 0 and n - li == 1 and fit >= 2:
                    fit -= 1
                    hh -= block_h(blocks[li + fit][2], pm)
            # keepNext: last line of this paragraph must share a page with the
            # first line of the next paragraph
            if (pm["keep_next"] and li + fit == n and idx + 1 < len(paras)
                    and li == 0 and not retried):
                after = room - hh
                if after < first_line_cost(idx + 1):
                    retried = True
                    new_page()
                    first_on_page = True
                    continue
            # commit
            start_li = li
            for k in range(fit):
                a, b, chunk = blocks[li + k]
                placed.append((a, b, state["page"]))
            state["used"] += hh + (pm["before"] if first_on_page else 0.0)
            first_on_page = False
            li += fit
            # footnotes referenced by the lines just placed
            for k in range(start_li, li):
                a, b, _ = blocks[k]
                for pos, fids in fns_by_pos.items():
                    if a <= pos < b or (pos == b == len(p["text"])):
                        for fid in fids:
                            fn_used[state["page"]] += footnote_height(fid)
            # footnote area may have overflowed: roll lines back
            guard = 0
            while page_room() < -1e-6 and li > start_li and guard < 50:
                guard += 1
                li -= 1
                placed.pop()
                state["used"] -= block_h(blocks[li][2], pm)
            if li < n:
                new_page()
                first_on_page = True

        if not placed:
            placed = [(0, len(p["text"]), state["page"])]
        state["used"] += pm["after"] + eng.para_pad
        result.append((idx, placed, placed[0][2]))
        if p["sect"]:
            new_page()
    return result, state["page"], fn_used
