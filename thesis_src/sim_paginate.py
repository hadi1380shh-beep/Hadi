# -*- coding: utf-8 -*-
"""شبیه‌ساز صفحه‌بندی Word با متریک واقعی فونت‌های سری B.

- گام سطر (pitch): از جدول usWin فونت (همان که Word برای line-spacing خودکار به‌کار می‌برد)
- پهنای سطر: از advance واقعی تک‌تک حروف متن با همان فونت
اعتبارسنجی: بیلد قدیمی باید دقیقاً ۸۱ صفحه (مشاهده کاربر در Word) بدهد.
"""
import json
import math
import os
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

ADV_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'font_adv.json')
WIDTH_SCALE = 1.0  # ضریب چگالی کالیبره‌شده با مشاهده واقعی کاربر (۸۱ صفحه)
PITCH = {
    ('B Lotus', False): 1.0,
    ('B Lotus', True): 1.0,
    ('B Badr', False): 1.0,
    ('B Badr', True): 1.0,
    ('B Titr', False): 1.0,
    ('B Titr', True): 1.0,
    ('B Traffic', False): 1.0,
    ('B Traffic', True): 1.0,
}
ZERO_WIDTH = set('\u200c\u200d\u200e\u200f\u202a\u202b\u202c\u202d\u202e\ufeff\xad')

_ADV = None


def advem(family, bold, ch):
    global _ADV
    if _ADV is None:
        with open(ADV_JSON, encoding='utf-8') as f:
            _ADV = json.load(f)
    tab = _ADV.get(f'{family}|{int(bold)}') or _ADV.get('B Lotus|0')
    v = tab.get(ch)
    if v is None:
        v = 0.0 if (ch in ZERO_WIDTH or unicodedata.category(ch) in ('Mn', 'Me', 'Cf')) else 0.5
    return v


def _f(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def run_font_bold(r):
    fam, bold = 'B Lotus', False
    rPr = r.find(f'{W}rPr')
    if rPr is not None:
        rf = rPr.find(f'{W}rFonts')
        if rf is not None:
            fam = rf.get(f'{W}ascii') or rf.get(f'{W}cs') or fam
        b = rPr.find(f'{W}b')
        if b is not None and (b.get(f'{W}val') or '1') not in ('0', 'false'):
            bold = True
        else:
            bcs = rPr.find(f'{W}bCs')
            if bcs is not None and (bcs.get(f'{W}val') or '1') not in ('0', 'false'):
                bold = True
    return fam, bold


def parse_docx(path):
    zf = zipfile.ZipFile(path)
    dx = ET.fromstring(zf.read('word/document.xml').decode('utf8'))
    st = ET.fromstring(zf.read('word/styles.xml').decode('utf8'))

    normal = {'size': 14.0, 'before': 0.0, 'after': 0.0, 'line': 240.0, 'rule': 'auto'}
    for s in st.findall(f'{W}style'):
        if s.get(f'{W}styleId') != 'Normal':
            continue
        pPr = s.find(f'{W}pPr')
        if pPr is not None:
            sp = pPr.find(f'{W}spacing')
            if sp is not None:
                normal['before'] = _f(sp.get(f'{W}before')) / 20.0
                normal['after'] = _f(sp.get(f'{W}after')) / 20.0
                normal['line'] = _f(sp.get(f'{W}line'), 240.0)
                normal['rule'] = sp.get(f'{W}lineRule') or 'auto'
        rPr = s.find(f'{W}rPr')
        if rPr is not None:
            sz = rPr.find(f'{W}sz')
            if sz is not None:
                normal['size'] = _f(sz.get(f'{W}val'), 28.0) / 2.0

    body = dx.find(f'{W}body')
    paras = []
    for p in body.findall(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        sp = pPr.find(f'{W}spacing') if pPr is not None else None
        ind = pPr.find(f'{W}ind') if pPr is not None else None
        size = None
        advEm = 0.0
        has_text = False
        pitch = PITCH[('B Lotus', False)]
        fnrefs = 0
        brk = False
        for r in p.findall(f'{W}r'):
            fam, bold = run_font_bold(r)
            pitch = max(pitch, PITCH.get((fam, bold), PITCH[('B Lotus', False)]))
            if size is None:
                rPr = r.find(f'{W}rPr')
                if rPr is not None:
                    e = rPr.find(f'{W}sz')
                    if e is not None:
                        size = _f(e.get(f'{W}val'), 28.0) / 2.0
                    else:
                        e = rPr.find(f'{W}szCs')
                        if e is not None:
                            size = _f(e.get(f'{W}val'), 28.0) / 2.0
            for t in r.findall(f'{W}t'):
                tx = t.text or ''
                if tx:
                    has_text = True
                for ch in tx:
                    advEm += advem(fam, bold, ch)
            advEm += 0.5 * len(r.findall(f'{W}tab'))
            for b in r.findall(f'{W}br'):
                if (b.get(f'{W}type') or 'textWrapping') == 'page':
                    brk = True
            fnrefs += len(r.findall(f'{W}footnoteReference'))
        if size is None:
            size = normal['size']
        if sp is not None:
            before = _f(sp.get(f'{W}before'), normal['before'] * 20.0) / 20.0
            after = _f(sp.get(f'{W}after'), normal['after'] * 20.0) / 20.0
            line = _f(sp.get(f'{W}line'), normal['line'])
            rule = sp.get(f'{W}lineRule') or normal['rule']
        else:
            before, after, line, rule = normal['before'], normal['after'], normal['line'], normal['rule']
        if rule == 'auto':
            lineH = size * pitch * (line / 240.0)
        else:
            lineH = line / 20.0
        indent = 0.0
        if ind is not None:
            indent = (_f(ind.get(f'{W}left')) + _f(ind.get(f'{W}right'))) / 20.0
        keepNext = pPr is not None and pPr.find(f'{W}keepNext') is not None
        bms = [b.get(f'{W}name') for b in p.findall(f'{W}bookmarkStart')]
        sect = pPr.find(f'{W}sectPr') if pPr is not None else None
        sty = ''
        if pPr is not None:
            ps = pPr.find(f'{W}pStyle')
            if ps is not None:
                sty = ps.get(f'{W}val') or ''
        paras.append({
            'advEm': advEm, 'has_text': has_text, 'fnrefs': fnrefs, 'break': brk,
            'size': size, 'before': before, 'after': after, 'lineH': lineH,
            'indent': indent, 'keepNext': keepNext, 'bookmarks': bms,
            'sect': sect, 'style': sty,
        })

    sectPrs = [pa['sect'] for pa in paras if pa['sect'] is not None]
    final_sect = body.find(f'{W}sectPr')
    if final_sect is not None:
        sectPrs.append(final_sect)
    g = sectPrs[0]
    pgW = _f(g.find(f'{W}pgSz').get(f'{W}w')) / 20.0
    pgH = _f(g.find(f'{W}pgSz').get(f'{W}h')) / 20.0
    mar = g.find(f'{W}pgMar')
    geom = {
        'width': pgW - (_f(mar.get(f'{W}left')) + _f(mar.get(f'{W}right'))) / 20.0,
        'height': pgH - (_f(mar.get(f'{W}top')) + _f(mar.get(f'{W}bottom'))) / 20.0,
    }
    sections = []
    for s in sectPrs:
        pg = s.find(f'{W}pgNumType')
        stype = s.find(f'{W}type')
        sections.append({
            'pgStart': int(pg.get(f'{W}start')) if pg is not None and pg.get(f'{W}start') else None,
            'type': (stype.get(f'{W}val') if stype is not None else 'nextPage'),
        })

    # ارتفاع دقیق هر پاورقی: اندازه واقعی متن یادداشت
    note_heights = []
    try:
        fnx = ET.fromstring(zf.read('word/footnotes.xml').decode('utf8'))
        for fn in fnx.findall(f'{W}footnote'):
            fid = fn.get(f'{W}id')
            if fid in ('0', '1'):
                continue
            adv = 0.0
            for t in fn.findall(f'.//{W}t'):
                for ch in (t.text or ''):
                    adv += advem('B Lotus', False, ch)
            nlines = max(1, math.ceil(adv * 10.0 / (geom['width'] * WIDTH_SCALE_REF[0])))
            note_heights.append(nlines * 10.0 * PITCH[('B Lotus', False)] + 2.0)
    except KeyError:
        pass
    zf.close()
    return paras, geom, normal, sections, note_heights


def simulate(path, wscale=None):
    if wscale is None:
        wscale = WIDTH_SCALE
    paras, geom, normal, sections, note_heights = parse_docx(path)
    pageH = geom['height']
    baseW = geom['width']

    # تخصیص ارتفاع یادداشت‌ها به ارجاع‌ها به ترتیب سند
    ref_ptr = [0]

    def para_height(pa, consume):
        width = (baseW - pa['indent']) * wscale
        if pa['has_text']:
            lines = max(1, math.ceil(pa['advEm'] * pa['size'] / width - 1e-9))
        else:
            lines = 1
        H = pa['before'] + lines * pa['lineH'] + pa['after']
        if pa['fnrefs']:
            for _ in range(pa['fnrefs']):
                if ref_ptr[0] < len(note_heights):
                    H += note_heights[ref_ptr[0]]
                else:
                    H += 20.0
                if consume:
                    ref_ptr[0] += 1
            H += 20.0  # خط جداکننده پاورقی
            lines = (H - pa['before'] - pa['after']) / pa['lineH']
        return H, lines

    phys, disp, y, seci = 1, 1, 0.0, 0
    mapping = {}

    def new_page():
        nonlocal phys, disp, y
        phys += 1
        disp += 1
        y = 0.0

    i, n = 0, len(paras)
    while i < n:
        pa = paras[i]
        if pa['break']:
            new_page()
        else:
            H, lines = para_height(pa, consume=False)
            if pa['keepNext'] and y > 0:
                need = H
                j = i + 1
                while j < n and not paras[j]['break']:
                    q = paras[j]
                    qH, qLines = para_height(q, consume=False)
                    if q['keepNext']:
                        need += qH
                        j += 1
                        continue
                    need += q['before'] + min(2.0, qLines) * q['lineH']
                    break
                if y + need > pageH:
                    new_page()
            for nm in pa['bookmarks']:
                mapping[nm] = disp
            H, lines = para_height(pa, consume=True)
            remH = lines * pa['lineH']
            first = True
            while True:
                avail = pageH - y
                cb = pa['before'] if first else 0.0
                if cb + remH + pa['after'] <= avail + 1e-6:
                    y += cb + remH + pa['after']
                    break
                fit = math.floor((avail - cb) / pa['lineH'] + 1e-9)
                remInt = math.ceil(remH / pa['lineH'] - 1e-9)
                if first and fit < 2 and remInt > 1 and y > 0:
                    new_page()
                    continue
                if fit < 1:
                    if y > 0:
                        new_page()
                        continue
                    fit = 1
                leftInt = remInt - fit
                if leftInt == 1 and fit > 1:
                    fit -= 1
                y += cb + fit * pa['lineH']
                remH -= fit * pa['lineH']
                new_page()
                first = False
        if pa['sect'] is not None:
            seci += 1
            assert seci < len(sections), 'section overflow'
            phys += 1
            y = 0.0
            st = sections[seci]['pgStart']
            disp = st if st is not None else disp + 1
        i += 1
    return {
        'total_phys': phys, 'mapping': mapping, 'sections': len(sections),
        'geom': geom, 'paras': n, 'refs_used': ref_ptr[0], 'notes': len(note_heights),
    }


WIDTH_SCALE_REF = [1.0]

if __name__ == '__main__':
    ws = float(sys.argv[2]) if len(sys.argv) > 2 else 1.0
    WIDTH_SCALE_REF[0] = ws
    r = simulate(sys.argv[1], ws)
    print(f"total_phys={r['total_phys']} sections={r['sections']} paras={r['paras']} "
          f"bookmarks={len(r['mapping'])} refs_used={r['refs_used']}/{r['notes']}")
