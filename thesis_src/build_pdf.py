# -*- coding: utf-8 -*-
"""ساخت PDF رساله از روی docx نهایی — با فونت‌های اصلی B و صفحه‌بندی مستقل و خودسازگار.

روش: تجزیه docx ← شکستن حریصانه سطرها با advance واقعی حروف ← چیدمان صفحه
(مشابه Word: keep-with-next، بیوه/یتیم، پاورقی، شروع مجدد شماره) ← رندر reportlab.
شماره‌های فهرست از چیدمان همین PDF می‌آید پس دقیق است.
نیازمندی‌ها (در هر نشست تازه نصب شود): reportlab arabic_reshaper python-bidi pypdf matplotlib fonttools
اجرا: python3 build_pdf.py
"""
import glob
import math
import os
import re
import subprocess
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sim_paginate as S

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

DOCX = '/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_Final2.docx'
PDF = '/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_Final2.pdf'
KAJ = '/tmp/kajfonts/plugins/bfonts/resources'
FA_DIGITS = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')
FA_RE = re.compile(r'[\u0600-\u06FF]')
TOC_STYLE_INDENT = {'TOC1': 0.0, 'TOC2': 284.0 / 20.0, 'TOC3': 567.0 / 20.0, 'TOC4': 851.0 / 20.0}


def ensure_bfonts():
    if not os.path.exists(f'{KAJ}/BLotus.ttf'):
        subprocess.run(['rm', '-rf', '/tmp/kajfonts'], check=False)
        subprocess.run(['git', 'clone', '-q', '--depth', '1', '--filter=blob:none',
                        '--sparse', 'https://github.com/HamedMasafi/Kaj.git',
                        '/tmp/kajfonts'], check=True)
        subprocess.run(['git', '-C', '/tmp/kajfonts', 'sparse-checkout', 'set',
                        'plugins/bfonts/resources'], check=True)
    return {
        ('B Lotus', False): f'{KAJ}/BLotus.ttf',
        ('B Lotus', True): f'{KAJ}/BLotusBd.ttf',
        ('B Badr', False): f'{KAJ}/BBadr.ttf',
        ('B Badr', True): f'{KAJ}/BBadrBd.ttf',
        ('B Titr', False): f'{KAJ}/BTitrBd.ttf',
        ('B Titr', True): f'{KAJ}/BTitrBd.ttf',
        ('B Traffic', False): f'{KAJ}/BTraffic.ttf',
        ('B Traffic', True): f'{KAJ}/BTraffic.ttf',
    }


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


def parse_paras(p_list):
    paras = []
    for p in p_list:
        pPr = p.find(f'{W}pPr')
        sty0 = ''
        if pPr is not None:
            _ps = pPr.find(f'{W}pStyle')
            if _ps is not None:
                sty0 = _ps.get(f'{W}val') or ''
        is_toc, toc_cut = sty0.startswith('TOC'), False
        runs, refs, instrs = [], [], []
        text_len, size_first, brk = 0, None, False
        for r in p.findall(f'{W}r'):
            if is_toc and toc_cut:
                it0 = r.find(f'{W}instrText')
                if it0 is not None and it0.text:
                    instrs.append(it0.text.strip())
                continue
            fam, bold = run_font_bold(r)
            sz = None
            rPr = r.find(f'{W}rPr')
            if rPr is not None:
                e = rPr.find(f'{W}sz')
                if e is not None:
                    sz = _f(e.get(f'{W}val'), 28.0) / 2.0
                else:
                    e = rPr.find(f'{W}szCs')
                    if e is not None:
                        sz = _f(e.get(f'{W}val'), 28.0) / 2.0
                va = rPr.find(f'{W}vertAlign')
                sup = va is not None and (va.get(f'{W}val') or '') == 'superscript'
            else:
                sup = False
            if sz is None:
                sz = 14.0
            if size_first is None:
                size_first = sz
            it = r.find(f'{W}instrText')
            if it is not None and it.text:
                instrs.append(it.text.strip())
            tx = ''.join(t.text or '' for t in r.findall(f'{W}t'))
            ntab = len(r.findall(f'{W}tab'))
            if is_toc and ntab and not toc_cut:
                if tx:
                    runs.append((fam, bold, sz, tx, sup))
                    text_len += len(tx)
                toc_cut = True
                refs += [text_len] * 0
            elif tx or ntab:
                if ntab:
                    tx += '\t' * ntab
                runs.append((fam, bold, sz, tx, sup))
                text_len += len(tx)
            for b in r.findall(f'{W}br'):
                if (b.get(f'{W}type') or 'textWrapping') == 'page':
                    brk = True
            refs += [text_len] * len(r.findall(f'{W}footnoteReference'))
        if pPr is not None:
            sp = pPr.find(f'{W}spacing')
            ind = pPr.find(f'{W}ind')
            if sp is not None:
                before = _f(sp.get(f'{W}before')) / 20.0
                after = _f(sp.get(f'{W}after')) / 20.0
                line = _f(sp.get(f'{W}line'), 240.0)
                rule = sp.get(f'{W}lineRule') or 'auto'
            else:
                before, after, line, rule = 0.0, 0.0, 240.0, 'auto'
            indent_r = _f(ind.get(f'{W}right')) / 20.0 if ind is not None else 0.0
            indent_l = _f(ind.get(f'{W}left')) / 20.0 if ind is not None else 0.0
            indent_f = _f(ind.get(f'{W}firstLine')) / 20.0 if ind is not None else 0.0
            keepNext = pPr.find(f'{W}keepNext') is not None
            sty = ''
            ps = pPr.find(f'{W}pStyle')
            if ps is not None:
                sty = ps.get(f'{W}val') or ''
            jc_el = pPr.find(f'{W}jc')
            jc = jc_el.get(f'{W}val') if jc_el is not None else 'right'
            shd = pPr.find(f'{W}shd')
            shd_fill = shd.get(f'{W}fill') if shd is not None else None
            pbdr = pPr.find(f'{W}pBdr')
            bdr = {}
            if pbdr is not None:
                for side in ('top', 'left', 'bottom', 'right'):
                    e = pbdr.find(f'{W}{side}')
                    if e is not None:
                        bdr[side] = (e.get(f'{W}color') or '000000',
                                     _f(e.get(f'{W}sz'), 6.0) / 8.0)
            sect = pPr.find(f'{W}sectPr')
        else:
            before = after = 0.0
            line, rule = 240.0, 'auto'
            indent_r = indent_l = indent_f = 0.0
            keepNext, sty, sect = False, '', None
            jc, shd_fill, bdr = 'right', None, {}
        pitch = S.PITCH[('B Lotus', False)]
        for fam, bold, _sz, _tx, _sup in runs:
            pitch = max(pitch, S.PITCH.get((fam, bold), pitch))
        size = size_first or 14.0
        lineH = size * pitch * (line / 240.0) if rule == 'auto' else line / 20.0
        bms = [b.get(f'{W}name') for b in p.findall(f'{W}bookmarkStart')]
        paras.append({
            'runs': runs, 'refs': refs, 'instrs': instrs, 'break': brk,
            'size': size, 'before': before, 'after': after, 'lineH': lineH,
            'pitch': pitch, 'indent_r': indent_r, 'indent_l': indent_l,
            'indent_f': indent_f, 'keepNext': keepNext, 'bookmarks': bms,
            'sect': sect, 'style': sty, 'jc': jc, 'shd': shd_fill, 'bdr': bdr,
            'text': ''.join(tx for _, _, _, tx, _ in runs),
        })
    return paras


def run_adv_list(run):
    fam, bold, _sz, tx, _sup = run
    return [S.advem(fam, bold, ch) if ch != '\t' else 0.5 for ch in tx]


def greedy_wrap(pa, width_pt):
    text = pa['text'].replace('\t', ' ')
    if not text.strip():
        return [(0, len(pa['text']))]
    advs = []
    for run in pa['runs']:
        advs.extend(run_adv_list(run))
    n = len(pa['text'])
    cap = width_pt / pa['size']
    lines, i = [], 0
    while i < n:
        first_cap = cap - (pa['indent_f'] / pa['size'] if not lines else 0.0)
        acc, j, last_sp = 0.0, i, -1
        while j < n:
            a = advs[j] if j < len(advs) else 0.3
            if acc + a > first_cap and j > i:
                break
            acc += a
            if text[j] == ' ':
                last_sp = j
            j += 1
        if j < n and last_sp > i:
            j = last_sp + 1
        elif j == i:
            j = i + 1
        lines.append((i, min(j, n)))
        i = min(j, n)
        while i < n and text[i] == ' ':
            i += 1
    return lines


def layout_document(docx_path):
    zf = zipfile.ZipFile(docx_path)
    dx = ET.fromstring(zf.read('word/document.xml').decode('utf8'))
    body = dx.find(f'{W}body')
    paras = parse_paras(body.findall(f'{W}p'))
    g = body.find(f'{W}sectPr')
    pgW = _f(g.find(f'{W}pgSz').get(f'{W}w')) / 20.0
    pgH = _f(g.find(f'{W}pgSz').get(f'{W}h')) / 20.0
    mar = g.find(f'{W}pgMar')
    marL = _f(mar.get(f'{W}left')) / 20.0
    marR = _f(mar.get(f'{W}right')) / 20.0
    marT = _f(mar.get(f'{W}top')) / 20.0
    marB = _f(mar.get(f'{W}bottom')) / 20.0
    marHdr = _f(mar.get(f'{W}header'), 720.0) / 20.0
    marFtr = _f(mar.get(f'{W}footer'), 720.0) / 20.0
    baseW = pgW - marL - marR
    pageH = pgH - marT - marB
    # بخش‌ها به ترتیب
    break_spots = [k for k, pa in enumerate(paras) if pa['sect'] is not None]
    sect_nodes = [paras[k]['sect'] for k in break_spots] + [g]
    sect_list = []
    for s in sect_nodes:
        pg = s.find(f'{W}pgNumType')
        d = {'pgStart': int(pg.get(f'{W}start')) if pg is not None and pg.get(f'{W}start') else None,
             'hdr': None, 'ftr': None, 'border': None}
        for tag, key in (('headerReference', 'hdr'), ('footerReference', 'ftr')):
            e = s.find(f'{W}{tag}')
            if e is not None:
                d[key] = e.get(R_ID)
        pb = s.find(f'{W}pgBorders')
        if pb is not None:
            top = pb.find(f'{W}top')
            if top is not None:
                d['border'] = (top.get(f'{W}color') or '000000', _f(top.get(f'{W}sz'), 6.0) / 8.0)
        sect_list.append(d)
    # نگاشت rId به پارت سرصفحه/پاصفحه
    rels = ET.fromstring(zf.read('word/_rels/document.xml.rels').decode('utf8'))
    rid2part = {}
    for rel in rels:
        t = rel.get('Type') or ''
        if 'header' in t or 'footer' in t:
            rid2part[rel.get('Id')] = 'word/' + (rel.get('Target') or '').split('/')[-1]
    # یادداشت‌ها
    notes = []
    try:
        fnx = ET.fromstring(zf.read('word/footnotes.xml').decode('utf8'))
        for fn in fnx.findall(f'{W}footnote'):
            if fn.get(f'{W}id') in ('0', '1'):
                continue
            notes.append(''.join(t.text or '' for t in fn.findall(f'.//{W}t')))
    except KeyError:
        pass
    zf.close()

    def eff_indent_r(pa):
        return pa['indent_r'] + TOC_STYLE_INDENT.get(pa['style'], 0.0)

    widths = [baseW - eff_indent_r(pa) - pa['indent_l'] for pa in paras]
    wraps = [greedy_wrap(pa, widths[k]) if not pa['break'] else [] for k, pa in enumerate(paras)]

    note_lineH = 10.0 * S.PITCH[('B Lotus', False)]
    note_wraps, note_heights = [], []
    for tx in notes:
        cap = baseW / 10.0
        words, lines = tx.split(' '), []
        cur_l, acc = '', 0.0
        for w_ in words:
            wa = sum(S.advem('B Lotus', False, ch) for ch in w_) + S.advem('B Lotus', False, ' ')
            if acc + wa > cap and cur_l:
                lines.append(cur_l)
                cur_l, acc = '', 0.0
            cur_l = (cur_l + ' ' + w_).strip()
            acc += wa
        if cur_l:
            lines.append(cur_l)
        note_wraps.append(lines or [''])
        note_heights.append(max(1, len(lines or [''])) * note_lineH + 2.0)

    def pheight(k):
        pa = paras[k]
        return pa['before'] + len(wraps[k]) * pa['lineH'] + pa['after']

    # پیش‌محاسبه ارجاع‌ها: ترتیب سراسری، خط هر ارجاع، ارتفاع یادداشت‌ها
    ref_base = {}
    _acc = 0
    for _k, _pa in enumerate(paras):
        ref_base[_k] = _acc
        _acc += len(_pa['refs'])

    def note_h(ordinal):
        return note_heights[ordinal] if ordinal < len(note_heights) else 20.0

    ref_lines = {}
    for _k, _pa in enumerate(paras):
        for _ri, _off in enumerate(_pa['refs']):
            _hit = None
            for _li, (_s, _e) in enumerate(wraps[_k]):
                if _s <= _off < _e or (_off == _e == len(_pa['text'])):
                    _hit = _li
                    break
            ref_lines[(_k, _ri)] = _hit if _hit is not None else len(wraps[_k]) - 1

    def para_fold_total(k):
        base = ref_base[k]
        return (sum(note_h(base + ri) for ri in range(len(paras[k]['refs']))) +
                (20.0 if paras[k]['refs'] else 0.0))

    def chunk_virt(k, a, b):
        base = ref_base[k]
        return sum(note_h(base + ri) for ri in range(len(paras[k]['refs']))
                   if a <= ref_lines[(k, ri)] <= b)

    def chunk_has_ref(k, a, b):
        return any(a <= ref_lines[(k, ri)] <= b for ri in range(len(paras[k]['refs'])))

    pages = []
    cur_page = {'items': [], 'sect': 0, 'fasl': '', 'disp': 1, 'has_ref': False}
    phys, disp, y, seci = 1, 1, 0.0, 0
    mapping = {}
    collapsed_after = set()
    fasl = ''

    def new_page():
        nonlocal phys, disp, y, cur_page
        pages.append(cur_page)
        phys += 1
        disp += 1
        y = 0.0
        cur_page = {'items': [], 'sect': seci, 'fasl': fasl, 'disp': disp, 'has_ref': False}

    i, n = 0, len(paras)
    while i < n:
        pa = paras[i]
        if pa['break']:
            new_page()
        else:
            H = pheight(i) + para_fold_total(i)
            if pa['keepNext'] and y > 0:
                need, j = H, i + 1
                while j < n and not paras[j]['break']:
                    if paras[j]['keepNext']:
                        need += pheight(j) + para_fold_total(j)
                        j += 1
                        continue
                    need += paras[j]['before'] + min(2.0, len(wraps[j])) * paras[j]['lineH']
                    break
                if y + need > pageH:
                    new_page()
            for nm in pa['bookmarks']:
                mapping[nm] = disp
            L = len(wraps[i])
            first, li = True, 0
            while True:
                avail = pageH - y
                cb = pa['before'] if first else 0.0
                rem_lines = L - li
                rem_virt = chunk_virt(i, li, L - 1)
                sep_full = 20.0 if (chunk_has_ref(i, li, L - 1) and not cur_page['has_ref']) else 0.0
                if cb + rem_lines * pa['lineH'] + rem_virt + sep_full + pa['after'] <= avail + 1e-6:
                    cur_page['items'].append((i, li, L - 1, y + cb))
                    y += cb + rem_lines * pa['lineH'] + rem_virt + sep_full + pa['after']
                    if chunk_has_ref(i, li, L - 1):
                        cur_page['has_ref'] = True
                    break
                fit = math.floor((avail - cb) / pa['lineH'] + 1e-9)
                if first and fit < 2 and rem_lines > 1 and y > 0:
                    new_page()
                    continue
                if fit < 1:
                    if y > 0:
                        new_page()
                        continue
                    fit = 1
                if rem_lines - fit == 1 and fit > 1:
                    fit -= 1
                fit = min(fit, rem_lines)
                while fit > 0:
                    _cv = chunk_virt(i, li, li + fit - 1)
                    _cs = 20.0 if (chunk_has_ref(i, li, li + fit - 1) and not cur_page['has_ref']) else 0.0
                    if fit * pa['lineH'] + _cv + _cs <= avail - cb + 1e-6:
                        break
                    fit -= 1
                if fit < 1:
                    if y > 0:
                        new_page()
                        continue
                    fit = 1
                if first and fit < 2 and rem_lines > 1 and y > 0:
                    new_page()
                    continue
                if fit == rem_lines:
                    # همه سطرها (ولی after نه) ← جاگذاری و شکست صفحه (after می‌ریزد)
                    cur_page['items'].append((i, li, L - 1, y + cb))
                    collapsed_after.add(i)
                    y += cb + rem_lines * pa['lineH'] + rem_virt + sep_full
                    if chunk_has_ref(i, li, L - 1):
                        cur_page['has_ref'] = True
                    new_page()
                    break
                cvirt = chunk_virt(i, li, li + fit - 1)
                csep = 20.0 if (chunk_has_ref(i, li, li + fit - 1) and not cur_page['has_ref']) else 0.0
                cur_page['items'].append((i, li, li + fit - 1, y + cb))
                y += cb + fit * pa['lineH'] + cvirt + csep
                if chunk_has_ref(i, li, li + fit - 1):
                    cur_page['has_ref'] = True
                li += fit
                new_page()
                first = False
            if pa['style'] == 'TH-Fasl':
                fasl = pa['text'].replace('\t', '').strip()
                cur_page['fasl'] = fasl
        if pa['sect'] is not None:
            seci += 1
            pages.append(cur_page)
            phys += 1
            y = 0.0
            st = sect_list[seci]['pgStart']
            disp = st if st is not None else disp + 1
            cur_page = {'items': [], 'sect': seci, 'fasl': fasl, 'disp': disp, 'has_ref': False}
        i += 1
    pages.append(cur_page)

    # خط هر ارجاع ← صفحه‌اش
    line_page = {}
    for pgi, pg in enumerate(pages):
        for (k, a, b, _yy) in pg['items']:
            for li in range(a, b + 1):
                line_page[(k, li)] = pgi
    ref_page = {}
    for k, pa in enumerate(paras):
        base = ref_base[k]
        for ri in range(len(pa['refs'])):
            ref_page[base + ri] = line_page.get((k, ref_lines[(k, ri)]), 0)
    page_notes = [[] for _ in pages]
    for rpi, pgi in ref_page.items():
        if rpi not in page_notes[pgi]:
            page_notes[pgi].append(rpi)
    for pgi, pg in enumerate(pages):
        pg['notes'] = sorted(page_notes[pgi])
    # اثبات عدم تداخل بدنه و پاورقی در هر صفحه
    for pgi, pg in enumerate(pages):
        V = 0.0
        for (k, a, b_, _yy) in pg['items']:
            V += (b_ - a + 1) * paras[k]['lineH']
            if a == 0:
                V += paras[k]['before']
            if b_ == len(wraps[k]) - 1 and k not in collapsed_after:
                V += paras[k]['after']
        ssum = sum(note_h(r) for r in pg['notes'])
        rhs = pageH - ssum - (20.0 if pg['notes'] else 0.0)
        assert V <= rhs + 1e-6, f'overlap risk pdf-p{pgi + 1}: V={V:.1f} > {rhs:.1f}'
    return {
        'paras': paras, 'pages': pages, 'mapping': mapping, 'ref_page': ref_page,
        'ref_base': ref_base, 'collapsed': collapsed_after, 'notes': notes, 'note_wraps': note_wraps,
        'note_heights': note_heights, 'sects': sect_list, 'wraps': wraps,
        'rid2part': rid2part,
        'geom': {'pgW': pgW, 'pgH': pgH, 'marL': marL, 'marR': marR, 'marT': marT,
                 'marB': marB, 'marHdr': marHdr, 'marFtr': marFtr, 'baseW': baseW,
                 'pageH': pageH},
    }


def main():
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.colors import HexColor, black
    import arabic_reshaper
    from bidi.algorithm import get_display
    from fontTools.ttLib import TTFont as FTFont

    paths = ensure_bfonts()
    dj = sorted(glob.glob('/usr/local/lib/python3*/dist-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf'))
    djb = sorted(glob.glob('/usr/local/lib/python3*/dist-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans-Bold.ttf'))
    paths[('DejaVu', False)] = dj[0]
    paths[('DejaVu', True)] = djb[0] if djb else dj[0]
    for key, pth in paths.items():
        pdfmetrics.registerFont(TTFont(f'F-{key[0]}-{int(key[1])}', pth))
    cmaps = {}
    for key, pth in paths.items():
        try:
            cmaps[key] = set(FTFont(pth).getBestCmap().keys())
        except Exception:
            cmaps[key] = set()

    def pick_font(fam, bold, ch):
        if ord(ch) in cmaps.get((fam, bold), set()):
            return f'F-{fam}-{int(bold)}'
        if ord(ch) in cmaps.get((fam, False), set()):
            return f'F-{fam}-0'
        if ord(ch) in cmaps.get(('DejaVu', bold), set()):
            return f'F-DejaVu-{int(bold)}'
        return 'F-DejaVu-0'

    L = layout_document(DOCX)
    paras, pages = L['paras'], L['pages']
    # مختصات فشرده بدنه (بدون شکاف مجازی پاورقی — مثل Word)
    for _pg in pages:
        _cy, _cy0 = 0.0, []
        for (_k, _a, _b, _y0) in _pg['items']:
            _pa = paras[_k]
            _cb = _pa['before'] if _a == 0 else 0.0
            _cy0.append(_cy + _cb)
            _cy += _cb + (_b - _a + 1) * _pa['lineH']
            if _b == len(L['wraps'][_k]) - 1 and _k not in L['collapsed']:
                _cy += _pa['after']
        _pg['cy0'] = _cy0
    g = L['geom']
    PW, PH = g['pgW'], g['pgH']
    xR = PW - g['marR']
    xL = g['marL']
    yTop = PH - g['marT']

    zf = zipfile.ZipFile(DOCX)
    hf_cache = {}

    def hf_paras(rid):
        if rid not in hf_cache:
            part = L['rid2part'].get(rid)
            if part is None:
                hf_cache[rid] = []
            else:
                root = ET.fromstring(zf.read(part).decode('utf8'))
                hf_cache[rid] = parse_paras(root.findall(f'.//{W}p'))
        return hf_cache[rid]

    c = canvas.Canvas(PDF, pagesize=(PW, PH))
    c.setTitle('عدول ولی فقیه از تعهدات بین‌المللی دولت اسلامی — از منظر فقه امامیه و حقوق بین‌الملل')
    c.setAuthor('هادی شبستانی')
    c.setSubject('رساله علمی سطح سه حوزه علمیه')
    reshaper = arabic_reshaper.ArabicReshaper(
        configuration={'delete_harakat': False, 'support_ligatures': True})

    def shape_run(tx):
        if FA_RE.search(tx):
            return get_display(reshaper.reshape(tx), base_dir='R')
        return get_display(tx, base_dir='L')

    def draw_runs_rtl(x_right, y, runs, size, fill=black):
        c.setFillColor(fill)
        for fam, bold, _sz, tx, _sup in runs:
            v = shape_run(tx)
            # گروه‌بندی بر اساس فونت واقعی هر حرف؛ سپس رسم چپ‌به‌راست (ترتیب دیداری)
            cur, cur_txt = None, ''
            groups = []
            for ch in v:
                fn = pick_font(fam, bold, ch)
                if fn != cur and cur_txt:
                    groups.append((cur, cur_txt))
                    cur_txt = ''
                cur = fn
                cur_txt += ch
            if cur_txt:
                groups.append((cur, cur_txt))
            widths = [pdfmetrics.stringWidth(gt, fn, size) for fn, gt in groups]
            x = x_right - sum(widths)
            for (fn, gt), w_ in zip(groups, widths):
                c.setFont(fn, size)
                c.drawString(x, y, gt)
                x += w_
            x_right -= sum(widths)
        return x_right

    def draw_runs_ltr_justified(x_left, width, y, runs, size, last_line, fill=black):
        c.setFillColor(fill)
        if last_line:
            x = x_left
            for fam, bold, _sz, tx, _sup in runs:
                v = get_display(tx, base_dir='L')
                fn = pick_font(fam, bold, 'a')
                c.setFont(fn, size)
                c.drawString(x, y, v)
                x += pdfmetrics.stringWidth(v, fn, size)
            return
        words = []
        for fam, bold, _sz, tx, _sup in runs:
            for w_ in tx.split(' '):
                words.append((fam, bold, w_))
        if len(words) < 2:
            draw_runs_ltr_justified(x_left, width, y, runs, size, True, fill)
            return
        tot = sum(pdfmetrics.stringWidth(w_, pick_font(f, b, 'a'), size) for f, b, w_ in words)
        sp = pdfmetrics.stringWidth(' ', pick_font('DejaVu', False, 'a'), size)
        extra = (width - tot - sp * (len(words) - 1)) / (len(words) - 1)
        x = x_left
        for idx, (fam, bold, w_) in enumerate(words):
            fn = pick_font(fam, bold, 'a')
            c.setFont(fn, size)
            c.drawString(x, y, w_)
            x += pdfmetrics.stringWidth(w_, fn, size) + sp + max(0.0, extra)

    def slice_runs(pa, s, e):
        out, pos = [], 0
        for fam, bold, sz, tx, sup in pa['runs']:
            ns, ne = pos, pos + len(tx)
            l, r_ = max(s, ns), min(e, ne)
            if l < r_:
                out.append((fam, bold, sz, tx[l - ns:r_ - ns], sup))
            pos = ne
        return out

    def run_adv_sum(runs):
        tot = 0.0
        for fam, bold, _sz, tx, _sup in runs:
            tot += sum(S.advem(fam, bold, ch) if ch not in '\t ' else 0.5 if ch == '\t' else S.advem(fam, bold, ' ') for ch in tx)
        return tot

    # شماره فهرست: ترتیب سطرها ← نشانک
    toc_paras = [k for k, pa in enumerate(paras) if pa['style'].startswith('TOC')]
    toc_num = {}
    for ti, k in enumerate(toc_paras):
        toc_num[k] = L['mapping'].get(f'_Toc{ti + 1:06d}', 0)

    for pgi, pg in enumerate(pages):
        sect = L['sects'][pg['sect']]
        if sect['border']:
            col, w_ = sect['border']
            c.setStrokeColor(HexColor('#' + col))
            c.setLineWidth(w_)
            c.rect(24, 24, PW - 48, PH - 48, stroke=1, fill=0)
        # سرصفحه / پاصفحه
        for kind, rid in (('hdr', sect['hdr']), ('ftr', sect['ftr'])):
            if not rid:
                continue
            for hpa in hf_paras(rid):
                if not hpa['text'] and not hpa['instrs']:
                    continue
                runs = []
                for fam, bold, sz, tx, sup in hpa['runs']:
                    runs.append((fam, bold, sz, tx, sup))
                if any('STYLEREF' in ins for ins in hpa['instrs']):
                    runs = [('B Lotus', False, hpa['size'], pg['fasl'], False)] if pg['fasl'] else []
                else:
                    for ins in hpa['instrs']:
                        if ins.startswith('PAGE'):
                            runs.append(('B Lotus', True, hpa['size'],
                                         str(pg['disp']).translate(FA_DIGITS), False))
                if not runs:
                    continue
                size = hpa['size']
                lh = hpa['lineH']
                if kind == 'hdr':
                    yb = PH - g['marHdr'] - hpa['before'] - (lh - size * 0.45)
                else:
                    yb = g['marFtr'] - hpa['before'] - (lh - size * 0.45) + 6
                draw_runs_rtl(xR, yb, runs, size)
                if 'bottom' in hpa['bdr']:
                    col, w_ = hpa['bdr']['bottom']
                    c.setStrokeColor(HexColor('#' + col))
                    c.setLineWidth(w_)
                    yy = yb - size * 0.45 - 4
                    c.line(xL, yy, xR, yy)
        # بدنه
        for _ii, (k, a, b, _y0) in enumerate(pg['items']):
            y0 = pg['cy0'][_ii]
            pa = paras[k]
            nlines = b - a + 1
            blk_h = pa['before'] + nlines * pa['lineH'] + (pa['after'] if b == len(L['wraps'][k]) - 1 else 0)
            y_top = yTop - (y0 - (pa['before'] if a == 0 else 0))
            y_bot = y_top - blk_h
            x0 = xL + pa['indent_l']
            x1 = xR - pa['indent_r'] - TOC_STYLE_INDENT.get(pa['style'], 0.0)
            if pa['shd']:
                c.setFillColor(HexColor('#' + pa['shd']))
                c.rect(x0, y_bot, x1 - x0, blk_h, stroke=0, fill=1)
            if pa['bdr']:
                for side in ('top', 'bottom'):
                    if side in pa['bdr']:
                        col, w_ = pa['bdr'][side]
                        c.setStrokeColor(HexColor('#' + col))
                        c.setLineWidth(w_)
                        c.line(x0, y_top if side == 'top' else y_bot, x1, y_top if side == 'top' else y_bot)
                if 'left' in pa['bdr'] and 'right' in pa['bdr']:
                    col, w_ = pa['bdr']['left']
                    c.setStrokeColor(HexColor('#' + col))
                    c.setLineWidth(w_)
                    c.rect(x0, y_bot, x1 - x0, blk_h, stroke=1, fill=0)
            yy = yTop - y0
            is_toc = pa['style'].startswith('TOC')
            for li in range(a, b + 1):
                s, e = L['wraps'][k][li]
                runs = slice_runs(pa, s, e)
                yb = yy - (pa['lineH'] - pa['size'] * 0.45)
                last_line = li == len(L['wraps'][k]) - 1
                if runs and any(tx.strip(' \t') for _, _, _, tx, _ in runs):
                    if is_toc:
                        truns = [(f, bo, sz, tx.replace('\t', ''), su) for f, bo, sz, tx, su in runs]
                        draw_runs_rtl(x1, yb, truns, pa['size'])
                        if last_line:
                            tw = 0.0
                            for f, bo, _sz, tx, _su in truns:
                                v = shape_run(tx)
                                cur, cur_txt = None, ''
                                for ch in v:
                                    fn = pick_font(f, bo, ch)
                                    if fn != cur and cur_txt:
                                        tw += pdfmetrics.stringWidth(cur_txt, cur, pa['size'])
                                        cur_txt = ''
                                    cur = fn
                                    cur_txt += ch
                                if cur_txt:
                                    tw += pdfmetrics.stringWidth(cur_txt, cur, pa['size'])
                            num = str(toc_num.get(k, '')).translate(FA_DIGITS)
                            fn = pick_font('B Lotus', False, '0')
                            nw = pdfmetrics.stringWidth(num, fn, pa['size']) if num else 0.0
                            c.setFillColor(black)
                            c.setFont(fn, pa['size'])
                            c.drawString(xL, yb, num)
                            gap = (x1 - tw) - (xL + nw)
                            if gap > 10 and num:
                                dw = pdfmetrics.stringWidth('.', fn, pa['size'])
                                nd = int(gap / (dw * 1.6))
                                dx = xL + nw + 4
                                for _di in range(nd):
                                    c.drawString(dx, yb, '.')
                                    dx += dw * 1.6
                    elif pa['jc'] == 'left' or (not FA_RE.search(pa['text']) and pa['text'].strip()):
                        draw_runs_ltr_justified(x0, x1 - x0, yb, runs, pa['size'], last_line)
                    else:
                        draw_runs_rtl(x1, yb, runs, pa['size'])
                    # نشان‌های ارجاع این سطر
                    base = L['ref_base'].get(k, 0)
                    for ri, off in enumerate(pa['refs']):
                        if s <= off < e or (off == e == len(pa['text']) and last_line):
                            pre = slice_runs(pa, s, off)
                            wpre = run_adv_sum(pre) * pa['size']
                            mk = str(base + ri + 1).translate(FA_DIGITS)
                            c.setFillColor(black)
                            c.setFont(pick_font('B Lotus', False, '0'), 9)
                            c.drawString(x1 - wpre - 9, yb + pa['size'] * 0.35, mk)
                yy -= pa['lineH']
        # پاورقی‌ها
        if pg['notes']:
            nh = 20.0 + sum(L['note_heights'][ni] for ni in pg['notes'])
            ysep = g['marB'] + nh - 20.0 + 6
            c.setStrokeColor(black)
            c.setLineWidth(0.75)
            c.line(xR - 70, ysep, xR, ysep)
            ny = ysep - 4
            nlh = 10.0 * S.PITCH[('B Lotus', False)]
            for ni in pg['notes']:
                for lni, ln in enumerate(L['note_wraps'][ni]):
                    ny -= nlh
                    nyb = ny - (nlh - 4.5)
                    if lni == 0:
                        mk = str(ni + 1).translate(FA_DIGITS)
                        c.setFillColor(black)
                        c.setFont(pick_font('B Lotus', False, '0'), 8)
                        mw = pdfmetrics.stringWidth(mk, pick_font('B Lotus', False, '0'), 8)
                        c.drawString(xR - mw, nyb + 3, mk)
                        draw_runs_rtl(xR - mw - 5, nyb, [('B Lotus', False, 10, ln, False)], 10)
                    else:
                        draw_runs_rtl(xR, nyb, [('B Lotus', False, 10, ln, False)], 10)
                ny -= 2.0
        c.showPage()
    c.save()
    zf.close()
    print(f'PDF saved: {PDF} pages={len(pages)}')
    return L


def verify(L):
    from pypdf import PdfReader
    import arabic_reshaper
    from bidi.algorithm import get_display
    reshaper = arabic_reshaper.ArabicReshaper(
        configuration={'delete_harakat': False, 'support_ligatures': True})
    rdr = PdfReader(PDF)
    print('pdf pages:', len(rdr.pages), '| layout pages:', len(L['pages']))
    assert len(rdr.pages) == len(L['pages'])
    # استخراج سطرها به ترتیب محتوا (pymupdf) — قابل اعتمادتر از pypdf برای راست‌به‌چپ
    import pymupdf
    d = pymupdf.open(PDF)
    tnorm = []
    for pgi in range(d.page_count):
        raw = d[pgi].get_text('rawdict')
        chars = []
        for blk in raw.get('blocks', []):
            for ln in blk.get('lines', []):
                for sp in ln.get('spans', []):
                    for ch in sp.get('chars', []):
                        chars.append((round(ch['bbox'][1], 1), ch['bbox'][0], ch['c']))
        chars.sort(key=lambda t: (t[0], t[1]))
        lines, cur_y, cur_t = [], None, ''
        for yy, xx, ch in chars:
            if cur_y is None or abs(yy - cur_y) > 2.0:
                if cur_t:
                    lines.append(cur_t)
                cur_y, cur_t = yy, ch
            else:
                cur_t += ch
        if cur_t:
            lines.append(cur_t)
        tnorm.append([re.sub('[\s\u200c\u064b-\u065f\ufe80\u0621]+', '', ln) for ln in lines])
    d.close()
    # هر سرفصل باید در صفحه نگاشتی‌اش پیدا شود (طولانی‌ترین ران ≥ ۱۲ حرف)
    paras = L['paras']
    names = sorted(L['mapping'])
    bad, checked = [], 0
    for idx, nm in enumerate(names):
        disp = L['mapping'][nm]
        # صفحه فیزیکی متناظر با شماره نمایشی: جست‌وجو در صفحات
        cand = [pgi for pgi, pg in enumerate(L['pages']) if pg['disp'] == disp and
                ((idx < 2 and pg['sect'] < 2) or (idx >= 2 and pg['sect'] == 2))]
        if not cand:
            bad.append((nm, disp, 'no-page'))
            continue
        # متن سرفصل از نشانک
        title = ''
        for pa in paras:
            if nm in pa['bookmarks']:
                title = pa['text'].replace('\t', '')
                break
        runs = [tx for _, _, _, tx, _ in pa['runs']] if title else []
        key = ''
        for tx in runs:
            if FA_RE.search(tx):
                v = get_display(reshaper.reshape(tx), base_dir='R')
            else:
                v = tx
            if len(v) > len(key):
                key = v
                key_logic = reshaper.reshape(tx) if FA_RE.search(tx) else tx
        key = re.sub('[\s\u200c\u064b-\u065f\ufe80\u0621]+', '', key.strip())
        if len(key) < 12:
            continue
        checked += 1
        hit = False
        for pgi in cand:
            LNs = tnorm[pgi]
            if any(key[:24] in ln or key[-24:] in ln for ln in LNs):
                hit = True
                break
            for a, b in zip(LNs, LNs[1:]):
                if key in (a + b) or key in (b + a):
                    hit = True
                    break
            if hit:
                break
        if not hit:
            bad.append((nm, disp, title[:30]))
    print(f'heading placement checked: {checked}, mismatches: {len(bad)}')
    for b in bad[:10]:
        print('  MISS:', b)
    # نمونه سطرهای فهرست
    for pgi, pg in enumerate(L['pages']):
        if any(paras[k]['style'].startswith('TOC') for k, _a, _b, _y in pg['items']):
            print(f'--- TOC lines sample (pdf p.{pgi + 1}) ---')
            shown = 0
            for ln in tnorm[pgi]:
                s = re.sub(r'\.{4,}', '....', ln)
                if s.strip('.'):
                    print('  ', s[:60])
                    shown += 1
                    if shown >= 6:
                        break
            break


if __name__ == '__main__':
    L = main()
    verify(L)
