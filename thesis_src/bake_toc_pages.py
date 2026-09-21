# -*- coding: utf-8 -*-
"""تزریق شماره صفحه (تخمین شبیه‌ساز کالیبره‌شده) به‌عنوان کش فیلدهای PAGEREF فهرست.

فیلدها زنده می‌مانند: Word هنگام باز شدن شماره‌های دقیق را جایگزین می‌کند.
نکته فنی: جراحی XML با lxml انجام می‌شود تا پیشوندهای فضای‌نام دقیقاً حفظ شوند
(ElementTree پیشوندها را به ns1/ns2 تغییر می‌داد و فایل در Word باز نمی‌شد).
اجرا: python3 bake_toc_pages.py  (بعد از build_thesis.py، روی فایل تازه)
"""
import re
import sys
import zipfile

from lxml import etree

import sim_paginate as S

W_SCALE = 1.56  # کالیبره با مشاهده واقعی کاربر: بیلد قدیمی = ۸۱ صفحه
OUT = '/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_Final2.docx'

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': WNS}
FA = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')


def main():
    S.WIDTH_SCALE_REF[0] = W_SCALE
    r = S.simulate(OUT, W_SCALE)
    mapping = r['mapping']
    assert len(mapping) == 179, f"bookmarks={len(mapping)}"
    print(f"sim: total_phys={r['total_phys']} bookmarks={len(mapping)}")

    zin = zipfile.ZipFile(OUT)
    items = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    orig = items['word/document.xml']
    root = etree.fromstring(orig)

    injected = 0
    for p in root.findall('./w:body/w:p', NS):
        armed = None
        for k in p:
            if k.tag != f'{{{WNS}}}r':
                continue
            it = k.find('w:instrText', NS)
            if it is not None and it.text and 'PAGEREF' in it.text:
                m = re.search(r'PAGEREF\s+(_Toc\d+)', it.text)
                assert m, f'bad PAGEREF instr: {it.text!r}'
                armed = m.group(1)
                assert armed in mapping, f'no page for {armed}'
                continue
            if armed is None:
                continue
            fc = k.find('w:fldChar', NS)
            if fc is not None and fc.get(f'{{{WNS}}}fldCharType') == 'separate':
                cr = etree.SubElement(p, f'{{{WNS}}}r')
                rPr = etree.SubElement(cr, f'{{{WNS}}}rPr')
                rf = etree.SubElement(rPr, f'{{{WNS}}}rFonts')
                rf.set(f'{{{WNS}}}ascii', 'B Lotus')
                rf.set(f'{{{WNS}}}hAnsi', 'B Lotus')
                rf.set(f'{{{WNS}}}cs', 'B Lotus')
                etree.SubElement(rPr, f'{{{WNS}}}sz').set(f'{{{WNS}}}val', '26')
                etree.SubElement(rPr, f'{{{WNS}}}szCs').set(f'{{{WNS}}}val', '26')
                etree.SubElement(rPr, f'{{{WNS}}}rtl')
                t = etree.SubElement(cr, f'{{{WNS}}}t')
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t.text = str(mapping[armed]).translate(FA)
                # انتقال درست بعد از separate (نه انتهای پاراگراف)
                p.remove(cr)
                k.addnext(cr)
                injected += 1
                armed = None
                continue
            t_el = k.find('w:t', NS)
            if t_el is not None and (t_el.text or '').strip():
                print('ABORT: already baked (stale cache present). Rebuild fresh first.',
                      file=sys.stderr)
                sys.exit(2)

    assert injected == 179, f'injected={injected}'
    new = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # نگهبان: پیشوندهای فضای‌نام نباید تغییر کنند (وگرنه Word باز نمی‌کند)
    assert not re.search(rb'\bns\d+:', new), 'namespace prefixes mangled!'
    oi = re.search(rb'Ignorable="[^"]*"', orig)
    ni = re.search(rb'Ignorable="[^"]*"', new)
    assert oi and ni and oi.group(0) == ni.group(0), 'mc:Ignorable changed!'

    items['word/document.xml'] = new
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)

    names = sorted(mapping)
    front = [mapping[n] for n in names[:2]]
    main = [mapping[n] for n in names[2:]]
    assert all(b >= a for a, b in zip(main, main[1:])), 'main TOC not monotonic!'
    print(f'baked 179 page caches; front={front} main=1..{max(main)}')
    print('sample:', [(names[i], mapping[names[i]]) for i in (0, 1, 2, 50, 100, 178)])


if __name__ == '__main__':
    main()
