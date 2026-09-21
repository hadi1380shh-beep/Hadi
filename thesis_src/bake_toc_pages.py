# -*- coding: utf-8 -*-
"""تزریق شماره صفحه (تخمین شبیه‌ساز کالیبره‌شده) به‌عنوان کش فیلدهای PAGEREF فهرست.

فیلدها زنده می‌مانند: Word هنگام باز شدن شماره‌های دقیق را جایگزین می‌کند.
اجرا: python3 bake_toc_pages.py  (بعد از build_thesis.py، روی فایل تازه)
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

import sim_paginate as S

W_SCALE = 1.56  # کالیبره با مشاهده واقعی کاربر: بیلد قدیمی = ۸۱ صفحه
OUT = '/home/user/Hadi/Resaleh_Sath3_Hadi_Shabestani_v2.docx'

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
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
    root = ET.fromstring(items['word/document.xml'].decode('utf8'))

    injected = 0
    for p in root.findall(f'{W}body/{W}p'):
        kids = list(p)
        armed = None
        for idx, k in enumerate(kids):
            if k.tag != f'{W}r':
                continue
            it = k.find(f'{W}instrText')
            if it is not None and it.text and 'PAGEREF' in it.text:
                m = re.search(r'PAGEREF\s+(_Toc\d+)', it.text)
                assert m, f'bad PAGEREF instr: {it.text!r}'
                armed = m.group(1)
                assert armed in mapping, f'no page for {armed}'
                continue
            if armed is None:
                continue
            fc = k.find(f'{W}fldChar')
            if fc is not None and fc.get(f'{W}fldCharType') == 'separate':
                cr = ET.Element(f'{W}r')
                rPr = ET.SubElement(cr, f'{W}rPr')
                rf = ET.SubElement(rPr, f'{W}rFonts')
                rf.set(f'{W}ascii', 'B Lotus')
                rf.set(f'{W}hAnsi', 'B Lotus')
                rf.set(f'{W}cs', 'B Lotus')
                ET.SubElement(rPr, f'{W}sz').set(f'{W}val', '26')
                ET.SubElement(rPr, f'{W}szCs').set(f'{W}val', '26')
                ET.SubElement(rPr, f'{W}rtl')
                t = ET.SubElement(cr, f'{W}t')
                t.set(XML_SPACE, 'preserve')
                t.text = str(mapping[armed]).translate(FA)
                p.insert(list(p).index(k) + 1, cr)
                injected += 1
                armed = None
                continue
            if k.find(f'{W}t') is not None:
                # متنی بین دستور PAGEREF و separate/end یعنی قبلاً پخته شده
                t_el = k.find(f'{W}t')
                if (t_el.text or '').strip():
                    print('ABORT: already baked (stale cache present). Rebuild fresh first.',
                          file=sys.stderr)
                    sys.exit(2)

    assert injected == 179, f'injected={injected}'
    items['word/document.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
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
