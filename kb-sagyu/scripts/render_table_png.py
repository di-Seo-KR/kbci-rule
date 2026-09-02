#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""사규 docx 안의 표를 전부 PNG로 렌더링한다 (사규관리시스템이 docx 표를 파싱하지 못해 이미지로 등재).

    python3 kb-sagyu/scripts/render_table_png.py "<docx>" [출력폴더]

- 셀 텍스트·정렬·열 너비(tblGrid 비율)·gridSpan/vMerge 병합을 docx에서 그대로 추출
- HTML → 헤드리스 Chromium(2배율) 캡처 → 여백 트리밍, 192dpi PNG
- 파일명: <docx 이름>_표N_<바로 앞 조문>.png  (예: (6-12)_복무규정 시행지침_260831_표1_제6조.png)
- 글꼴: scripts/fonts/NanumGothic (KB금융 서체 미보유). 크로미움: /opt/pw-browsers/chromium-*/chrome-linux/chrome
"""
import sys, os, re, glob, json, html, subprocess
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, 'fonts')
CHROME = (glob.glob('/opt/pw-browsers/chromium-*/chrome-linux/chrome') or ['chromium'])[0]
CSS = '''@font-face{font-family:NG;src:url("file://%(f)s/NanumGothic-Regular.ttf");font-weight:400}
@font-face{font-family:NG;src:url("file://%(f)s/NanumGothic-Bold.ttf");font-weight:700}
body{margin:0;padding:24px;background:#fff;font-family:NG,"WenQuanYi Zen Hei",sans-serif;color:#111}
table{border-collapse:collapse;table-layout:fixed;font-size:22px;line-height:1.5}
td{border:1.5px solid #222;padding:9px 14px;vertical-align:middle;word-break:keep-all}
tr.h td{background:#efefef;font-weight:700;text-align:center}
td.c{text-align:center} td.l{text-align:left} td p{margin:0}
td p.hang{padding-left:1.6em;text-indent:-1.6em}''' % {'f': FONT_DIR}


def extract(path):
    """[(앞 조문 라벨, {'grid':[...], 'rows':[[cell,...],...]}), ...]  cell={'lines':[...], 'span':n, 'vm':None|'restart'|'cont'}"""
    d = docx.Document(path); out = []; last_art = ''
    for el in d.element.body.iterchildren():
        if el.tag == qn('w:p'):
            t = Paragraph(el, d).text.strip()
            m = re.match(r'(제\d+조(?:의\d+)?)', t)
            if m: last_art = m.group(1)
            continue
        if el.tag != qn('w:tbl'): continue
        t = Table(el, d); rows = []
        for row in t.rows:
            cells = []
            for tc in row._tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr')); span = 1; vm = None
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan')); span = int(gs.get(qn('w:val'))) if gs is not None else 1
                    v = tcPr.find(qn('w:vMerge')); vm = (v.get(qn('w:val')) or 'cont') if v is not None else None
                lines = []
                for p in tc.findall(qn('w:p')):
                    pp = Paragraph(p, d); pPr = p.find(qn('w:pPr'))
                    jc = pPr.find(qn('w:jc')).get(qn('w:val')) if (pPr is not None and pPr.find(qn('w:jc')) is not None) else None
                    if pp.text.strip(): lines.append({'text': pp.text, 'jc': jc})
                cells.append({'lines': lines or [{'text': '', 'jc': None}], 'span': span, 'vm': vm})
            rows.append(cells)
        grid = [int(g.get(qn('w:w'))) for g in el.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]
        out.append((last_art, {'grid': grid, 'rows': rows}))
    return out


def cell_html(lines):
    merged = []
    for l in lines:                       # 강제 줄바꿈된 이어지는 줄(앞 공백 시작)은 앞 줄에 붙인다
        t = l['text']
        if merged and t.startswith(' ') and not t.strip()[:2].rstrip('.').isdigit():
            merged[-1] = merged[-1].rstrip() + ' ' + t.strip()
        else:
            merged.append(t.strip())
    ps = []
    for t in merged:
        cls = ' class="hang"' if (len(merged) > 1 and t[:2].rstrip('.').isdigit()) else ''
        ps.append('<p%s>%s</p>' % (cls, html.escape(t).replace('  ', ' &nbsp;')))
    return ''.join(ps)


def render(tbl, png, total_px=None):
    grid = tbl['grid']; ncol = len(grid)
    total_px = total_px or min(1100, max(600, 150 * ncol))
    s = sum(grid); widths = [round(total_px * g / s) for g in grid]
    # vMerge: 세로 병합 rowspan 계산
    rows = tbl['rows']; rowspan = {}
    for ri, row in enumerate(rows):
        ci = 0
        for cj, c in enumerate(row):
            if c['vm'] == 'restart':
                k = ri + 1; n = 1
                while k < len(rows):
                    cc = [x for x in rows[k]]; pos = 0; hit = None
                    for x in cc:
                        if pos == ci: hit = x; break
                        pos += x['span']
                    if hit and hit['vm'] == 'cont': n += 1; k += 1
                    else: break
                rowspan[(ri, cj)] = n
            ci += c['span']
    trs = []
    for ri, row in enumerate(rows):
        tds = []; ci = 0
        for cj, c in enumerate(row):
            if c['vm'] == 'cont': ci += c['span']; continue
            w = sum(widths[ci:ci + c['span']]) - 30
            jc = c['lines'][0]['jc']; multi = len(c['lines']) > 1 or len(c['lines'][0]['text']) > 18
            cls = 'c' if (jc == 'center' or ri == 0) else ('l' if (multi or jc == 'both') else 'c')
            attrs = ' colspan="%d"' % c['span'] if c['span'] > 1 else ''
            if (ri, cj) in rowspan: attrs += ' rowspan="%d"' % rowspan[(ri, cj)]
            tds.append('<td class="%s" style="width:%dpx"%s>%s</td>' % (cls, w, attrs, cell_html(c['lines'])))
            ci += c['span']
        trs.append('<tr%s>%s</tr>' % (' class="h"' if ri == 0 else '', ''.join(tds)))
    hp = png[:-4] + '.html'
    open(hp, 'w', encoding='utf-8').write('<!doctype html><html><head><meta charset="utf-8"><style>%s</style></head><body><table>%s</table></body></html>' % (CSS, ''.join(trs)))
    subprocess.run([CHROME, '--headless=new', '--no-sandbox', '--disable-gpu', '--hide-scrollbars',
                    '--force-device-scale-factor=2', '--window-size=%d,1400' % (total_px + 60),
                    '--screenshot=%s' % png, 'file://' + os.path.abspath(hp)], capture_output=True, timeout=180)
    os.remove(hp)
    im = Image.open(png).convert('RGB')
    bbox = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255))).getbbox()
    if bbox:
        pad = 36; im = im.crop((max(0, bbox[0] - pad), max(0, bbox[1] - pad), min(im.width, bbox[2] + pad), min(im.height, bbox[3] + pad)))
    im.save(png, dpi=(192, 192)); return im.size


def main():
    if len(sys.argv) < 2:
        print(__doc__); return 1
    src = sys.argv[1]; outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(src)), 'img')
    os.makedirs(outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(src))[0]
    tables = extract(src)
    if not tables:
        print('표 없음:', src); return 0
    made = []
    for i, (art, tbl) in enumerate(tables, 1):
        png = os.path.join(outdir, '%s_표%d_%s.png' % (base, i, art or '위치미상'))
        size = render(tbl, png); made.append(png)
        print('생성: %s  %dx%d  (%d행 %d열)' % (os.path.basename(png), size[0], size[1], len(tbl['rows']), len(tbl['grid'])))
    return 0


if __name__ == '__main__':
    sys.exit(main())
