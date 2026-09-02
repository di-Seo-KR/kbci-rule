# -*- coding: utf-8 -*-
import json, os, subprocess, html
from PIL import Image, ImageChops
CHROME='/opt/pw-browsers/chromium-1194/chrome-linux/chrome'
HERE=os.path.abspath('.')
T=json.load(open('tables.json',encoding='utf-8'))
CSS='''@font-face{font-family:NG;src:url("file://%s/fonts/NanumGothic-Regular.ttf");font-weight:400}
@font-face{font-family:NG;src:url("file://%s/fonts/NanumGothic-Bold.ttf");font-weight:700}
body{margin:0;padding:24px;background:#fff;font-family:NG,"WenQuanYi Zen Hei",sans-serif;color:#111}
table{border-collapse:collapse;table-layout:fixed;font-size:22px;line-height:1.5}
td{border:1.5px solid #222;padding:9px 14px;vertical-align:middle;word-break:keep-all}
tr.h td{background:#efefef;font-weight:700;text-align:center}
td.c{text-align:center} td.l{text-align:left}
td p{margin:0}
td p.hang{padding-left:1.6em;text-indent:-1.6em}''' % (HERE, HERE)

def cell_html(lines, hdr):
    # 줄바꿈으로 끊긴 이어지는 줄(앞 공백으로 시작)은 앞 줄에 붙인다
    merged=[]
    for l in lines:
        t=l['text']
        if merged and t.startswith(' ') and not t.strip()[:2].rstrip('.').isdigit():
            merged[-1]['text']=merged[-1]['text'].rstrip()+' '+t.strip()
        else:
            merged.append(dict(l, text=t.strip()))
    ps=[]
    for l in merged:
        cls=' class="hang"' if (len(merged)>1 and l['text'][:2].rstrip('.').isdigit()) else ''
        ps.append('<p%s>%s</p>' % (cls, html.escape(l['text']).replace('  ',' &nbsp;')))
    return ''.join(ps)

def build(idx, tbl, total_px):
    grid=tbl['grid']; s=sum(grid); widths=[round(total_px*g/s) for g in grid]
    rows=[]
    for ri,row in enumerate(tbl['rows']):
        tds=[]
        for ci,cell in enumerate(row):
            jc=cell[0]['jc']; multi=len(cell)>1 or len(cell[0]['text'])>18
            cls='c' if (jc=='center' or ri==0) else ('l' if multi or jc=='both' else 'c')
            tds.append('<td class="%s" style="width:%dpx">%s</td>' % (cls, widths[ci]-30, cell_html(cell, ri==0)))
        rows.append('<tr%s>%s</tr>' % (' class="h"' if ri==0 else '', ''.join(tds)))
    doc='<!doctype html><html><head><meta charset="utf-8"><style>%s</style></head><body><table>%s</table></body></html>' % (CSS, ''.join(rows))
    hp='t%d.html'%idx; open(hp,'w',encoding='utf-8').write(doc)
    png='t%d.png'%idx
    subprocess.run([CHROME,'--headless=new','--no-sandbox','--disable-gpu','--hide-scrollbars',
                    '--force-device-scale-factor=2','--window-size=%d,900'%(total_px+60),
                    '--screenshot=%s'%png,'file://%s/%s'%(HERE,hp)],capture_output=True,timeout=120)
    im=Image.open(png).convert('RGB')
    bg=Image.new('RGB',im.size,(255,255,255)); bbox=ImageChops.difference(im,bg).getbbox()
    if bbox:
        pad=36; im=im.crop((max(0,bbox[0]-pad),max(0,bbox[1]-pad),min(im.width,bbox[2]+pad),min(im.height,bbox[3]+pad)))
    im.save(png, dpi=(192,192)); print(png, im.size)

build(1, T[0], 1000)   # 제6조 안식년 휴가 표 (7열)
build(2, T[1], 860)    # 제24조 무단결근 조치 표 (2열)
