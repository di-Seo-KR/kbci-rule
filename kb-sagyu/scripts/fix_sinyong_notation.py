# -*- coding: utf-8 -*-
"""(8-6) 신용정보관리 지침 260901 표기 정정 + (9-1) 신용정보업무 규정 개정 전문 생성.
기준: 「사규관리규정」(2026. 8. 12. 개정-2026-024) 제11조③·제12조⑥·제14조②④·제16조·제17조."""
import re, shutil, sys
import docx
from notation_lib import (ptext, set_text, replace, paint, paint_marks, MARK, BLUE, BLACK,
                    paras, merge_into_prev, fix_common, final_pass, LOG)
from byeolpyo_lib import add_list

LAW_BRACKET = re.compile(r'「([^」]*(?:법률|법|시행령|시행규칙|감독규정))」')
SAGYU_DOUBLE = {'『정관』': '「정관」', '『신용정보 관리지침』': '「신용정보관리 지침」',
                '『신용정보관리지침』': '「신용정보관리 지침」',
                '『개인정보보호 내부 관리계획』': '「개인정보보호 내부 관리계획」',
                '『채권추심업무처리지침』': '「채권추심업무 처리 지침」'}


def pre_normalize_marks(doc):
    """[…]·<…>·역순(<개정 일자>) 형태의 개정 표기를 (일자 개정) 형태로 통일한다."""
    n = 0
    for p in paras(doc):
        t = ptext(p)
        new = t
        new = re.sub(r'\[\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?\s*(본[조항호목]\s*)?(개정|신설|삭제)\s*\]',
                     lambda m: '(%s. %d. %d. %s)' % (m.group(1), int(m.group(2)), int(m.group(3)), m.group(5)), new)
        new = re.sub(r'<\s*(개정|신설)\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?\s*>',
                     lambda m: '(%s. %d. %d. %s)' % (m.group(2), int(m.group(3)), int(m.group(4)), m.group(1)), new)
        new = re.sub(r'삭제\s*<\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?\s*>',
                     lambda m: '(%s. %d. %d. 삭제)' % (m.group(1), int(m.group(2)), int(m.group(3))), new)
        # "① 삭제 (일자 삭제)" / "제12조 삭제 (일자 삭제)" → 문장을 지우고 표기만 (제16조②3호)
        new = re.sub(r'(^[①-⑮]|^제\d+조)\s*삭제\s*[\(（]', r'\1 (', new)
        if new != t:
            set_text(p, new); n += 1
    if n: LOG.append('  [제16조②] [ ]·< >·역순 개정 표기를 "(일자 변경사항)" 형식으로 통일 %d건' % n)


def cite_fix(doc, extra_laws=(), extra_sagyu=()):
    n = 0
    for p in paras(doc):
        t = ptext(p)
        for m in list(LAW_BRACKET.finditer(t)):
            if replace(p, m.group(0), '『%s』' % m.group(1), must=False): n += 1
        for a, b in SAGYU_DOUBLE.items():
            while a in ptext(p):
                if not replace(p, a, b, must=False): break
                n += 1
        for a, b in list(extra_laws) + list(extra_sagyu):
            while a in ptext(p) and b not in ptext(p):
                if not replace(p, a, b, must=False): break
                n += 1
    if n: LOG.append('  [제12조⑥] 법령 『』·사규 「」 인용 표기 정정 %d건' % n)


def byeol_fix(doc, pairs):
    n = 0
    for p in paras(doc):
        for a, b in pairs:
            while a in ptext(p):
                if not replace(p, a, b, must=False): break
                n += 1
    if n: LOG.append('  [제11조③] 별표·서식 표기 [별표 제N호]·[서식 제N호]로 정정 %d건' % n)


def merge_stray_marks(doc):
    n = 0
    for i in range(len(paras(doc)) - 1, 0, -1):
        if MARK.fullmatch(ptext(paras(doc)[i]).strip()):
            merge_into_prev(doc, i, sep=' '); n += 1
    if n: LOG.append('  [제16조②] 별도 단락으로 떨어진 개정 표기를 해당 조·항 문장 끝으로 이동 %d건' % n)


def merge_title_and_first(doc):
    """조 제목 단독 단락 + 다음 줄 ① → 한 단락 (제14조④: 제2항부터 줄바꿈)."""
    n = 0
    ps = paras(doc)
    for i in range(len(ps) - 1, 0, -1):
        prev = ptext(ps[i - 1]).strip(); t = ptext(ps[i]).strip()
        if re.fullmatch(r'제\d+조(?:의\d+)?\([^)]*\)(?:\s*\(\d{4}\. \d{1,2}\. \d{1,2}\. (?:개정|신설)\))?', prev) and t.startswith('①'):
            merge_into_prev(doc, i, sep=' '); n += 1
    if n: LOG.append('  [제14조④] 조 제목과 제1항이 나뉜 단락 병합 %d건' % n)


def merge_continuation(doc, starts):
    n = 0
    ps = paras(doc)
    for i in range(len(ps) - 1, 0, -1):
        t = ptext(ps[i]).strip()
        if any(t.startswith(s) for s in starts):
            merge_into_prev(doc, i, sep=' '); n += 1
    if n: LOG.append('  [제12조④] 한 문장이 나뉜 단락 병합 %d건' % n)


# ═══════════════ A. (8-6) 신용정보관리 지침 ═══════════════
def fix_A(src, out):
    shutil.copy(src, out)
    doc = docx.Document(out)
    LOG.append('■ 신용정보관리 지침 (2026. 9. 1.)')
    pre_normalize_marks(doc)
    fix_common(doc, '지침')
    merge_stray_marks(doc)
    merge_continuation(doc, ('이 경우 대통령령으로 정하는 바에 따라 해당 신용정보주체의 동의가',))
    byeol_fix(doc, [('<별표1>', '[별표 제1호]'), ('“<별표 2> 제재양정기준”', '[별표 제2호]'),
                    ('서식<별지서식1>에 따라', '서식[서식 제1호]에 따라'),
                    ('“<별지서식2> 신용정보 (열람, 정정, 삭제, 동의철회)요구서”', '신용정보(열람, 정정, 삭제, 동의철회) 요구서[서식 제2호]'),
                    ('“<별지서식2> 신용정보 (열람, 정정, 삭제, 동의철회) 요구서”', '신용정보(열람, 정정, 삭제, 동의철회) 요구서[서식 제2호]'),
                    ('“<별지서식2> 신용정보(열람, 정정, 삭제, 동의철회) 요구서”', '신용정보(열람, 정정, 삭제, 동의철회) 요구서[서식 제2호]'),
                    ('“<별지서식3> 신용정보 (열람, 일부열람, 열람연기, 열람거절)통지서”', '신용정보(열람, 일부열람, 열람연기, 열람거절) 통지서[서식 제3호]'),
                    ('<별지서식4> “신용정보(정정,삭제,동의 철회)요구에 대한 결과 통지서”', '신용정보(정정, 삭제, 동의철회) 요구에 대한 결과 통지서[서식 제4호]'),
                    ('요청자로부터신용정보(', '요청자로부터 신용정보('), ('확인한 후신용정보(', '확인한 후 신용정보(')])
    cite_fix(doc,
             extra_laws=[('가. 주민등록법 제7조의2', '가. 『주민등록법』 제7조의2'),
                         ('나. 여권법 제7조', '나. 『여권법』 제7조'),
                         ('다. 도로교통법 제80조', '다. 『도로교통법』 제80조'),
                         ('라. 출입국관리법 제31조', '라. 『출입국관리법』 제31조'),
                         ('마. 재외동포의 출입국과 법적 지위에 관한 법률 제7조', '마. 『재외동포의 출입국과 법적 지위에 관한 법률』 제7조'),
                         ('(신용정보법 제32조제4항의', '(『신용정보법』 제32조제4항의')],
             extra_sagyu=[('의거 인사규정 및', '의거 「인사규정」 및')])
    final_pass(doc)
    doc.save(out)
    add_list(out, out, [
        ('[별표 제1호] 개인신용정보의 제공 사실 및 이유 등을 알리거나 공시하는 시기 및 방법', '<○○○○. ○. ○. 개정>'),
        ('[별표 제2호] 제재 양정 기준', '<○○○○. ○. ○. 개정>'),
        ('[서식 제1호] 신용정보관리·보호인의 업무수행 실적 등 보고서', '<○○○○. ○. ○. 개정>'),
        ('[서식 제2호] 신용정보(열람, 정정, 삭제, 동의철회) 요구서', '<○○○○. ○. ○. 개정>'),
        ('[서식 제3호] 신용정보(열람, 일부열람, 열람연기, 열람거절) 통지서', '<○○○○. ○. ○. 개정>'),
        ('[서식 제4호] 신용정보(정정, 삭제, 동의철회) 요구에 대한 결과 통지서', '<○○○○. ○. ○. 개정>'),
        ('[서식 제5호] 신용정보시정요청서', '<○○○○. ○. ○. 개정>'),
    ])


# ═══════════════ B. (9-1) 신용정보업무 규정 — 개정 반영 + 표기 정정 ═══════════════
NEW_2_1 = ('① "신용정보"란 금융거래 등 상거래에서 거래 상대방의 신용을 판단할 때 필요한 정보로서 '
           '『신용정보의 이용 및 보호에 관한 법률』 제2조(정의)를 따른다. (2026. 9. 1. 개정)')
NEW_15_4 = '④ 『신용정보의 이용 및 보호에 관한 법률』 제32조제6항 각 호의 경우 (2026. 9. 1. 신설)'


def build_B(src, out):
    shutil.copy(src, out)
    doc = docx.Document(out)
    LOG.append('■ 신용정보업무 규정 (2026. 9. 1. 개정 반영)')

    # ── 개정 반영 (신구대비표)
    for p in paras(doc):
        t = ptext(p).strip()
        if t.startswith('① "신용정보"라 함은 금융거래등 상거래에 있어서'):
            set_text(p, NEW_2_1); LOG.append('  [개정] 제2조제1항 신용정보 정의 → 신용정보법 제2조 인용'); break
    ps = paras(doc)
    for i, p in enumerate(ps):
        if ptext(p).strip().startswith('③ 개인이 직접 제공한 개인신용정보'):
            from copy import deepcopy
            new = deepcopy(p._element); p._element.addnext(new)
            np = docx.text.paragraph.Paragraph(new, p._parent)
            for r in np.runs[1:]: r._element.getparent().remove(r._element)
            np.runs[0].text = NEW_15_4
            LOG.append('  [개정] 제15조제4항 신설 — 신용정보법 제32조제6항 각 호의 경우'); break
    # 부칙 추가
    last = paras(doc)[-1]
    from copy import deepcopy
    head_src = [p for p in paras(doc) if ptext(p).strip().startswith('부 칙')][-1]
    body_src = [p for p in paras(doc) if ptext(p).strip().startswith('① (시행일)')][-1]
    blank = deepcopy(last._element); 
    for ch in list(blank):
        if ch.tag != docx.oxml.ns.qn('w:pPr'): blank.remove(ch)
    h = deepcopy(head_src._element); b = deepcopy(body_src._element); bl2 = deepcopy(blank)
    last._element.addnext(blank); blank.addnext(h); h.addnext(bl2); bl2.addnext(b)
    hp = docx.text.paragraph.Paragraph(h, last._parent); bp = docx.text.paragraph.Paragraph(b, last._parent)
    set_text(hp, '부 칙 (2026. 9. 1.)'); set_text(bp, '① (시행일) 이 규정은 2026. 9. 1. 부터 시행한다.')
    LOG.append('  [개정] 부칙 신설 — 2026. 9. 1. 시행')
    # 머리 표기
    for p in paras(doc)[:5]:
        if '[시행' in ptext(p):
            set_text(p, '[시행 2026. 9. 1.] [2026. 9. 1. 개정-2026-0○○]'); break

    # ── 표기 정정
    pre_normalize_marks(doc)
    for p in paras(doc):
        if ptext(p).strip().startswith('제19조2('):
            replace(p, '제19조2(', '제19조의2('); LOG.append('  [제14조②] "제19조2" → "제19조의2"')
    fix_common(doc, '규정')
    merge_stray_marks(doc)
    # 제15조의2 본조 신설 표기: 제목 뒤 → 마지막 항 끝
    ps = paras(doc)
    for i, p in enumerate(ps):
        if ptext(p).strip().startswith('제15조의2('):
            m = MARK.search(ptext(p))
            if m:
                replace(p, ' ' + m.group(0), '')
                j = i + 1
                while j < len(ps) and not ptext(ps[j]).strip().startswith('제4장'): j += 1
                k = j - 1
                while not ptext(ps[k]).strip(): k -= 1
                last_run = ps[k].runs[-1]
                from fixlib import clone_run
                last_run._element.addnext(clone_run(last_run._element, ' ' + m.group(0)))
                paint_marks(ps[k])
                LOG.append('  [제16조②] 제15조의2 본조 신설 표기를 제목 뒤 → 마지막 항 끝으로 이동')
            break
    merge_title_and_first(doc)
    byeol_fix(doc, [('『별지 서식 제1호』', '[서식 제1호]'), ('『별표 제1호』', '[별표 제1호]')])
    cite_fix(doc, extra_laws=[('「동법 시행령」, 「동법시행규칙」', '『동법 시행령』, 『동법 시행규칙』'),
                              ('정관 제2조(목적 및 업무의 범위)', '「정관」 제2조(목적 및 업무의 범위)'),
                              ('신용정보업 감독규정 별지 제18호 서식', '『신용정보업 감독규정』 별지 제18호 서식')])
    final_pass(doc)
    doc.save(out)
    add_list(out, out, [
        ('[별표 제1호] 신용정보집중기관을 통하여 집중관리·활용되는 신용정보의 범위', '<○○○○. ○. ○. 개정>'),
        ('[서식 제1호] 신분증', '<○○○○. ○. ○. 개정>'),
    ])


if __name__ == '__main__':
    fix_A('A_신용정보관리지침_260901.docx', 'out_신용정보관리지침_260901.docx')
    LOG.append('')
    build_B('/home/user/kbci-rule/kb-sagyu/sources/(9-1)_신용정보업무 규정_250401.docx', 'out_신용정보업무규정_260901.docx')
    print('\n'.join(LOG))
