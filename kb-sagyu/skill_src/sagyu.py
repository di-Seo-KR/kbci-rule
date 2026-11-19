#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KB신용정보 사규 조회 도구 (표준 라이브러리만 사용).

  python3 sagyu.py list [키워드]                 사규 목록(분류번호·소관부서·시행일)
  python3 sagyu.py find 키워드 [키워드2 …]        조문 단위 검색(모든 키워드 포함). --or, --reg 사규, --no-att, --max N
  python3 sagyu.py show 사규 [조번호 …]           사규 전문 또는 지정한 조 (예: show 감사규정 5 6 / show 7-1 71)
  python3 sagyu.py toc 사규                      조 제목 목록 + 그 사규의 별표·서식 목록
  python3 sagyu.py att 키워드 [키워드2 …]         별표·서식 찾기. 1건이면 전문 출력, 여러 건이면 목록
  python3 sagyu.py jg 키워드 [키워드2 …]          직무전결기준표에서 업무 찾기(전결권자·비고). 키워드 없이 --notes 는 적용사항
  python3 sagyu.py map 키워드 [키워드2 …]         정독 카드에서 관련 사규 찾기(주제·성격으로 탐색, 하나라도 포함) — 질문을 받으면 가장 먼저
  python3 sagyu.py card 사규                     그 사규의 정독 요약 카드(핵심 규칙·예외·다른 사규와의 연결·정비 포인트)
  python3 sagyu.py exc 키워드 [키워드2 …]         예외·생략·위임·준용 조항만 검색(불구하고/생략/갈음/준용/위임/적용하지 아니)
  python3 sagyu.py where 키워드 [키워드2 …]       키워드가 나오는 모든 조를 한 줄씩(건수 제한 없음)
  python3 sagyu.py cite 사규                     이 사규를 인용하는 사규 / 이 사규가 인용하는 사규
  python3 sagyu.py ledger [키워드]               대장 목록(근거 조문·서식)

사규 지정: 분류번호(11-1) 또는 사규명 일부(공백 무시). 정확히 같은 이름이 있으면 그것을 우선한다.
"""
import sys, re, json, csv
from pathlib import Path

DATA = Path(__file__).resolve().parents[1] / 'data'
INDEX = json.load(open(DATA / 'index.json', encoding='utf-8'))
ART = re.compile(r'^제\s*(\d+)\s*조(?:\s*의\s*(\d+))?\s*[\(（]')


def sq(s):
    return re.sub(r'\s+', '', s)


def resolve(q):
    q0 = q.strip('()（）'); qs = sq(q0)
    regs = INDEX['regs']
    hit = [r for r in regs if r['key'] == q0]
    if not hit: hit = [r for r in regs if sq(r['name']) == qs]
    if not hit: hit = [r for r in regs if qs in sq(r['name'])]
    if not hit:
        sys.exit(f'사규를 찾지 못함: {q}  (python3 sagyu.py list 로 이름 확인)')
    if len(hit) > 1:
        print('여러 사규가 해당됨 — 분류번호로 다시 지정:', ', '.join(f"({r['key']}) {r['name']}" for r in hit), file=sys.stderr)
    return hit[0]


def reg_lines(r):
    return (DATA / 'regs' / r['file']).read_text(encoding='utf-8').split('\n')


def art_no(a):
    return a['no']


def article_blocks(r):
    """[(no, title, [lines])] — 본칙 조 단위. 앞머리(장 제목 등)는 no='' 블록."""
    lines = reg_lines(r); arts = r['articles']
    end = r['buchik_line'] if r['buchik_line'] is not None else len(lines)
    out = []
    for i, a in enumerate(arts):
        stop = arts[i + 1]['line'] if i + 1 < len(arts) else end
        blk = lines[a['line']:stop]
        while blk and (not blk[-1].strip() or re.match(r'^제\s*\d+\s*(장|절|관)(\s*의\s*\d+)?\s', blk[-1].strip())):
            blk.pop()                                   # 다음 장·절 제목은 이 조에 속하지 않음
        out.append((a['no'], a['title'], blk))
    return out, lines[end:]


def cmd_list(args):
    kw = [sq(a) for a in args]
    print('분류번호 | 사규명 | 소관부서 | 시행일 | 이사회결의대상')
    for r in INDEX['regs']:
        s = sq(r['name'] + r['소관부서'] + r['분류'])
        if all(k in s for k in kw):
            print(f"({r['key']}) | {r['name']} | {r['소관부서']} | {r['시행일']} | {r['이사회결의대상']}")
    print(f"\n기준일(빌드): {INDEX['built']} — 사규 {len(INDEX['regs'])}건, 별표·서식 {len(INDEX['atts'])}건")


def opts(args):
    o = {'or': False, 'reg': None, 'att': True, 'max': 30, 'notes': False}; kw = []
    it = iter(args)
    for a in it:
        if a == '--or': o['or'] = True
        elif a == '--no-att': o['att'] = False
        elif a == '--notes': o['notes'] = True
        elif a == '--reg': o['reg'] = next(it)
        elif a == '--max': o['max'] = int(next(it))
        else: kw.append(a)
    return o, kw


def match(text, kw, any_):
    t = sq(text); ks = [sq(k) for k in kw]
    return any(k in t for k in ks) if any_ else all(k in t for k in ks)


def snippet(lines, kw, width=3):
    ks = [sq(k) for k in kw]; keep = []
    for i, ln in enumerate(lines):
        if any(k in sq(ln) for k in ks): keep.append(i)
    if len(lines) <= 14: return [l for l in lines if l.strip()]
    out = [lines[0]] if 0 not in keep else []
    for i in keep[:width * 2]:
        out.append(('   … ' if i else '') + lines[i].strip())
    return out


EXC = re.compile(r'불구하고|생략할 수|생략한다|갈음|준용|위임|적용하지 아니|그러하지 아니|제외한다|우선한다|예외로')


def exc_lines(lines):
    """조문 안에서 예외·생략·위임·준용을 정한 줄만."""
    out, i = [], 0
    while i < len(lines):
        l = lines[i].strip()
        if EXC.search(l):
            out.append(l)
            if re.search(r'다음\s*각\s*호|다음과 같|다음의', l):      # 뒤따르는 호·목까지 함께
                j = i + 1
                while j < len(lines) and re.match(r'^\s*(\d+\.|[가-하]\.|\(\d+\))', lines[j]):
                    if lines[j].strip() not in out: out.append('  ' + lines[j].strip())
                    j += 1
        i += 1
    return out


def name_hits(kw):
    """키워드가 사규 '이름'에 들어 있는 사규 — 본문은 약칭(협의회·위원회 등)을 써서 키워드 검색에 안 걸린다."""
    ks = [sq(k) for k in kw]
    return [r for r in INDEX['regs'] if any(k in sq(r['name']) for k in ks)]


def print_name_hits(kw, shown):
    for r in name_hits(kw):
        blocks, _ = article_blocks(r)
        print(f"★ 「{r['name']}」({r['key']})은 사규명이 키워드와 일치 — 본문이 약칭을 쓰므로 검색에 안 걸린 조가 있다. 답하기 전에 카드와 전문을 읽을 것: card {r['key']} / show {r['key']}")
        print('   조 목록: ' + ', '.join(f"제{no}조({t})" for no, t, _ in blocks))
        for no, t, lines in blocks:
            ex = exc_lines(lines)
            if ex:
                print(f"   ⚠ 제{no}조({t}) 예외·생략·위임 조항:")
                for l in ex[:12]: print('      ' + l[:220])
        print()


def cmd_find(args):
    o, kw = opts(args)
    if not kw: sys.exit('키워드를 주세요')
    regs = [resolve(o['reg'])] if o['reg'] else INDEX['regs']
    n = 0
    if not o['reg']: print_name_hits(kw, None)
    for r in regs:
        blocks, buchik = article_blocks(r)
        for no, title, lines in blocks:
            if match('\n'.join(lines), kw, o['or']):
                n += 1
                if n > o['max']: break
                print(f"■ 「{r['name']}」 제{no}조({title})  [({r['key']}), 소관 {r['소관부서']}]")
                for l in snippet(lines, kw): print('   ' + l.strip())
                print()
        if n > o['max']: break
        if buchik and match('\n'.join(buchik), kw, o['or']) and n <= o['max']:
            hits = [l.strip() for l in buchik if any(sq(k) in sq(l) for k in kw)][:4]
            print(f"■ 「{r['name']}」 부칙·말미  [({r['key']})]"); [print('   ' + h) for h in hits]; print()
    if o['att'] and n <= o['max']:
        keys = {r['key'] for r in regs}
        for a in INDEX['atts']:
            if a['key'] not in keys: continue
            text = (DATA / 'atts' / a['file']).read_text(encoding='utf-8')
            if match(a['name'] + text, kw, o['or']):
                n += 1
                if n > o['max']: break
                hits = [l.strip() for l in text.split('\n') if any(sq(k) in sq(l) for k in kw)][:5]
                print(f"□ 별표·서식 ({a['key']}) {a['name']}"); [print('   ' + h[:200]) for h in hits]; print()
    if n == 0: print('해당 없음 — 키워드를 줄이거나 --or 로 다시 검색. 사규에 없는 것은 "사규에 규정 없음"으로 답할 것.')
    elif n > o['max']: print(f'… 결과가 {o["max"]}건을 넘음. 키워드를 추가하거나 --reg 로 좁힐 것.')


def want_no(s):
    s = sq(s).replace('제', '').replace('조', '')
    return s


def cmd_show(args):
    if not args: sys.exit('사규를 지정하세요')
    r = resolve(args[0]); nos = [want_no(a) for a in args[1:]]
    print(f"「{r['name']}」 ({r['key']}) — 소관 {r['소관부서']}, 시행 {r['시행일']}, 원본 {r['source']}\n")
    if not nos:
        print('\n'.join(reg_lines(r))); return
    blocks, _ = article_blocks(r); found = False
    for no, title, lines in blocks:
        if no in nos:
            found = True; print('\n'.join(l for l in lines if l.strip())); print()
    if not found: print('해당 조 없음. toc 로 조 목록 확인.')


def cmd_toc(args):
    r = resolve(args[0])
    print(f"「{r['name']}」 ({r['key']}) — 소관 {r['소관부서']}, 시행 {r['시행일']}")
    for a in r['articles']: print(f"  제{a['no']}조({a['title']})")
    att = [a for a in INDEX['atts'] if a['key'] == r['key']]
    if att:
        print('  별표·서식:'); [print('   -', a['name']) for a in att]


def cmd_att(args):
    o, kw = opts(args)
    hit = [a for a in INDEX['atts'] if match(f"({a['key']}) {a['name']}", kw, False)]
    if not hit: print('해당 별표·서식 없음'); return
    if len(hit) > 1:
        for a in hit: print(f"({a['key']}) {a['name']}")
        print(f'\n{len(hit)}건 — 키워드를 더 주면 전문을 출력'); return
    a = hit[0]
    print(f"({a['key']}) {a['name']} — 원본 {a['source']}\n")
    print((DATA / 'atts' / a['file']).read_text(encoding='utf-8'))


def cmd_jg(args):
    o, kw = opts(args)
    J = json.load(open(DATA / 'jeongyeol.json', encoding='utf-8'))
    print(f"「직무전결규정」 [별표 제1호] 직무전결기준표 — {J['머리']}")
    if o['notes'] or not kw:
        print('※ 적용사항'); [print('  ' + n) for n in J['적용사항']]; print('  ' + J['범례'])
        if not kw: return
    n = 0
    for r in J['rows']:
        if match(r['분야'] + r['업무'] + r['비고'], kw, o['or']):
            n += 1
            who = ', '.join(f"{m['직위']}{'(소관부서 전결 ●)' if m['표시'] == '●' else ''}" for m in r['전결']) or '— (하위 행 또는 비고 참조)'
            print(f"· ({r['분야']}) {r['업무']}\n    전결: {who}" + (f"\n    비고: {r['비고']}" if r['비고'] else ''))
    if n == 0:
        print('해당 행 없음 → 적용사항 1: 기준표에 명시되지 않은 사항은 대표이사 결재 사항. (다른 규정·시행문서에 별도 정함이 있는지 find 로 확인)')


def cmd_exc(args):
    o, kw = opts(args)
    if not kw: sys.exit('키워드를 주세요')
    regs = [resolve(o['reg'])] if o['reg'] else INDEX['regs']
    n = 0
    for r in regs:
        blocks, _ = article_blocks(r)
        for no, title, lines in blocks:
            ex = exc_lines(lines)
            if not ex: continue
            # 조 전체가 키워드와 관련되거나(사규명 포함) 예외 줄 자체에 키워드가 있는 경우
            scope = r['name'] + '\n'.join(lines)
            if match(scope, kw, o['or']):
                n += 1
                print(f"⚠ 「{r['name']}」 제{no}조({title})  [({r['key']})]")
                for l in ex[:14]: print('   ' + l[:240])
                print()
    if n == 0: print('해당 예외·생략·위임 조항 없음 (키워드를 줄이거나 --or)')


def cmd_where(args):
    o, kw = opts(args)
    if not kw: sys.exit('키워드를 주세요')
    for r in INDEX['regs']:
        blocks, _ = article_blocks(r)
        hit = [f"제{no}조({t})" for no, t, lines in blocks if match('\n'.join(lines), kw, o['or'])]
        tag = ' ★사규명 일치(전문 확인)' if any(sq(k) in sq(r['name']) for k in kw) else ''
        if hit or tag: print(f"({r['key']}) 「{r['name']}」{tag}: " + (', '.join(hit) or '—'))
    for a in INDEX['atts']:
        text = (DATA / 'atts' / a['file']).read_text(encoding='utf-8')
        if match(a['name'] + text, kw, o['or']): print(f"   □ 별표·서식 ({a['key']}) {a['name']}")


CARDS = Path(__file__).resolve().parents[1] / 'references' / 'cards'


def card_path(r):
    return CARDS / (r['file'][:-4] + '.md')


def cmd_card(args):
    r = resolve(args[0]); f = card_path(r)
    if not f.exists(): print('카드 없음 — show 로 원문을 읽을 것'); return
    print(f.read_text(encoding='utf-8'))
    print('\n※ 카드는 지도다. 답의 근거로 쓰는 조문은 show 로 원문을 다시 확인하고 인용할 것.')


def cmd_map(args):
    o, kw = opts(args)
    if not kw: sys.exit('키워드를 주세요')
    ks = [sq(k) for k in kw]; res = []
    for r in INDEX['regs']:
        f = card_path(r)
        if not f.exists(): continue
        sec = ''; hits = []; found = set()
        for ln in f.read_text(encoding='utf-8').split('\n'):
            if ln.startswith('## '): sec = ln[3:5].strip('. '); continue
            hk = [k for k in ks if k in sq(ln)]
            if hk:
                found |= set(hk); hits.append((sec, ln.strip()))
        nm = [k for k in ks if k in sq(r['name'])]
        found |= set(nm)
        if found: res.append((len(found), len(hits), r, hits))
    res.sort(key=lambda x: (-x[0], -x[1]))
    if not res: print('관련 카드 없음 — 유의어로 다시 검색'); return
    print(f'키워드 {len(ks)}개 중 일치 개수 순. [2]=예외·생략 [3]=다른 사규와의 연결 [5]=찾는 말. 상위 사규는 card 로 읽고, 근거 조문은 show 로 원문 확인.\n')
    for nf, nh, r, hits in res[:o['max'] if o['max'] != 30 else 15]:
        print(f"■ ({r['key']}) 「{r['name']}」 — 일치 {nf}/{len(ks)}, {nh}줄  [소관 {r['소관부서']}]")
        pri = sorted(hits, key=lambda h: {'2': 0, '3': 1, '5': 2, '1': 3}.get(h[0], 4))
        for sec, ln in pri[:5]: print(f'   [{sec}] {ln[:230]}')
        print()


def cmd_cite(args):
    r = resolve(args[0])
    rows = list(csv.DictReader(open(DATA / '인용관계.csv', encoding='utf-8-sig')))
    print(f"「{r['name']}」을 인용하는 사규 (개정 시 연동 검토 대상):")
    for x in rows:
        if x['인용되는사규'] == r['key']: print(f"  ({x['인용하는사규']}) {x['인용하는사규명']} — {x['횟수']}회")
    print(f"\n「{r['name']}」이 인용하는 사규:")
    for x in rows:
        if x['인용하는사규'] == r['key']: print(f"  ({x['인용되는사규']}) {x['인용되는사규명']} — {x['횟수']}회")
    print('\n※ 인용관계는 「사규명」 표기 기준 자동 추출이라 누락이 있을 수 있음. find 로 사규명·조문 키워드를 한 번 더 검색할 것.')


def cmd_ledger(args):
    kw = args
    for x in csv.DictReader(open(DATA / '대장목록.csv', encoding='utf-8-sig')):
        if match(''.join(x.values()), kw, False) if kw else True:
            print(f"({x['분류번호']}) {x['사규명']} | {x['대장명']} | {x['근거조문']} | {x['양식']} | {x['상태']}")


CMDS = {'list': cmd_list, 'find': cmd_find, 'show': cmd_show, 'toc': cmd_toc, 'att': cmd_att,
        'jg': cmd_jg, 'cite': cmd_cite, 'ledger': cmd_ledger, 'exc': cmd_exc, 'where': cmd_where, 'card': cmd_card, 'map': cmd_map}

if __name__ == '__main__':
    if len(sys.argv) < 2 or sys.argv[1] not in CMDS:
        print(__doc__); sys.exit(0)
    try:
        CMDS[sys.argv[1]](sys.argv[2:])
    except BrokenPipeError:          # head 등으로 출력이 끊긴 경우
        sys.stderr.close()
