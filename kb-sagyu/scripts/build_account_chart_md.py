# -*- coding: utf-8 -*-
"""「회계규정」 [별표 제1호] 계정과목 해설(xlsx) → Markdown 변환 + 예산과목표 대조.

입력: kb-sagyu/attachments/(4-3) 별표 제1호_계정과목 해설.xlsx
      kb-sagyu/attachments/(4-1) 별표 제1호_예산과목표.xlsx (대조용)
출력: kb-sagyu/drafts/계정과목.md

시트: 자산계정 / 부채자본(부채+자본) / 수익계정 / 비용계정.
열: A=계정코드, B~F=관·항·목·절·세과목(들여쓰기 위치가 곧 레벨), G=통합 COA, H=집계계정.
"""
import re, sys, collections
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'attachments' / '(4-3) 별표 제1호_계정과목 해설.xlsx'
OUT = ROOT / 'drafts' / '계정과목.md'
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_budget_accounts_md as budget  # noqa: E402

LV = ['관', '항', '목', '절', '세과목']


def norm(s):
    return re.sub(r'\s+', '', s or '')


def parse_sheet(ws, section):
    out, cur, prev_code = [], [''] * 5, None
    rows = list(ws.iter_rows(min_row=5, values_only=True))
    for idx, r in enumerate(rows):
        r = list(r) + [None] * 8
        code = str(r[0]).strip() if r[0] is not None else ''
        names = [norm(str(x)) if x is not None else '' for x in r[1:6]]
        coa = str(r[6]).strip() if r[6] is not None else ''
        agg = str(r[7]).strip() if r[7] is not None else ''
        if code.startswith('(') and code.endswith('계정)'):      # '(자본계정)' 같은 구간 표시
            section = code.strip('()').replace('계정', '')
            cur = [''] * 5
            continue
        if code.startswith('계') or names[0] == '관':              # 반복 머리행
            continue
        if not code and not any(names):
            continue
        lvl = next((i for i, n in enumerate(names) if n), None)
        if lvl is None:                                            # 명칭 빈칸 행
            if code == prev_code:
                continue                                           # 같은 코드 중복 빈칸 → 버림
            nxt = next((list(x)[1:6] for x in rows[idx + 1:] if any(list(x)[1:6])), None)
            nl = next((i for i, n in enumerate(nxt or []) if n), 4)
            lvl, name = max(nl - 1, 0), '(명칭 없음)'
        else:
            name = names[lvl]
        contra = name.startswith('(') and name != '(명칭 없음)'   # 차감계정: 상위 경로로 쓰지 않음
        path = cur[:lvl] + [name]
        if not contra:
            cur[lvl] = name
            for j in range(lvl + 1, 5):
                cur[j] = ''
        out.append({'section': section, 'code': code, 'level': lvl, 'name': name, 'contra': contra,
                    'path': path, 'coa': coa, 'agg': agg})
        prev_code = code
    # 리프: 자기 경로를 접두로 갖는 더 긴 (비차감) 경로가 없으면 리프
    prefixes = set()
    for x in out:
        if not x['contra']:
            for k in range(1, len(x['path'])):
                prefixes.add((x['section'],) + tuple(x['path'][:k]))
    for x in out:
        x['leaf'] = ((x['section'],) + tuple(x['path'])) not in prefixes
    return out


def load_accounts():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    sec_of = {'자산계정': '자산', '부채자본': '부채', '수익계정': '수익', '비용계정': '비용'}
    allrows = []
    for ws in wb.worksheets:
        allrows += parse_sheet(ws, sec_of.get(ws.title, ws.title))
    return allrows


def cmp_key(name):
    """대조용 정규화: 공백 제거, 계정 측 구분 접미어 제거."""
    s = norm(name)
    s = re.sub(r'\((물건비|ACS)\)', '', s)
    return s


def compare(acc, bud):
    # 예산과목표 범위에 대응하는 계정: 비용 > 경비, 자산 > 비유동자산
    def in_scope(a):
        return (a['section'] == '비용' and a['path'][0] == '경비') or \
               (a['section'] == '자산' and a['path'][0] == '비유동자산')
    scope = [a for a in acc if in_scope(a) and not a['contra']]
    acc_all = {cmp_key(n) for a in scope for n in a['path']}
    bud_all = {cmp_key(x) for r in bud for x in (r['관'], r['항'], r['목'], r['절'], r['세과목']) if x}
    only_acc = [a for a in scope if a['leaf'] and cmp_key(a['name']) not in bud_all]
    only_bud = []
    for r in bud:
        leaf = r['세과목'] or r['절'] or r['목']
        if cmp_key(leaf) not in acc_all:
            only_bud.append(r)
    return only_acc, only_bud


ACCT_ONLY = re.compile(r'사용권|감가상각|확정급여|사외적립|복구충당|사용제한|부가가치세|퇴직급여\(제수당\)|리스료')


def nature(a):
    return '결산·회계처리성' if ACCT_ONLY.search(''.join(a['path'])) else '집행성'


def md_table_rows(rows):
    L = ['| 계정코드 | 레벨 | 관 | 항 | 목 | 절 | 세과목 | 통합COA | 집계계정 |',
         '|---|---|---|---|---|---|---|---|---|']
    for a in rows:
        p = a['path'] + [''] * (5 - len(a['path']))
        L.append('| {} | {} | {} | {} | {} | {} | {} | {} | {} |'.format(
            a['code'], LV[a['level']], *[x or '' for x in p], a['coa'], a['agg']))
    return L


def build():
    acc = load_accounts()
    ws = openpyxl.load_workbook(budget.SRC, data_only=True).worksheets[0]
    bud = budget.parse(ws)
    for i, r in enumerate(bud, 1):
        r['no'] = i
    only_acc, only_bud = compare(acc, bud)

    cnt = collections.Counter(a['section'] for a in acc)
    leaf_cnt = collections.Counter(a['section'] for a in acc if a['leaf'])
    L = []
    L.append('# 계정과목 — 「회계규정」 [별표 제1호] 계정과목 해설')
    L.append('')
    L.append('- 출처: 「회계규정」(2023. 10. 30.) [별표 제1호] 계정과목 해설, 사규관리시스템 등재본 xlsx')
    L.append('- 근거: 「회계규정」 제45조제1항 "회사의 계정과목은 [별표 제1호]에 의한다". 계정과목의 신설·변경·폐지는 회계주무부서장이 총괄하고 기업회계기준 변경·신상품 개발 등에 한하며, 카드사 검토를 거쳐 승인한다(제47조).')
    L.append('- 구조: 10자리 계정코드 + 관 → 항 → 목 → 절 → 세과목. 별표 이름은 \'해설\'이지만 실제로는 설명문 없이 코드와 체계만 있다.')
    L.append('- 통합COA·집계계정: 일부 계정에만 기재된 대응 코드다. 「회계규정」 제45조제2항(카드사 연결재무제표용 계정체계 적용)에 따른 카드사 계정 대응 코드로 보인다.')
    L.append('- 표기 정규화: 명칭의 자간 공백("경      비", "유  치  원")과 셀 안 줄바꿈은 제거했다. 명칭이 빈칸인 행은 "(명칭 없음)"으로 표시했다.')
    L.append('- 예산과목표와의 관계: 예산은 「예산규정」 [별표 제1호] 예산과목표(2023. 5. 3.)로 편성하는데, 이 계정과목 해설이 더 나중 것이고 세과목이 더 많다. 아래 "예산과목표 대조"를 먼저 본다.')
    L.append('')
    L.append('## 규모')
    L.append('')
    L.append('| 구분 | 전체 행 | 최하위(리프) 계정 |')
    L.append('|---|---|---|')
    for s in ['자산', '부채', '자본', '수익', '비용']:
        if cnt[s]:
            L.append(f'| {s} | {cnt[s]} | {leaf_cnt[s]} |')
    L.append('')

    L.append('## 예산과목표 대조 (예산사업 ↔ 세과목 매핑 시 필독)')
    L.append('')
    L.append('대조 범위는 예산과목표가 다루는 경비(비용계정 > 경비)와 비유동자산(자산계정 > 비유동자산)이다. 명칭은 공백과 계정 측 구분 접미어 "(물건비)"·"(ACS)"를 떼고 비교했다. 차감 계정("(감가상각누계액)" 등)은 제외했다.')
    L.append('')
    L.append(f'### 계정에는 있고 예산과목표에는 없는 최하위 과목 ({len(only_acc)}개)')
    L.append('')
    L.append('예산과목표 개정 시 추가 검토 대상이다. 27년 예산사업이 이 계정으로 집행될 예정이면 예산과목표에 세과목(집행구분 포함)을 신설해야 배정·통제가 가능하다.')
    L.append('')
    L.append('성격 열은 작성자 분류다. "집행성"은 부서가 예산을 받아 집행하는 비용, "결산·회계처리성"은 리스(사용권자산)·감가상각 세분·확정급여 퇴직급여 구성요소처럼 결산 분개로 생기는 계정이라 보통 예산 편성 대상이 아니다.')
    L.append('')
    L.append('| 계정코드 | 경로 | 과목 | 성격 |')
    L.append('|---|---|---|---|')
    for a in sorted(only_acc, key=lambda a: (nature(a) != '집행성', a['code'])):
        L.append(f"| {a['code']} | {' > '.join(a['path'][:-1])} | {a['name']} | {nature(a)} |")
    L.append('')
    L.append(f'### 예산과목표에는 있고 계정과목에서 같은 이름을 찾지 못한 과목 ({len(only_bud)}개)')
    L.append('')
    L.append('계정 쪽 명칭이 다르거나 폐지됐을 수 있다. 매핑 시 계정코드를 특정할 수 없으므로 회계 담당과 대응 계정을 확인한다.')
    L.append('')
    if not only_bud:
        L.append('없음. 예산과목표의 모든 최하위 과목은 계정과목 해설에서 같은 이름을 찾을 수 있다(명칭 표기 차이 "개인포상비/개인 포상비", "고정급/고정급(ACS)" 포함).')
    else:
        L.append('| 예산과목표 No. | 경로 | 과목 | 집행구분 |')
        L.append('|---|---|---|---|')
    for r in only_bud:
        leaf = r['세과목'] or r['절'] or r['목']
        path = ' > '.join(x for x in (r['관'], r['항'], r['목'], r['절']) if x and x != leaf)
        L.append(f"| {r['no']} | {path} | {leaf} | {r['집행구분']} |")
    L.append('')
    L.append('### 구조 차이')
    L.append('')
    dup = [c for c, n in collections.Counter(a['code'] for a in acc if a['code']).items() if n > 1]
    for c in dup:
        names = ', '.join(a['name'] for a in acc if a['code'] == c)
        L.append(f'- **원문 계정코드 중복**: {c} — {names}. 별표 원문 오류로 보이며 회계 담당 확인이 필요하다.')
    L.append('- 계정은 복리후생비를 **인건비성복리후생비(인건비 > 5110600000)** 와 **복리후생비(물건비 > 5120100000, 명칭에 "(물건비)" 접미)** 로 나눠 두 벌 운영한다. 예산과목표에는 물건비 > 복리후생비 한 벌만 있다. 같은 경조금·장학금이라도 어느 쪽 계정으로 집행하는지 회계 담당 기준을 확인한다.')
    L.append('- 계정은 자가운전보조비·단체보험을 물건비 쪽(차량비 > 자가운전보조비(물건비), 복리후생비 > 기업단체보험)에 두면서 인건비 > **기타인건비**(인건비성차량비 > 자가운전보조비, 당번근무비, 단체상해보험)에도 둔다. 예산과목표에는 물건비 쪽만 있고, 당번비는 제수수료등의 취급내용에 들어 있다.')
    L.append('- 계정의 잡급 > 채권추심원·임대차조사원·오토론추심원은 절 명칭이 빈칸(5110400400)이다. 예산과목표에서는 일반사무인력 절 아래에 있다.')
    L.append('')

    titles = [('자산', '재무상태표 — 자산'), ('부채', '재무상태표 — 부채'), ('자본', '재무상태표 — 자본'),
              ('수익', '손익계산서 — 수익'), ('비용', '손익계산서 — 비용')]
    for i, (s, t) in enumerate(titles, 1):
        rows = [a for a in acc if a['section'] == s]
        if not rows:
            continue
        L.append(f'## {i}. {t} ({len(rows)}행)')
        L.append('')
        L += md_table_rows(rows)
        L.append('')
    return '\n'.join(L), acc, only_acc, only_bud


if __name__ == '__main__':
    body, acc, oa, ob = build()
    OUT.write_text(body, encoding='utf-8')
    print(OUT, len(body), 'chars;', len(acc), 'accounts;', 'only_acc', len(oa), 'only_bud', len(ob))
