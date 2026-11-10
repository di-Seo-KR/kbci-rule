# -*- coding: utf-8 -*-
"""kb-sagyu 스킬 패키지 빌드.

입력: kb-sagyu/data(코퍼스·목록·인용관계·대장), attachments(직무전결기준표 xlsx), drafts(예산과목표·계정과목 MD),
      kb-sagyu/skill_src/(SKILL.md, sagyu.py — 사람이 고치는 원본)
출력: kb-sagyu/skill/kb-sagyu/ (스킬 폴더), kb-sagyu/skill/kb-sagyu.zip (업로드용)

사규가 개정되면: rebuild.py → build_budget_accounts_md.py·build_account_chart_md.py → 이 스크립트 순으로 실행.
"""
import csv, json, re, shutil, zipfile, datetime
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'skill_src'
OUT = ROOT / 'skill' / 'kb-sagyu'
ZIP = ROOT / 'skill' / 'kb-sagyu.zip'

ART = re.compile(r'^제\s*(\d+)\s*조(?:\s*의\s*(\d+))?\s*[\(（]')
BUCHIK = re.compile(r'^부\s*칙')


def reg_key(fname):
    m = re.match(r'\((\d+-\d+)\)', fname)
    return m.group(1) if m else ''


def clean_name(fname):
    s = re.sub(r'^\(\d+-\d+\)[_ ]*', '', fname)
    s = re.sub(r'\.(docx|xlsx)$', '', s)
    s = re.sub(r'_\d{6}$', '', s)
    return s.strip()


def split_articles(text):
    """본칙을 조 단위로 자른다. 반환: (articles[{no, title, start_line}], 부칙 시작 줄)."""
    lines = text.split('\n')
    arts, buchik = [], None
    for i, ln in enumerate(lines):
        s = ln.strip()
        if buchik is None and BUCHIK.match(s) and arts:
            buchik = i
        if buchik is None:
            m = ART.match(s)
            if m:
                no = m.group(1) + ('의' + m.group(2) if m.group(2) else '')
                t = re.match(r'^제\s*\d+\s*조(?:\s*의\s*\d+)?\s*[\(（]([^)）]*)', s)
                arts.append({'no': no, 'title': t.group(1).strip() if t else '', 'line': i})
    return arts, buchik


def build_jeongyeol():
    """직무전결기준표 xlsx → 행 단위 JSON (업무 경로 + 전결권자 + 비고)."""
    f = next((ROOT / 'attachments').glob('(2-2)*직무전결기준표*.xlsx'))
    ws = openpyxl.load_workbook(f, data_only=True).worksheets[0]
    rows = [[('' if x is None else str(x).strip()) for x in r[:6]] for r in ws.iter_rows(values_only=True)]
    notes = [r[0] for r in rows[:12] if re.match(r'^\d\.', r[0])]
    head = rows[0][0]
    cols = ['점포장', '부서장/본부장', '집행임원', '대표이사']
    out, cat, stack = [], '', []

    def level(s):
        if re.match(r'^\d+\.', s): return 1
        if re.match(r'^[가-하]\.', s): return 2
        if re.match(r'^\(\d+\)', s): return 3
        if re.match(r'^\([가-하]\)', s): return 4
        return 5

    for i, r in enumerate(rows, 1):
        name = re.sub(r'\s+', ' ', r[0])
        if not name or i <= 11:
            continue
        if re.match(r'^\(.+\)$', name) and not any(r[1:5]):
            cat, stack = name.strip('()'), []
            if r[5]:
                out.append({'row': i, '분야': cat, '업무': '(분야 공통)', '전결': [], '비고': re.sub(r'\s+', ' ', r[5])})
            continue
        lv = level(name)
        stack = [x for x in stack if x[0] < lv] + [(lv, name)]
        marks = []
        for c, v in zip(cols, r[1:5]):
            if v:
                marks.append({'직위': c if v in ('○', '●') else v, '표시': v if v in ('○', '●') else '○'})
        out.append({'row': i, '분야': cat, '업무': ' > '.join(x[1] for x in stack),
                    '전결': marks, '비고': re.sub(r'\s+', ' ', r[5])})
    return {'출처': f.name, '머리': head, '적용사항': notes,
            '범례': '○ = 해당 직위 전결, ● = 「직제규정」상 해당 업무 소관부서의 전결(그 열의 직위). 전결 표시가 없는 행은 하위 행을 본다.',
            'rows': out}


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    (OUT / 'data' / 'regs').mkdir(parents=True)
    (OUT / 'data' / 'atts').mkdir(parents=True)
    (OUT / 'scripts').mkdir()
    (OUT / 'references').mkdir()

    meta = {r['분류번호']: r for r in csv.DictReader(open(ROOT / 'data' / '사규목록.csv', encoding='utf-8-sig'))}
    corpus = list(json.load(open(ROOT / 'data' / 'corpus92.json', encoding='utf-8')).values())[0]
    atts = list(json.load(open(ROOT / 'data' / 'corpus_attachments.json', encoding='utf-8')).values())[0]

    index = {'built': datetime.date.today().isoformat(), 'regs': [], 'atts': []}
    for it in corpus:
        key = reg_key(it['file']); name = clean_name(it['file'])
        text = it['text'].strip() + '\n'
        fn = f'({key})_{name}.txt'
        (OUT / 'data' / 'regs' / fn).write_text(text, encoding='utf-8')
        arts, buchik = split_articles(text)
        m = meta.get(key, {})
        index['regs'].append({'key': key, 'name': name, 'file': fn, 'source': it['file'],
                              '분류': m.get('분류', ''), '소관부서': m.get('소관부서', ''), '시행일': m.get('시행일', ''),
                              '이사회결의대상': m.get('이사회결의대상', ''),
                              'articles': arts, 'buchik_line': buchik})
    for it in atts:
        key = reg_key(it['file']); name = clean_name(it['file'])
        fn = f'({key})_{name}.txt'.replace('/', '-')
        (OUT / 'data' / 'atts' / fn).write_text(it['text'].strip() + '\n', encoding='utf-8')
        index['atts'].append({'key': key, 'name': name, 'file': fn, 'source': it['file']})
    index['regs'].sort(key=lambda r: [int(x) for x in r['key'].split('-')])
    index['atts'].sort(key=lambda r: ([int(x) for x in r['key'].split('-')], r['name']))

    json.dump(index, open(OUT / 'data' / 'index.json', 'w', encoding='utf-8'), ensure_ascii=False)
    json.dump(build_jeongyeol(), open(OUT / 'data' / 'jeongyeol.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
    for f in ['사규목록.csv', '인용관계.csv', '대장목록.csv']:
        shutil.copy(ROOT / 'data' / f, OUT / 'data' / f)
    for f in ['예산과목표.md', '계정과목.md']:
        shutil.copy(ROOT / 'drafts' / f, OUT / 'references' / f)
    for f in sorted((ROOT / 'references').glob('0*.md')):
        shutil.copy(f, OUT / 'references' / f.name)
    shutil.copy(SRC / 'sagyu.py', OUT / 'scripts' / 'sagyu.py')
    for f in SRC.glob('*.md'):
        dst = OUT / f.name if f.name == 'SKILL.md' else OUT / 'references' / f.name
        s = f.read_text(encoding='utf-8').replace('{{BUILT}}', index['built']) \
             .replace('{{N_REGS}}', str(len(index['regs']))).replace('{{N_ATTS}}', str(len(index['atts'])))
        dst.write_text(s, encoding='utf-8')

    if ZIP.exists():
        ZIP.unlink()
    with zipfile.ZipFile(ZIP, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(OUT.rglob('*')):
            if f.is_file():
                z.write(f, Path('kb-sagyu') / f.relative_to(OUT))
    size = sum(f.stat().st_size for f in OUT.rglob('*') if f.is_file())
    print(f'regs {len(index["regs"])}, atts {len(index["atts"])}, folder {size/1e6:.2f} MB, zip {ZIP.stat().st_size/1e6:.2f} MB')


if __name__ == '__main__':
    main()
