# -*- coding: utf-8 -*-
"""사규관리규정 종합 개정안 신구조문대비표(.docx) 생성.
표현 정비(법제처)는 자동 치환, 핵심 조문(목 삭제·효력순위 통합·시스템 현행화·정본 신설 등)은 권장안 직접 반영.
"""
import re, os
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT="맑은 고딕"; GREY=RGBColor(0x33,0x33,0x33); YELLOW="FFB800"; HDR=RGBColor(0xFF,0xFF,0xFF)

def kfont(run,size=9,bold=False,color=GREY):
    run.font.size=Pt(size); run.font.bold=bold; run.font.name=FONT; run.font.color.rgb=color
    rPr=run._element.get_or_add_rPr(); rf=rPr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rPr.append(rf)
    for a in ('w:eastAsia','w:ascii','w:hAnsi'): rf.set(qn(a),FONT)

def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd')
    sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),hexc); tcPr.append(sh)

def cell_text(cell,head,body,size=9,head_color=None):
    cell.text=""
    p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(2); p.paragraph_format.line_spacing=1.15
    if head:
        r=p.add_run(head); kfont(r,size=size,bold=True,color=head_color or GREY)
        if body: p.add_run("\n")
    if body:
        for i,line in enumerate(body.split("\n")):
            if i>0: p.add_run("\n")
            r=p.add_run(line); kfont(r,size=size)

# ---------- 현행 조문 로드 ----------
t=open('/tmp/kbci_reg.txt').read()
parts=re.split(r'(제\s?\d+\s?조(?:의\s?\d+)?\s*\([^)]*\))', t)
arts={}; order=[]
i=1
while i<len(parts):
    head=re.sub(r'\s+',' ',parts[i].strip())
    if '부 칙' in head: break
    body=re.sub(r'\s+',' ',(parts[i+1] if i+1<len(parts) else "")).strip()
    key=re.sub(r'제\s*(\d+)\s*조의\s*(\d+)',r'제\1조의\2',head); key=re.sub(r'제\s*(\d+)\s*조',r'제\1조',key)
    arts[key]=body; order.append(key); i+=2

def num(key):
    m=re.match(r'제(\d+)조(?:의(\d+))?',key); return (int(m.group(1)), int(m.group(2) or 0)) if m else (999,0)

def revise(s):
    s=s.replace("이라 함은","이란").replace("라 함은","란")
    for a,b in [("에 의하여야","에 따라야"),("에 의하여","에 따라"),("에 의한다","에 따른다"),("에 의할","에 따를"),("에 의한 ","에 따른 ")]:
        s=s.replace(a,b)
    s=s.replace("하여서는 아니된다","해서는 안 된다").replace("되지 아니하는","되지 않는").replace("하지 아니한","하지 않은").replace("아니된다","안 된다")
    s=re.sub(r'아니한(?=\s|$|다|\.)','않은',s)
    s=s.replace("개폐","개정·폐지")
    for a,b in [("시달절차","공포절차"),("시달한다","공포한다"),("시달하","공포하"),("시달일","공포일"),("시달된","공포된"),("시달","공포")]:
        s=s.replace(a,b)
    s=s.replace("회송한다","돌려보낸다").replace("당해","해당").replace("등사","복사").replace("견양","견본")
    s=s.replace("등록 필","등록을 마친").replace("기타","그 밖의").replace("ㆍ","·")
    return s

# ---------- 권장안(핵심 조문 직접 작성) ----------
NEW={
 "제4조(규정화시 유의사항)":("제4조(규정화 시 유의사항) 규정화 시에는 계속성·탄력성·균형성·적용성·능률성을 고려하여야 한다.",
   "[경량화] 5개 호의 추상 설명을 1개 문장으로 축약"),
 "제5조(규정의 효력순위)":("제5조(규정의 효력순위) ① 외규는 규정에 우선하여 효력을 가진다. ② 규정은 정관, 이사회규정, 이사회 부의대상 규정, 대표이사 전결규정, 그 밖의 일반규정의 순위로 하위규정을 구속하고, 상위규정에 저촉되는 하위규정의 조항은 효력을 상실한다. ③ 동일 순위의 규정 사이에서 기존 조항이 새로 제정·개정된 조항과 저촉되는 부분은 특별한 정함이 없는 한 새 규정의 시행과 동시에 효력을 잃는다. ④ 지침은 규정에 저촉되지 않는 범위에서 규정과 동일한 효력을 가진다.",
   "[정합성·경량화] 제10조의2와 중복·불일치 → 제5조로 통합(위계 목록 일원화). ※위계 목록 [확인]"),
 "제8조(등록)":("제8조(등록) 규정의 제정·개정·폐지안이 결재되었을 때에는 주무부서장은 「사규관리시스템」에 등록하며, 제30조의2에 따라 문서번호(구분-연도-순번)를 부여한다. 다만, 감독기관의 협의가 필요한 때에는 그 동의서 사본을 첨부한다.",
   "[시스템·연계] ‘문서관리규정 연계’→사규관리시스템 등록·제30조의2 문서번호 부여로 현행화, 개폐→제정·개정·폐지"),
 "제9조(시달절차)":("제9조(공포 절차) ① 주무부서장은 등록을 마친 사규를, 규정은 대표이사 명의로, 지침은 전결권자 명의로 공포한다. ② 사규를 공포하는 경우에는 제정·개정·폐지의 사유를 간략히 명시하여야 하며, 필요에 따라 다음 사항을 따로 통지할 수 있다. 1. 규정의 해설 2. 규정의 시행에 필요한 조치와 그 밖의 준비사항",
   "[정합성·시스템·표현] 비문(‘규정을 규정은’) 정리, 시달→공포, 등록 필→등록을 마친, 개폐→제정·개정·폐지"),
 "제10조의2(효력순위)":("〈삭제〉",
   "[경량화·정합성] 제5조와 중복 → 삭제(제5조로 통합)"),
 "제15조(규정집 간행)":("제15조(등재 및 열람) ① 주무부서장은 규정을 「사규관리시스템」에 등재하여 관리하며, 임직원은 사규관리시스템을 통하여 열람한다. ② 현행 규정의 세부사항은 『별표 제1호』의 사규목차에 따른다.",
   "[시스템·경량화] 전산시스템 등→사규관리시스템, 종이 ‘규정집 간행’ 완화, 조명 현행화"),
 "제18조(본칙의 구성)":("제18조(본칙의 구성) ① 본칙은 「조」로 구성한다. ② 여러 조문으로 된 규정은 보통 「장」으로 나누며, 소분류가 더 필요할 때에는 절·관의 순으로 사용하고, 특히 조문이 많을 때에는 장의 분류 위에 「편」을 둔다. ③ 편·장·절·관에는 각각 일련번호와 제목을 붙인다. 다만, 장·절·관의 일련번호는 이들이 속하는 편·장·절이 바뀔 때마다 1부터 시작한다.",
   "[정합성] ‘목’은 조 하위 단위 → 조 위 묶음(편·장·절·관)에서 삭제"),
 "제20조(조)":("제20조(조) ① 조에는 제1조부터 시작하는 일련번호를 붙이되 편·장·절·관이 바뀌더라도 그 번호순을 바꾸지 않는다. ② 조에는 조문의 내용을 요약한 조명을 소괄호 안에 표시한다.",
   "[정합성·표현] ‘목’ 삭제, ‘아니한다’→‘않는다’"),
}
# 신설(정본)
EXTRA=[("〈신설〉","제○조(정본) 「사규관리시스템」에 등재된 사규를 정본(正本)으로 하며, 그 내용이 인쇄물 등 다른 사본과 다를 때에는 정본에 따른다.","[시스템·신설] ※정책 [확인]")]

# ---------- 변경 조문 수집 ----------
entries=[]  # (head, cur, new, note)
for key in order:
    cur=key+" "+arts[key]
    if key in NEW:
        new,note=NEW[key]; entries.append((key,cur,new,note))
    else:
        nv=revise(cur)
        if nv!=cur:
            entries.append((key,cur,nv,"[표현] 법제처 권고 표현 정비"))
for cur,new,note in EXTRA:
    entries.append(("〈신설〉",cur,new,note))

# 헤드 표기 정규화(조번호 띄어쓰기)
def norm_head(h):
    h=re.sub(r'제\s*(\d+)\s*조의\s*(\d+)',r'제\1조의\2',h); return re.sub(r'제\s*(\d+)\s*조',r'제\1조',h)

# ---------- docx ----------
doc=Document(); sec=doc.sections[0]
sec.page_width,sec.page_height=Mm(297),Mm(210)  # 가로(A4 landscape)
sec.top_margin=sec.bottom_margin=Mm(15); sec.left_margin=sec.right_margin=Mm(15)
st=doc.styles['Normal']; st.font.name=FONT; st.font.size=Pt(9); st.element.rPr.rFonts.set(qn('w:eastAsia'),FONT)

p=doc.add_paragraph(); p.alignment=AL.CENTER
r=p.add_run("사규관리규정 신·구조문대비표 (종합 개정안)"); kfont(r,size=15,bold=True)
p2=doc.add_paragraph(); r=p2.add_run("※ [표현]법제처 표현정비 · [정합성]논리·일관성 · [경량화]현대화 · [시스템]사규관리시스템 반영 · [확인]정책 결정 필요")
kfont(r,size=8,color=RGBColor(0x88,0x88,0x88)); p2.paragraph_format.space_after=Pt(6)

tbl=doc.add_table(rows=1,cols=3); tbl.style='Table Grid'; tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
widths=[Mm(128),Mm(128),Mm(60)]
hdr=tbl.rows[0].cells
for c,txt in zip(hdr,["현 행","개 정 (안)","비 고"]):
    shade(c,YELLOW); c.paragraphs[0].alignment=AL.CENTER
    rr=c.paragraphs[0].add_run(txt); kfont(rr,size=10,bold=True,color=GREY)
for head,cur,new,note in entries:
    row=tbl.add_row().cells
    cell_text(row[0],"",cur,9)
    delete = new.startswith("〈삭제〉")
    cell_text(row[1],"",new,9, head_color=RGBColor(0xC0,0x00,0x00) if (delete or new.startswith("〈신설〉") or head=="〈신설〉") else None)
    cell_text(row[2],"",note,8)
for c,w in zip(tbl.columns,widths):
    for cell in c.cells: cell.width=w

out="사규정비/사규관리규정_신구조문대비표_종합개정안.docx"
os.makedirs(os.path.dirname(out),exist_ok=True); doc.save(out)
print("생성:",out,"| 변경/신설 조문:",len(entries))
