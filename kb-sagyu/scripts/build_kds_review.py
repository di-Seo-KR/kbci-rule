# -*- coding: utf-8 -*-
"""KB데이타시스템 수의계약 사규 검토 보고(워드) 생성."""
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

F='맑은 고딕'
d=Document(); s=d.sections[0]
s.page_width,s.page_height=Mm(210),Mm(297); s.left_margin=s.right_margin=Mm(20); s.top_margin=Mm(18); s.bottom_margin=Mm(15)
st=d.styles['Normal']; st.font.name=F; st.font.size=Pt(10.5); st.element.rPr.rFonts.set(qn('w:eastAsia'),F)

def run(p,t,b=False,sz=None,color=None):
    r=p.add_run(t); r.bold=b; r.font.name=F; r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),F)
    if sz: r.font.size=Pt(sz)
    if color: r.font.color.rgb=RGBColor.from_string(color)
    return r
def para(t='',b=False,sz=None,align=None,before=0,after=2,indent=0):
    p=d.add_paragraph(); f=p.paragraph_format; f.space_before=Pt(before); f.space_after=Pt(after); f.line_spacing=1.15
    if indent: f.left_indent=Mm(indent); f.first_line_indent=Mm(-3)
    if align is not None: p.alignment=align
    if t:
        for i,seg in enumerate(t.split('**')): run(p,seg,b or i%2==1,sz)
    return p
def shade(c,hexc):
    tcPr=c._element.get_or_add_tcPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),hexc); tcPr.append(sh)
def section(no,title):
    t=d.add_table(rows=1,cols=2); t.alignment=WD_TABLE_ALIGNMENT.LEFT; t.autofit=False
    a,b=t.rows[0].cells; a.width=Mm(9); b.width=Mm(161)
    shade(a,'60584C'); shade(b,'FBF4DF')
    pa=a.paragraphs[0]; pa.alignment=WD_ALIGN_PARAGRAPH.CENTER; run(pa,no,True,11,'FFFFFF')
    run(b.paragraphs[0],' '+title,True,12)
    para(after=0,sz=4)
def table(head,rows,widths):
    t=d.add_table(rows=1,cols=len(head)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
    for i,h in enumerate(head):
        c=t.rows[0].cells[i]; shade(c,'F6F4F0'); p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; run(p,h,True,9.5)
    for r in rows:
        cs=t.add_row().cells
        for i,v in enumerate(r):
            p=cs[i].paragraphs[0]; p.paragraph_format.space_after=Pt(0)
            if i==0: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for k,seg in enumerate(v.split('**')): run(p,seg,k%2==1,9.5)
    for row in t.rows:
        for i,w in enumerate(widths): row.cells[i].width=Mm(w)
    para(after=0,sz=4)

t=d.add_table(rows=1,cols=2); t.autofit=False; a,b=t.rows[0].cells; a.width=Mm(130); b.width=Mm(40)
run(a.paragraphs[0],'KB데이타시스템 수의계약 추진 관련 사규 검토',True,16)
pb=b.paragraphs[0]; pb.alignment=WD_ALIGN_PARAGRAPH.RIGHT; run(pb,'경영전략부\n2026. 11. 12. (목)',False,9.5)
bd=OxmlElement('w:tblBorders')
for e in ('top','bottom'):
    x=OxmlElement(f'w:{e}'); x.set(qn('w:val'),'single'); x.set(qn('w:sz'),'12'); x.set(qn('w:color'),'FFBC00'); bd.append(x)
t._element.tblPr.append(bd)
para(after=2,sz=4)

section('Ⅰ','검토 대상')
para('- IT운영부 문의: KB데이타시스템과 수의계약 추진 (총 12억원 내외, 업무용고정자산 취득)',indent=3)
para('- 지급 계획: 착수금 6.5억원 / 잔금 6억원(최종검수 후, `27년 예산 미확정)',indent=3)
para('- 추진 절차(안): 「계약규정」 제54조제2항제8호 수의계약 → 경영협의회 부의, 재무전략협의회 생략',indent=3,after=4)

section('Ⅱ','검토 결과')
para('1. 사규에 부합하는 사항',True,before=2)
table(['항목','판단','근거'],[
 ['수의계약','가능 (KB지주 및 그룹사와의 계약)','「계약규정」 제54조제2항제8호'],
 ['견적서','1개 업체 견적으로 가능','「계약규정」 제61조제1항제1호'],
 ['계약·하자보수보증금','면제 가능. 귀속사유 발생 시 현금 납입 확약서 징구 필요','「계약규정」 제72조제1항제1호·제2항, 제74조제2항·제3항'],
 ['경영협의회 부의','필요 (5천만원 초과 업무용고정자산 취득)','「경영협의회 규정」 제6조제1항제4호'],
 ['재무전략협의회','생략 가능 (경영협의회 결의사항은 부의 생략). 경영전략부 합의는 필요','「재무전략협의회 규정」 제5조제2항제2호, 「재무관리규정」 제4조제1항제4호·제5호'],
],[32,78,60])
para('2. 보완이 필요한 사항',True,before=2)
table(['항목','현행 사규','보완 방안'],[
 ['잔금 6억원\n(`27년 예산)','예산은 다음 회계연도 이월 사용 불가, 이월 시 이사회 승인(「예산규정」 제19조). **총액 일괄 계약 시 예산 없는 금액까지 약정**','① 장기계속계약: 총액 부기, `26년 예산 범위 내 1차 계약(「계약규정」 제109조) 또는 ② 계속비 편성(「예산규정」 제10조). `27년 예산안에 잔금 명시, 안건에 재원 계획 기재'],
 ['착수금 6.5억원\n(검수 전 지급)','대가는 검사 완료 후 지급 원칙(「계약규정」 제93조제1항)','예외 근거를 품의에 명시: 관계회사로서 불가피한 경우(제93조제1항 단서) 또는 용역계약 70% 이내 선지급(제93조제6항). 6.5억원은 약 54%로 한도 이내'],
],[30,68,72])

section('Ⅲ','추가 확인 사항')
for x in [
 '- **사전감사**: 경영협의회 부의 안건으로 결재 전 감사 사전감사 대상. 기안 상단에 근거조항 표기 (「감사규정」 제6조, [별표 제1호] 제3호·제4호 다목)',
 '- **중요한 계약**: 계약금액이 자본금의 5%를 초과하면 「경영협의회 규정」 제6조제1항제3호에도 해당. 안건 근거에 3호·4호 병기',
 '- **IT투자협의회**: 5,000만원 초과 IT투자예산이 소요되는 신규 IT사업은 IT투자협의회 심의·의결 대상(그룹공동IT사업 제외). 전산 관련 사업이면 경영협의회 부의 전에 의결 필요 (「IT투자협의회 운영지침」 제4조·제5조)',
 '- **그룹 내부거래 사전검토**: 주무부서(IT운영부)가 관련 법령상 제한되는 내부거래인지 사전 검토, 불명확하면 법무업무 담당부서에 검토 의뢰. 정상가격 대비 현저히 유리한 조건의 거래 금지 (「그룹 내부거래 운영 규정」 제5조·제11조)',
 '- **이사회 승인 여부**: 상법 제398조의 "이사등과 회사간의 거래"에 해당하면 사전에 이사회 승인 필요. 해당 여부는 위 사전검토에서 확인 (「그룹 내부거래 운영 규정」 제9조·제10조, 「이사회규정」 제9조제1항제11호)',
 '- **지주회사 IT 사전협의**: 그룹사 공동대처·IT 시너지가 예상되는 사항은 지주회사 소관부서와 사전협의, 사전협의 대상 사업은 진행경과·결과 보고 (「그룹 IT 협력 지침」 제3조·제4조)',
 '- **수의계약 소명자료**: 그룹사 확인 자료, 가격 적정성 자료 구비 (「계약규정」 제63조)',
 '- **계약서 기재**: 이행기간, 위험부담, 지체상금, 검수 기준 (「계약규정」 제89조·제99조)',
 '- **검수**: 수요부서장이 검사자 지정, 검사조서 작성. 계약담당 직원은 검사자 불가 (「계약규정」 제91조)',
 '- **집행 결재**: 경영협의회 결의 내용과 같으면 부서장 결재 + 경영전략부 합의, 다르면 대표이사 결재 (직무전결기준표 (예산) 6.나)',
 '- **보안성 검토**: 시스템 구축 시 대상 (「정보보호규정」 제23조·제29조)']:
    para(x,indent=3)
para('끝.',align=WD_ALIGN_PARAGRAPH.RIGHT,before=4)
out='kb-sagyu/drafts/261112_KB데이타시스템 수의계약 사규 검토_v1.2.docx'; d.save(out); print(out)
