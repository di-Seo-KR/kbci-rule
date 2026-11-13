#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""간이 전산 대장(HTML) 생성 — data/대장목록.csv + 등록 서식의 실제 열 구조를 담은 단일 파일.

    python3 kb-sagyu/scripts/build_daejang_html.py [출력경로]

- 외부 CDN·폰트·라이브러리 일절 없음(내부망 오프라인 환경 전제)
- 저장은 브라우저 localStorage. 공유·백업은 JSON 내보내기/가져오기로 한다.
- 대장 열 정의는 attachments/ 의 등록 서식에서 뽑은 실제 항목을 그대로 옮긴 것.
"""
import csv, json, os, sys, datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
T, D, N, S, M, C, TM = 'text', 'date', 'num', 'select', 'memo', 'check', 'time'

# ── 등록 서식이 있는 대장의 실제 열 구조 (키: "분류번호|대장명")
COLS = {
"4-7|통장 반출·입 관리대장": [("일자",D),("계좌번호",T),("사용목적",T),("금액",N),("반출 시간",TM),("반출 인수자",T),("동행자",T),("반납 시간",TM),("반납 인수자",T),("직인 날인자",T),("입회자",T),("비고",M)],
"4-7|인터넷 뱅킹 보안매체 관리대장": [("구분",S,["인증서","OTP 발생기"]),("인증서 ID / OTP S-N",T),("관리자",T),("발급·인수일",D),("인도일",D),("폐기일",D),("부점장 확인",T),("비고",M)],
"4-7|권한변경이력 관리대장": [("제청일",D),("제청부점",T),("제청신청인",T),("제청사유",M),("변경승인자",T),("비고",M)],
"7-6|열쇠 관리대장": [("열쇠번호",T),("반출일자",D),("사용자",T),("담당자 확인",T),("부점장 확인",T),("반납일자",D),("반납확인(부점장)",T),("비고",M)],
"7-6|모바일카드 관리대장": [("구분",S,["임직원","임직원 외"]),("사용자",T),("핸드폰번호",T),("등록일자",D),("등록 담당자 확인",T),("등록 부점장 확인",T),("해제일자",D),("해제 담당자 확인",T),("해제 부점장 확인",T),("비고",M)],
"7-6|실물카드 관리대장": [("구분",S,["반입","반출"]),("카드번호",T),("반출입일자",D),("담당자 확인",T),("부점장 확인",T),("총량",N),("잔량",N),("사용자",T),("비고",M)],
"8-5|고객정보명세 관리대장": [("추출(의뢰)일자",D),("의뢰부서(담당자)",T),("내용",M),("고객수",N),("목적",T),("관련 문서번호",T),("방법",S,["파일","인쇄물","우편함","인편","기타"]),("발송(수령)일자",D),("추출(수령·발송) 확인",T),("고객정보관리책임자 확인",T),("폐기일자",D)],
"8-7|보조기억매체 관리대장": [("연번",N),("관리번호(S/N)",T),("매체형태",S,["USB 메모리","이동형 하드디스크","CD","DVD","디스켓","기타"]),("용도",T),("등록일자",D),("정보보호 관리자",T),("폐기(취소)일자",D),("비고",M)],
"8-7|보조기억매체 점검대장": [("점검일자",D),("보유수량",N),("확인자",T),("서명",T),("비고",M)],
"8-7|보조기억매체 반·출입 대장": [("매체형태",S,["USB 메모리","이동형 하드디스크","CD","DVD","기타"]),("관리번호",T),("사용자",T),("용도",T),("입·출 구분",S,["반입","반출"]),("일시",D),("확인자(서명)",T)],
"8-7|업무용 PC 사용자 관리대장": [("변경일자",D),("PC 관리번호",T),("변경 사유",S,["신규입사","인사발령","고장","반납","폐기","기타"]),("종전 사용자",T),("신규 사용자",T),("하드디스크 포맷 여부",S,["완료","해당없음"]),("확인자",T),("비고",M)],
"8-9|출입자 관리대장": [("날짜",D),("입실시간",TM),("퇴실시간",TM),("출입구역",T),("소속",T),("성명",T),("출입목적",T),("반출입 장비",S,["없음","노트북","태블릿","보조기억매체","기타"]),("장비 용도",T),("반출입 확인",C),("저장매체 사용확인",C),("담당자 확인(서명)",T),("비고",M)],
"8-11|개인영상정보 관리대장": [("번호",N),("구분",S,["이용","제공","열람","파기"]),("일시",D),("파일명·형태",T),("관리책임자 또는 입회자 확인",T),("목적·사유",M),("이용·제공받는 제3자 / 열람 등 요구자",T),("이용·제공 근거",T),("이용·제공 형태",T),("기간",T)],
"8-11|녹취정보 관리대장": [("열람일자",D),("열람자(서명)",T),("입회자(서명)",T),("녹취일시",D),("녹취 전화번호",T),("열람사유",M),("파일명",T),("폐기일자",D),("고객정보관리책임자 또는 입회자 확인",T),("보안의 날 점검일자",D),("고객정보보호관리자 관리실태 확인",T),("비고",M)],
"8-11|생체정보 관리대장": [("순번",N),("지점명",T),("등록일",D),("식별정보",T),("성명",T),("계약구분",T),("등록결재",T),("계약해지일(발령일)",D),("지문정보 삭제여부",S,["삭제","미삭제"]),("지문삭제일",D),("삭제결재",T)],
"8-11|비대면 전자계약 정보 관리대장(전자인장날인 겸용)": [("계약일자",D),("계약 대상자",T),("계약 개인정보 반출여부",S,["여","부"]),("반출일자",D),("반출 대상자",T),("파일명",T),("개인정보 반출 항목",T),("폐기일자(내부망)",D),("폐기일자(외부망)",D),("보관·폐기 여부 점검 기준년월",T),("부점장 서명",T),("비고",M)],
"8-11|출입 관리대장": [("날짜",D),("입실시간",TM),("퇴실시간",TM),("소속",T),("성명",T),("출입목적",T),("담당자 확인",T),("비고",M)],
"9-3|백업 관리 대장": [("백업일자",D),("시스템명",T),("보안등급",T),("내용",M),("관리자 확인",T),("비고",M)],
"9-3|센터 운영실 출입관리 대장": [("년월일",D),("이름",T),("회사명(부서명)",T),("출입사유",T),("입실시간",TM),("퇴실시간",TM),("직원확인",T)],
"9-3|암호화키 관리대장": [("일련번호",T),("구분",S,["키생성","키백업","키폐기"]),("일자",D),("용도",T),("암호키",T),("부서/사용자명",T),("사용자 서명",T),("개인정보보호책임자 서명",T)],
"9-4|유색조 양식 관리대장": [("관리번호",T),("문서유형",T),("보관의뢰자",T),("접수자",T),("접수일자",D),("종료일자",D),("비고",M)],
"9-4|유색조 인영 관리대장": [("순번",N),("문서유형",T),("도장소유자",T),("도장이미지",T),("접수일자",D),("종료일자",D),("비고",M)],
"9-4|약관 및 정형화된 양식 관리대장": [("관리번호",T),("문서유형",T),("보관의뢰자",T),("접수자",T),("접수일자",D),("종료일자",D),("비고",M)],
"9-4|이관 대상문서 관리대장": [("관리번호",T),("문서종류",T),("보관의뢰자",T),("공인전자문서센터 사업자",T),("계약일자",D),("해지일자",D),("비고",M)],
"9-4|대상문서 압축률 관리대장": [("관리번호",T),("보관의뢰자",T),("양식명",T),("압축방식",T),("압축률",N),("종료일자",D),("비고",M)],
"10-3|소셜미디어 관리대장": [("관리번호",T),("개설인",T),("폐쇄일",D),("소속",T),("구분",S,["회사","부서","단체","기타"]),("목적",T),("계정구분",T),("계정명",T),("운영자",T),("담당자(연락처)",T)],
}
# ── 서식 미등록 대장의 기본 열(업무 성격에 맞춘 최소 구성 — 부점에서 조정)
GENERIC = {
"문서발송대장": [("발송일자",D),("문서번호",T),("수신처",T),("제목",T),("발송방법",S,["전자문서","우편","등기","내용증명","인편","기타"]),("직인 날인",C),("발송자",T),("비고",M)],
"문서접수대장": [("접수일자",D),("문서번호",T),("발신처",T),("제목",T),("접수방법",S,["전자문서","우편","팩스","인편","기타"]),("담당자",T),("처리결과",T),("비고",M)],
"보존문서기록대장": [("등록일자",D),("문서번호",T),("문서명",T),("보존기간",S,["1년","3년","5년","10년","준영구","영구"]),("보존시작일",D),("폐기예정일",D),("보관위치",T),("폐기일자",D),("비고",M)],
"인장관리대장": [("인장종류",T),("규격",T),("조제일자",D),("등록번호",T),("관리자",T),("보관장소",T),("폐기일자",D),("비고",M)],
"인장날인대장": [("날인일자",D),("문서번호",T),("문서명",T),("인장종류",T),("날인 건수",N),("신청자",T),("날인자",T),("비고",M)],
}
DEFAULT = [("일자",D),("내용",M),("담당자",T),("확인자",T),("비고",M)]


def norm(cols):
    out = []
    for c in cols:
        item = {"name": c[0], "type": c[1]}
        if len(c) > 2: item["opts"] = c[2]
        out.append(item)
    return out


def build_defs():
    rows = list(csv.DictReader(open(os.path.join(BASE, 'data/대장목록.csv'), encoding='utf-8-sig')))
    defs = []
    for i, r in enumerate(rows):
        key = '%s|%s' % (r['분류번호'], r['대장명'])
        cols = COLS.get(key) or GENERIC.get(r['대장명']) or DEFAULT
        defs.append({
            "id": "d%02d" % i, "code": r['분류번호'], "reg": r['사규명'], "name": r['대장명'],
            "art": r['근거조문'], "form": r['양식'], "state": r['상태'],
            "cols": norm(cols),
            "src": "서식" if key in COLS else ("기본" if r['대장명'] in GENERIC else "표준"),
        })
    return defs


HTML = r'''<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>KB신용정보 간이 전산 대장</title>
<style>
*{box-sizing:border-box}
:root{--bd:#d8dce3;--bg:#f5f6f8;--ink:#1c1f24;--mut:#6b7280;--pri:#1a4fa0;--pri-l:#eaf1fb;--warn:#b03030;--ok:#0f7b3f}
html,body{margin:0;height:100%}
body{font-family:"맑은 고딕","Malgun Gothic",AppleGothic,sans-serif;font-size:13px;color:var(--ink);background:var(--bg);display:flex;flex-direction:column}
button,input,select,textarea{font-family:inherit;font-size:13px}
button{cursor:pointer;border:1px solid var(--bd);background:#fff;border-radius:4px;padding:5px 11px}
button:hover{background:#f0f2f5}
button.pri{background:var(--pri);color:#fff;border-color:var(--pri)}
button.pri:hover{background:#15408a}
button.danger{color:var(--warn);border-color:#e3c2c2}
header{background:#fff;border-bottom:1px solid var(--bd);padding:9px 16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
header h1{font-size:15px;margin:0;font-weight:700;letter-spacing:-.3px}
header .sp{flex:1}
header input{border:1px solid var(--bd);border-radius:4px;padding:4px 7px}
#saveState{font-size:11px;color:var(--mut);min-width:110px}
main{flex:1;display:flex;min-height:0}
aside{width:290px;background:#fff;border-right:1px solid var(--bd);display:flex;flex-direction:column;min-height:0}
aside .find{padding:9px}
aside .find input{width:100%;border:1px solid var(--bd);border-radius:4px;padding:6px 8px}
#list{overflow:auto;flex:1;padding-bottom:10px}
#list .grp{font-size:11px;color:var(--mut);padding:9px 12px 4px;font-weight:700;background:#fafbfc;border-top:1px solid #eef0f3;position:sticky;top:0}
#list .it{padding:7px 12px;border-left:3px solid transparent;display:flex;gap:6px;align-items:baseline}
#list .it:hover{background:#f7f9fc}
#list .it.on{background:var(--pri-l);border-left-color:var(--pri)}
#list .it .nm{flex:1;line-height:1.35}
#list .it .ct{font-size:11px;color:var(--mut);background:#eef0f3;border-radius:9px;padding:1px 7px}
#list .it.on .ct{background:#fff}
section{flex:1;display:flex;flex-direction:column;min-width:0}
#head{background:#fff;border-bottom:1px solid var(--bd);padding:11px 16px}
#head h2{margin:0 0 5px;font-size:16px}
#head .meta{font-size:11.5px;color:var(--mut);line-height:1.6}
#head .meta b{color:#374151;font-weight:600}
.tag{display:inline-block;font-size:10.5px;border-radius:3px;padding:1px 6px;margin-left:5px;vertical-align:1px}
.tag.f{background:#e6f2ea;color:var(--ok)} .tag.n{background:#fdeeee;color:var(--warn)} .tag.g{background:#eef0f3;color:var(--mut)}
#bar{padding:8px 16px;display:flex;gap:7px;align-items:center;background:#fff;border-bottom:1px solid var(--bd);flex-wrap:wrap}
#bar input[type=search]{border:1px solid var(--bd);border-radius:4px;padding:5px 8px;width:210px}
#wrap{flex:1;overflow:auto;padding:12px 16px 40px}
table{border-collapse:collapse;background:#fff;width:100%}
th,td{border:1px solid var(--bd);padding:6px 8px;text-align:left;vertical-align:top;line-height:1.45}
th{background:#f0f2f5;font-weight:700;white-space:nowrap;position:sticky;top:0;z-index:1;cursor:pointer;font-size:12px}
th:hover{background:#e6e9ee}
td.n,th.n{text-align:right}
tr:hover td{background:#fafbfd}
td.act{white-space:nowrap;text-align:center}
td.act button{padding:2px 7px;font-size:11.5px}
.empty{padding:38px;text-align:center;color:var(--mut);background:#fff;border:1px dashed var(--bd);border-radius:6px}
.mask{position:fixed;inset:0;background:rgba(20,24,31,.42);display:flex;align-items:center;justify-content:center;padding:20px;z-index:50}
.modal{background:#fff;border-radius:7px;max-width:760px;width:100%;max-height:88vh;display:flex;flex-direction:column}
.modal h3{margin:0;padding:13px 17px;border-bottom:1px solid var(--bd);font-size:14.5px}
.modal .body{padding:15px 17px;overflow:auto}
.modal .foot{padding:11px 17px;border-top:1px solid var(--bd);display:flex;gap:8px;justify-content:flex-end}
.fld{margin-bottom:11px}
.fld label{display:block;font-size:11.5px;color:#4b5563;margin-bottom:3px;font-weight:600}
.fld input,.fld select,.fld textarea{width:100%;border:1px solid var(--bd);border-radius:4px;padding:6px 8px}
.fld textarea{min-height:56px;resize:vertical}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:0 14px}
.note{background:#fffbe9;border:1px solid #f0e2b0;border-radius:5px;padding:9px 11px;font-size:11.5px;line-height:1.6;color:#6b5a1e;margin-bottom:12px}
#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#1c1f24;color:#fff;padding:9px 16px;border-radius:5px;font-size:12.5px;opacity:0;transition:opacity .2s;pointer-events:none;z-index:99}
#toast.on{opacity:.95}
@media print{
 header,aside,#bar,td.act,th.act,#toast{display:none!important}
 body{background:#fff} main{display:block} #wrap{overflow:visible;padding:0}
 #head{border:0;padding:0 0 8px} th{position:static}
 table{font-size:11px} th,td{padding:4px 6px}
 @page{size:A4 landscape;margin:12mm}
}
</style></head><body>

<header>
 <h1>간이 전산 대장</h1>
 <label style="font-size:11.5px;color:var(--mut)">부점 <input id="mDept" size="12" placeholder="경영전략부"></label>
 <label style="font-size:11.5px;color:var(--mut)">연도 <input id="mYear" size="5" placeholder="2026"></label>
 <label style="font-size:11.5px;color:var(--mut)">작성자 <input id="mUser" size="8" placeholder="성명"></label>
 <span class="sp"></span>
 <span id="saveState"></span>
 <button onclick="exportJSON()">전체 백업(JSON)</button>
 <button onclick="document.getElementById('imp').click()">복원</button>
 <input type="file" id="imp" accept=".json" style="display:none" onchange="importJSON(this)">
 <button onclick="window.print()">인쇄</button>
</header>

<main>
 <aside>
  <div class="find"><input id="q" placeholder="대장·사규 검색" oninput="renderList()"></div>
  <div id="list"></div>
 </aside>
 <section>
  <div id="head"></div>
  <div id="bar"></div>
  <div id="wrap"></div>
 </section>
</main>
<div id="toast"></div>

<script>
"use strict";
const DEFS = __DEFS__;
const KEY = "kbci-daejang-v1";
let DB = {meta:{dept:"",year:"",user:""}, data:{}, custom:[]};
let cur = null, sortBy = null, sortDir = 1, filter = "";

/* ---------- 저장 ---------- */
function load(){
  try{ const s = localStorage.getItem(KEY); if(s) DB = Object.assign(DB, JSON.parse(s)); }
  catch(e){ toast("저장된 자료를 읽지 못했습니다"); }
}
function save(){
  try{
    DB.meta.savedAt = new Date().toISOString();
    localStorage.setItem(KEY, JSON.stringify(DB));
    document.getElementById("saveState").textContent = "저장됨 " + fmtTime(DB.meta.savedAt);
  }catch(e){
    document.getElementById("saveState").textContent = "저장 실패";
    alert("브라우저 저장 공간이 가득 찼습니다. [전체 백업(JSON)]으로 내려받은 뒤 오래된 기록을 정리하세요.");
  }
}
function fmtTime(iso){ const d=new Date(iso); return d.toLocaleString("ko-KR",{month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"}); }
function toast(m){ const t=document.getElementById("toast"); t.textContent=m; t.classList.add("on"); setTimeout(()=>t.classList.remove("on"),1900); }

/* ---------- 대장 ---------- */
function allDefs(){ return DEFS.concat(DB.custom||[]); }
function defOf(id){ return allDefs().find(d=>d.id===id); }
function recs(id){ return (DB.data[id] = DB.data[id] || []); }

function renderList(){
  const q = (document.getElementById("q").value||"").trim().toLowerCase();
  const box = document.getElementById("list"); box.textContent = "";
  const groups = {};
  allDefs().forEach(d=>{
    if(q && !(d.name+d.reg+d.code+(d.art||"")).toLowerCase().includes(q)) return;
    (groups[d.reg] = groups[d.reg] || []).push(d);
  });
  const keys = Object.keys(groups);
  if(!keys.length){ const e=document.createElement("div"); e.className="grp"; e.textContent="검색 결과 없음"; box.appendChild(e); return; }
  keys.forEach(g=>{
    const h=document.createElement("div"); h.className="grp"; h.textContent=g; box.appendChild(h);
    groups[g].forEach(d=>{
      const it=document.createElement("div"); it.className="it"+(cur===d.id?" on":"");
      const nm=document.createElement("div"); nm.className="nm"; nm.textContent=d.name;
      const ct=document.createElement("span"); ct.className="ct"; ct.textContent=recs(d.id).length;
      it.appendChild(nm); it.appendChild(ct);
      it.onclick=()=>{ cur=d.id; sortBy=null; filter=""; renderAll(); };
      box.appendChild(it);
    });
  });
  const add=document.createElement("div"); add.style.padding="10px 12px";
  const b=document.createElement("button"); b.textContent="+ 대장 추가"; b.onclick=newLedger;
  add.appendChild(b); box.appendChild(add);
  const warn=document.createElement("div");
  warn.style.cssText="margin:4px 10px 14px;padding:8px 10px;background:#fffbe9;border:1px solid #f0e2b0;border-radius:5px;font-size:11px;line-height:1.6;color:#6b5a1e";
  warn.textContent="기록은 이 브라우저(이 PC)에만 저장됩니다. 다른 사람과 공유하거나 백업하려면 상단 [전체 백업(JSON)]으로 파일을 내려받아 공용 폴더에 두고, 받는 쪽에서 [복원]으로 불러오세요. 브라우저 방문기록·쿠키를 지우면 기록도 함께 지워집니다.";
  box.appendChild(warn);
}

function renderHead(){
  const h=document.getElementById("head"), b=document.getElementById("bar");
  h.textContent=""; b.textContent="";
  if(!cur){ h.innerHTML='<h2>대장을 선택하세요</h2><div class="meta">왼쪽 목록에서 대장을 고르면 기록을 입력할 수 있습니다.</div>'; return; }
  const d=defOf(cur);
  const t=document.createElement("h2"); t.textContent=d.name;
  const tag=document.createElement("span");
  if(d.state==="근거조문없음"){ tag.className="tag n"; tag.textContent="근거 조문 없음"; }
  else if(d.state==="양식등록"){ tag.className="tag f"; tag.textContent="서식 등록"; }
  else { tag.className="tag g"; tag.textContent="서식 미등록"; }
  t.appendChild(tag); h.appendChild(t);
  const m=document.createElement("div"); m.className="meta";
  m.innerHTML = "<b>근거</b> 「"+esc(d.reg)+"」 "+esc(d.art||"—")+"  ·  <b>분류</b> "+esc(d.code)+
    (d.form&&d.form!=="서식없음"?"  ·  <b>서식</b> "+esc(d.form):"")+
    "<br><b>부점</b> "+esc(DB.meta.dept||"—")+"  ·  <b>연도</b> "+esc(DB.meta.year||"—")+
    "  ·  <b>기록</b> "+recs(cur).length+"건"+
    (d.src==="기본"?'  ·  <span style="color:#b03030">열 구성은 임시안입니다 — 부점 실무에 맞게 [열 편집]에서 조정하세요</span>':"");
  h.appendChild(m);

  mk(b,"pri","+ 기록 추가",()=>editRec(null));
  const s=document.createElement("input"); s.type="search"; s.placeholder="이 대장에서 검색"; s.value=filter;
  s.oninput=e=>{ filter=e.target.value; renderTable(); }; b.appendChild(s);
  const sp=document.createElement("span"); sp.style.flex="1"; b.appendChild(sp);
  mk(b,"","CSV 내보내기",exportCSV);
  mk(b,"","열 편집",editCols);
  if(d.custom) mk(b,"danger","대장 삭제",delLedger);
}
function mk(p,cls,txt,fn){ const x=document.createElement("button"); if(cls)x.className=cls; x.textContent=txt; x.onclick=fn; p.appendChild(x); return x; }
function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

function view(){
  const d=defOf(cur); let r=recs(cur).slice();
  if(filter){ const q=filter.toLowerCase(); r=r.filter(x=>d.cols.some(c=>String(x[c.name]||"").toLowerCase().includes(q))); }
  if(sortBy){ const c=d.cols.find(c=>c.name===sortBy);
    r.sort((a,b)=>{ let x=a[sortBy]||"",y=b[sortBy]||"";
      if(c&&c.type==="num"){ x=parseFloat(x)||0; y=parseFloat(y)||0; return (x-y)*sortDir; }
      return String(x).localeCompare(String(y),"ko")*sortDir; });
  }
  return r;
}

function renderTable(){
  const w=document.getElementById("wrap"); w.textContent="";
  if(!cur) return;
  const d=defOf(cur), rows=view();
  if(!rows.length){
    const e=document.createElement("div"); e.className="empty";
    e.textContent = recs(cur).length ? "검색 결과가 없습니다." : "아직 기록이 없습니다. [+ 기록 추가]로 첫 건을 입력하세요.";
    w.appendChild(e); return;
  }
  const tb=document.createElement("table"), th=document.createElement("thead"), hr=document.createElement("tr");
  const no=document.createElement("th"); no.textContent="No"; no.className="n"; hr.appendChild(no);
  d.cols.forEach(c=>{
    const t=document.createElement("th"); if(c.type==="num") t.className="n";
    t.textContent=c.name+(sortBy===c.name?(sortDir>0?" ▲":" ▼"):"");
    t.onclick=()=>{ if(sortBy===c.name) sortDir=-sortDir; else {sortBy=c.name;sortDir=1;} renderTable(); };
    hr.appendChild(t);
  });
  const ta=document.createElement("th"); ta.className="act"; ta.textContent="관리"; ta.style.cursor="default"; hr.appendChild(ta);
  th.appendChild(hr); tb.appendChild(th);
  const bd=document.createElement("tbody");
  rows.forEach((r,i)=>{
    const tr=document.createElement("tr");
    const n=document.createElement("td"); n.className="n"; n.textContent=i+1; tr.appendChild(n);
    d.cols.forEach(c=>{
      const td=document.createElement("td"); if(c.type==="num") td.className="n";
      td.textContent = c.type==="check" ? (r[c.name]?"O":"") : (r[c.name]||"");
      tr.appendChild(td);
    });
    const a=document.createElement("td"); a.className="act";
    mk(a,"","수정",()=>editRec(r._id)); mk(a,"danger","삭제",()=>delRec(r._id));
    tr.appendChild(a); bd.appendChild(tr);
  });
  tb.appendChild(bd); w.appendChild(tb);
  const f=document.createElement("div"); f.style.cssText="margin-top:9px;font-size:11.5px;color:var(--mut)";
  f.textContent="표시 "+rows.length+"건 / 전체 "+recs(cur).length+"건";
  w.appendChild(f);
}
function renderAll(){ renderList(); renderHead(); renderTable(); }

/* ---------- 기록 입력 ---------- */
function editRec(id){
  const d=defOf(cur), rec = id ? recs(cur).find(r=>r._id===id) : {};
  const body=document.createElement("div"); body.className="grid2";
  d.cols.forEach(c=>{
    const f=document.createElement("div"); f.className="fld";
    if(c.type==="memo") f.style.gridColumn="1 / -1";
    const l=document.createElement("label"); l.textContent=c.name; f.appendChild(l);
    let el;
    if(c.type==="select"){ el=document.createElement("select");
      el.appendChild(new Option("",""));
      (c.opts||[]).forEach(o=>el.appendChild(new Option(o,o)));
    } else if(c.type==="memo"){ el=document.createElement("textarea");
    } else { el=document.createElement("input");
      el.type = c.type==="date"?"date" : c.type==="time"?"time" : c.type==="num"?"number" : c.type==="check"?"checkbox" : "text";
    }
    if(c.type==="check") el.checked = !!rec[c.name]; else el.value = rec[c.name]||"";
    el.dataset.k=c.name; el.dataset.t=c.type;
    if(c.type==="check"){ const w=document.createElement("div"); w.style.paddingTop="4px"; el.style.width="auto"; w.appendChild(el); f.appendChild(w); }
    else f.appendChild(el);
    body.appendChild(f);
  });
  openModal((id?"기록 수정":"기록 추가")+" — "+d.name, body, ()=>{
    const o = id ? rec : {_id:"r"+Date.now()+Math.random().toString(36).slice(2,6), _by:DB.meta.user||"", _at:new Date().toISOString()};
    body.querySelectorAll("[data-k]").forEach(el=>{ o[el.dataset.k] = el.dataset.t==="check" ? el.checked : el.value; });
    o._upd = new Date().toISOString(); if(!id) recs(cur).push(o);
    save(); renderAll(); toast(id?"수정했습니다":"추가했습니다"); return true;
  });
}
function delRec(id){
  if(!confirm("이 기록을 삭제합니다. 되돌릴 수 없습니다.")) return;
  DB.data[cur]=recs(cur).filter(r=>r._id!==id); save(); renderAll(); toast("삭제했습니다");
}

/* ---------- 대장 정의 ---------- */
function colEditor(cols){
  const box=document.createElement("div");
  const note=document.createElement("div"); note.className="note";
  note.textContent="열 이름을 바꾸거나 순서를 조정할 수 있습니다. 이미 입력된 기록은 예전 열 이름으로 저장되어 있으므로, 이름을 바꾸면 그 열의 기존 값은 표에 보이지 않습니다.";
  box.appendChild(note);
  const list=document.createElement("div"); box.appendChild(list);
  function draw(){
    list.textContent="";
    cols.forEach((c,i)=>{
      const row=document.createElement("div"); row.style.cssText="display:flex;gap:6px;margin-bottom:6px;align-items:center";
      const nm=document.createElement("input"); nm.value=c.name; nm.style.flex="1"; nm.oninput=e=>c.name=e.target.value;
      const ty=document.createElement("select"); ty.style.width="110px";
      [["text","텍스트"],["date","날짜"],["time","시간"],["num","숫자"],["select","선택"],["memo","여러 줄"],["check","체크"]]
        .forEach(([v,t])=>{ const o=new Option(t,v); if(c.type===v)o.selected=true; ty.appendChild(o); });
      ty.onchange=e=>{ c.type=e.target.value; draw(); };
      row.appendChild(nm); row.appendChild(ty);
      if(c.type==="select"){ const op=document.createElement("input"); op.style.width="170px"; op.placeholder="선택지, 쉼표로";
        op.value=(c.opts||[]).join(","); op.oninput=e=>c.opts=e.target.value.split(",").map(s=>s.trim()).filter(Boolean); row.appendChild(op); }
      const up=document.createElement("button"); up.textContent="↑"; up.onclick=()=>{ if(i>0){ cols.splice(i-1,0,cols.splice(i,1)[0]); draw(); } };
      const dn=document.createElement("button"); dn.textContent="↓"; dn.onclick=()=>{ if(i<cols.length-1){ cols.splice(i+1,0,cols.splice(i,1)[0]); draw(); } };
      const rm=document.createElement("button"); rm.className="danger"; rm.textContent="삭제"; rm.onclick=()=>{ cols.splice(i,1); draw(); };
      row.appendChild(up); row.appendChild(dn); row.appendChild(rm); list.appendChild(row);
    });
    const add=document.createElement("button"); add.textContent="+ 열 추가";
    add.onclick=()=>{ cols.push({name:"새 열",type:"text"}); draw(); }; list.appendChild(add);
  }
  draw(); return box;
}
function editCols(){
  const d=defOf(cur), cols=JSON.parse(JSON.stringify(d.cols));
  openModal("열 편집 — "+d.name, colEditor(cols), ()=>{
    if(!cols.length){ alert("열이 하나는 있어야 합니다."); return false; }
    DB.custom = DB.custom||[];
    let c = DB.custom.find(x=>x.id===d.id);
    if(c) c.cols=cols;
    else { DB.custom.push(Object.assign({}, d, {cols:cols, custom:d.custom||false, over:true})); }
    save(); renderAll(); toast("열 구성을 저장했습니다"); return true;
  });
}
function newLedger(){
  const body=document.createElement("div");
  const f1=document.createElement("div"); f1.className="fld";
  f1.innerHTML='<label>대장명</label>'; const n1=document.createElement("input"); f1.appendChild(n1);
  const f2=document.createElement("div"); f2.className="fld";
  f2.innerHTML='<label>근거 사규</label>'; const n2=document.createElement("input"); f2.appendChild(n2);
  const f3=document.createElement("div"); f3.className="fld";
  f3.innerHTML='<label>근거 조문</label>'; const n3=document.createElement("input"); f3.appendChild(n3);
  body.appendChild(f1); body.appendChild(f2); body.appendChild(f3);
  const cols=[{name:"일자",type:"date"},{name:"내용",type:"memo"},{name:"담당자",type:"text"}];
  body.appendChild(colEditor(cols));
  openModal("대장 추가", body, ()=>{
    if(!n1.value.trim()){ alert("대장명을 입력하세요."); return false; }
    DB.custom = DB.custom||[];
    const id="c"+Date.now();
    DB.custom.push({id:id, code:"직접", reg:n2.value.trim()||"직접 등록", name:n1.value.trim(), art:n3.value.trim(),
                    form:"", state:n3.value.trim()?"양식없음":"근거조문없음", cols:cols, custom:true, src:"직접"});
    cur=id; save(); renderAll(); toast("대장을 추가했습니다"); return true;
  });
}
function delLedger(){
  const d=defOf(cur);
  if(!d.custom) return;
  if(!confirm("['"+d.name+"'] 대장과 기록 "+recs(cur).length+"건을 모두 삭제합니다. 되돌릴 수 없습니다.")) return;
  DB.custom=DB.custom.filter(x=>x.id!==cur); delete DB.data[cur]; cur=null; save(); renderAll();
}

/* ---------- 모달 ---------- */
function openModal(title, bodyEl, onOk){
  const mask=document.createElement("div"); mask.className="mask";
  const m=document.createElement("div"); m.className="modal";
  const h=document.createElement("h3"); h.textContent=title;
  const b=document.createElement("div"); b.className="body"; b.appendChild(bodyEl);
  const f=document.createElement("div"); f.className="foot";
  const c=document.createElement("button"); c.textContent="취소"; c.onclick=()=>mask.remove();
  const o=document.createElement("button"); o.className="pri"; o.textContent="저장";
  o.onclick=()=>{ if(onOk()!==false) mask.remove(); };
  f.appendChild(c); f.appendChild(o);
  m.appendChild(h); m.appendChild(b); m.appendChild(f); mask.appendChild(m); document.body.appendChild(mask);
  mask.onclick=e=>{ if(e.target===mask) mask.remove(); };
  const first=b.querySelector("input,select,textarea"); if(first) first.focus();
}

/* ---------- 내보내기 ---------- */
function dl(name, text, mime){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([text],{type:mime||"text/plain;charset=utf-8"}));
  a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500);
}
function csvCell(v){ v=String(v==null?"":v); return /[",\n]/.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; }
function exportCSV(){
  const d=defOf(cur), rows=view();
  const head=["No"].concat(d.cols.map(c=>c.name)).concat(["작성자","작성일시"]);
  const lines=[head.map(csvCell).join(",")];
  rows.forEach((r,i)=>{
    lines.push([i+1].concat(d.cols.map(c=>c.type==="check"?(r[c.name]?"O":""):(r[c.name]||"")))
      .concat([r._by||"", r._at?fmtTime(r._at):""]).map(csvCell).join(","));
  });
  const title=[d.name, DB.meta.dept, DB.meta.year].filter(Boolean).join("_");
  dl(title+".csv", "﻿"+lines.join("\r\n"), "text/csv;charset=utf-8");
  toast("CSV로 내려받았습니다");
}
function exportJSON(){
  const stamp=new Date().toISOString().slice(0,10).replace(/-/g,"");
  dl("간이대장_백업_"+(DB.meta.dept||"부점")+"_"+stamp+".json", JSON.stringify(DB,null,1), "application/json");
  toast("전체 백업을 내려받았습니다");
}
function importJSON(inp){
  const f=inp.files[0]; if(!f) return;
  const rd=new FileReader();
  rd.onload=()=>{
    try{
      const o=JSON.parse(rd.result);
      if(!o.data) throw new Error("형식이 다릅니다");
      if(!confirm("현재 브라우저에 저장된 내용을 덮어씁니다. 계속할까요?")) return;
      DB=Object.assign({meta:{},data:{},custom:[]}, o);
      cur=null; fillMeta(); save(); renderAll(); toast("복원했습니다");
    }catch(e){ alert("복원 실패: "+e.message); }
    inp.value="";
  };
  rd.readAsText(f);
}

/* ---------- 초기화 ---------- */
function fillMeta(){
  document.getElementById("mDept").value=DB.meta.dept||"";
  document.getElementById("mYear").value=DB.meta.year||"";
  document.getElementById("mUser").value=DB.meta.user||"";
}
["mDept","mYear","mUser"].forEach(id=>{
  document.addEventListener("DOMContentLoaded",()=>{
    document.getElementById(id).addEventListener("change",e=>{
      DB.meta[id.slice(1).toLowerCase()]=e.target.value; save(); renderHead();
    });
  });
});
window.addEventListener("DOMContentLoaded",()=>{
  load(); fillMeta();
  if(DB.meta.savedAt) document.getElementById("saveState").textContent="저장됨 "+fmtTime(DB.meta.savedAt);
  else document.getElementById("saveState").textContent="새 문서";
  if(!DB.meta.year){ DB.meta.year=String(new Date().getFullYear()); document.getElementById("mYear").value=DB.meta.year; }
  renderAll();
});
</script></body></html>'''


def main():
    defs = build_defs()
    html = HTML.replace('__DEFS__', json.dumps(defs, ensure_ascii=False, separators=(',', ':')))
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, 'drafts/간이전산대장.html')
    open(out, 'w', encoding='utf-8').write(html)
    n_form = sum(1 for d in defs if d['src'] == '서식')
    print('생성: %s  (%.0f KB)' % (out, os.path.getsize(out) / 1024))
    print('  대장 %d건 / 등록 서식 열 구조 반영 %d건 / 기본안 %d건'
          % (len(defs), n_form, sum(1 for d in defs if d['src'] == '기본')))
    return 0


if __name__ == '__main__':
    sys.exit(main())
