"""검사성적서 관리 시스템 - 웹 UI (FastAPI 서버 렌더링).

실행:  python run.py      또는      uvicorn app:app --host 0.0.0.0 --port 8000

운영 흐름
  1) 모델 등록
  2) 출하검사 기준서(Word) 업로드 -> 표 자동 분석 -> 열 매핑 확인 -> 검사항목 적용
  3) 성적서 생성(검사항목 스냅샷) -> 측정값 입력 -> 자동 판정
  4) Excel 성적서 출력
  5) 모든 단계가 이력으로 남는다 (기준서 개정 이력 + 변경 이력)

ponytail: 템플릿 엔진·프런트엔드 빌드 없이 표준 라이브러리 문자열과 html.escape 로
HTML 을 만든다. 이미 설치된 fastapi/uvicorn/openpyxl 외에 추가 의존성이 없다.
"""

from __future__ import annotations

import os
import sqlite3
from datetime import date
from html import escape as esc

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response

import db
import docx_parse
import report_excel

app = FastAPI(title="검사성적서 관리 시스템")

SPEC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "spec_docs")
MAX_UPLOAD = 20 * 1024 * 1024          # 기준서 업로드 상한 20MB


@app.on_event("startup")
def _startup() -> None:
    db.init_db()
    os.makedirs(SPEC_DIR, exist_ok=True)


def actor_of(request: Request) -> str:
    """작성자 식별. 사내 SSO 가 붙기 전까지는 쿠키에 담은 이름을 쓴다.

    쿠키 값은 latin-1 만 담을 수 있어 한국어 이름을 percent-encoding 해서 넣는다.

    ponytail: 인증 없음. 사내망 전용 도구이고 이력에 '누가' 를 남기는 것이 목적이다.
    업그레이드 경로 - 사내 SSO 헤더 또는 ERP 계정 연동 시 이 함수만 교체한다.
    """
    from urllib.parse import unquote
    raw = request.cookies.get("actor") or ""
    try:
        return unquote(raw).strip()[:30]
    except (ValueError, UnicodeDecodeError):
        return ""


def spec_path(doc_id: int) -> str:
    return os.path.join(SPEC_DIR, f"{doc_id}.docx")


# ------------------------------------------------------------------ 렌더 헬퍼

CSS = """
*{box-sizing:border-box}
body{margin:0;font:14px/1.55 "Malgun Gothic","맑은 고딕",system-ui,sans-serif;color:#1b1f24;background:#f6f7f9}
a{color:#1a56b3;text-decoration:none}
a:hover{text-decoration:underline}
header{background:#1f3a5f;color:#fff;padding:0 20px;display:flex;align-items:center;gap:22px;height:52px}
header .brand{font-weight:700;font-size:16px;margin-right:10px}
header a{color:#cfe0f5;padding:4px 2px;border-bottom:2px solid transparent}
header a.on{color:#fff;border-bottom-color:#8fc0ff}
main{max-width:1500px;margin:0 auto;padding:20px}
h1{font-size:19px;margin:0 0 4px}
h2{font-size:15px;margin:22px 0 8px;padding-bottom:5px;border-bottom:1px solid #d9dee5}
.sub{color:#6b7280;font-size:12.5px;margin:0 0 16px}
.card{background:#fff;border:1px solid #dfe3e8;border-radius:6px;padding:16px;margin-bottom:16px}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px}
th,td{border:1px solid #dfe3e8;padding:6px 8px;text-align:left;vertical-align:middle}
th{background:#eef2f7;font-weight:600;white-space:nowrap}
td.c,th.c{text-align:center}
tr:hover td{background:#fafbfd}
input,select,textarea{font:13px inherit;padding:5px 7px;border:1px solid #c3c9d2;border-radius:4px;background:#fff}
input[readonly]{background:#f1f3f5;color:#6b7280}
textarea{width:100%;min-height:52px}
button{font:13px inherit;padding:6px 14px;border:1px solid #1f3a5f;border-radius:4px;background:#1f3a5f;color:#fff;cursor:pointer}
button.alt{background:#fff;color:#1f3a5f}
button.danger{background:#fff;border-color:#b42318;color:#b42318}
button:hover{opacity:.88}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end}
.f{display:flex;flex-direction:column;gap:3px}
.f label{font-size:12px;color:#4b5563;font-weight:600}
.stat{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}
.stat div{background:#fff;border:1px solid #dfe3e8;border-radius:6px;padding:10px 16px;min-width:108px}
.stat b{display:block;font-size:21px;line-height:1.25}
.stat span{font-size:12px;color:#6b7280}
.msg{padding:9px 12px;border-radius:5px;margin-bottom:14px;font-size:13px}
.msg.ok{background:#e7f6ec;border:1px solid #a7d9b9;color:#13613a}
.msg.err{background:#fdeceb;border:1px solid #f1b5b0;color:#9c2019}
.tag{display:inline-block;padding:1px 7px;border-radius:9px;font-size:11.5px;font-weight:600;border:1px solid}
.t-ok{background:#e7f6ec;border-color:#a7d9b9;color:#13613a}
.t-ng{background:#fdeceb;border-color:#f1b5b0;color:#9c2019}
.t-na{background:#f1f3f5;border-color:#d4d9df;color:#6b7280}
.val{width:62px;text-align:center}
.val.ng{background:#fdeceb;border-color:#e4897f}
.cat{background:#e8eef6;font-weight:700}
.muted{color:#6b7280;font-size:12px}
.right{text-align:right}
.inline{display:inline}
.tight td,.tight th{padding:4px 6px}
.scroll{overflow-x:auto}
.who{margin-left:auto;display:flex;gap:5px;align-items:center}
.who input{padding:3px 6px;font-size:12px}
.who button{padding:3px 9px;font-size:12px}
.steps{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 16px;padding:0;list-style:none;font-size:12.5px}
.steps li{background:#fff;border:1px solid #dfe3e8;border-radius:20px;padding:4px 13px;color:#6b7280}
.steps li.on{background:#1f3a5f;border-color:#1f3a5f;color:#fff;font-weight:600}
.steps li.done{background:#e7f6ec;border-color:#a7d9b9;color:#13613a}
.prev{font-size:11.5px;border-collapse:collapse}
.prev td,.prev th{border:1px solid #dfe3e8;padding:3px 5px;max-width:190px;overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap}
.prev tr.hdr td{background:#eef2f7;font-weight:600}
.badge{display:inline-block;background:#1f3a5f;color:#fff;border-radius:3px;padding:1px 6px;font-size:11px}
.badge.cur{background:#13613a}
code{background:#f1f3f5;padding:1px 4px;border-radius:3px;font-size:12px}
details summary{cursor:pointer;font-weight:600;font-size:13px;padding:4px 0}
"""

NAV = [("/", "대시보드"), ("/models", "모델 · 검사항목"), ("/specs", "기준서 이력"),
       ("/reports", "성적서"), ("/history", "변경 이력")]


def page(title: str, body: str, active: str = "", request: Request | None = None) -> HTMLResponse:
    flash, who = "", ""
    if request is not None:
        if msg := request.query_params.get("msg"):
            flash += f'<div class="msg ok">{esc(msg)}</div>'
        if err := request.query_params.get("err"):
            flash += f'<div class="msg err">{esc(err)}</div>'
        name = actor_of(request)
        who = (f'<form method="post" action="/actor" class="who">'
               f'<input name="actor" value="{esc(name)}" placeholder="작성자 이름" style="width:110px">'
               f'<button type="submit" class="alt">변경</button></form>'
               if name else
               '<form method="post" action="/actor" class="who">'
               '<input name="actor" placeholder="작성자 이름 입력" style="width:130px">'
               '<button type="submit" class="alt">저장</button></form>')
    nav = "".join(
        f'<a href="{href}" class="{"on" if href == active else ""}">{esc(label)}</a>'
        for href, label in NAV
    )
    return HTMLResponse(
        "<!doctype html><html lang=\"ko\"><head><meta charset=\"utf-8\">"
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f"<title>{esc(title)} · 검사성적서 관리</title><style>{CSS}</style></head><body>"
        f'<header><span class="brand">검사성적서 관리 시스템</span>{nav}{who}</header>'
        f"<main>{flash}{body}</main></body></html>"
    )


@app.post("/actor")
async def set_actor(request: Request):
    from urllib.parse import quote
    form = await request.form()
    name = str(form.get("actor", "") or "").strip()[:30]
    ref = request.headers.get("referer") or "/"
    resp = RedirectResponse(ref, status_code=303)
    # 쿠키 값은 latin-1 로만 인코딩되므로 한국어 이름은 percent-encoding 한다
    resp.set_cookie("actor", quote(name), max_age=60 * 60 * 24 * 365,
                    httponly=True, samesite="lax")
    return resp


def see(url: str, msg: str = "", err: str = "") -> RedirectResponse:
    from urllib.parse import quote
    if msg:
        url += ("&" if "?" in url else "?") + "msg=" + quote(msg)
    if err:
        url += ("&" if "?" in url else "?") + "err=" + quote(err)
    return RedirectResponse(url, status_code=303)


def field(label: str, name: str, value="", kind: str = "text", width: int = 150, **attrs) -> str:
    extra = " ".join(f'{k.replace("_", "-")}="{esc(str(v))}"' for k, v in attrs.items())
    v = "" if value is None else esc(str(value))
    return (f'<div class="f"><label>{esc(label)}</label>'
            f'<input type="{kind}" name="{name}" value="{v}" style="width:{width}px" {extra}></div>')


def select(label: str, name: str, options: list[tuple[str, str]], selected: str = "",
           width: int = 170, **attrs) -> str:
    extra = " ".join(f'{k.replace("_", "-")}="{esc(str(v))}"' for k, v in attrs.items())
    opts = "".join(
        f'<option value="{esc(str(val))}" {"selected" if str(val) == str(selected) else ""}>{esc(text)}</option>'
        for val, text in options
    )
    return (f'<div class="f"><label>{esc(label)}</label>'
            f'<select name="{name}" style="width:{width}px" {extra}>{opts}</select></div>')


def judge_tag(judge: str) -> str:
    cls = {"OK": "t-ok", "NG": "t-ng"}.get(judge, "t-na")
    return f'<span class="tag {cls}">{esc(judge or "미판정")}</span>'


def limit_text(row) -> str:
    lo, hi = row["lsl"], row["usl"]
    if lo is None and hi is None:
        return '<span class="muted">수동판정</span>'
    fmt = lambda x: ("" if x is None else f"{x:g}")
    return f'<span class="muted">{fmt(lo) or "-"} ~ {fmt(hi) or "-"}</span>'


# -------------------------------------------------------------------- 대시보드

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    s = db.stats()
    recent = db.list_reports()[:12]
    models = db.list_models()
    no_item = [m for m in models if not m["item_count"]]

    cards = "".join(
        f"<div><b>{v}</b><span>{esc(k)}</span></div>" for k, v in [
            ("등록 모델", s["models"]), ("기준서 개정", s["docs"]), ("검사항목(활성)", s["items"]),
            ("성적서", s["reports"]), ("작성중", s["by_status"]["작성중"]),
            ("검토요청", s["by_status"]["검토요청"]), ("승인", s["by_status"]["승인"]),
            ("합격", s["pass"]), ("불합격", s["fail"]), ("이력 건수", s["events"]),
        ]
    )

    rows = "".join(
        f"<tr><td><a href=\"/reports/{r['id']}\">{esc(r['report_no'])}</a></td>"
        f"<td>{esc(r['model_name'])}</td><td>{esc(r['part_no'])}</td>"
        f"<td class=\"c\">{esc(r['inspect_date'] or '-')}</td>"
        f"<td class=\"c\">{esc(r['status'])}</td>"
        f"<td class=\"c\">{judge_tag(r['overall'])}</td>"
        f"<td class=\"c\">{r['line_count'] - r['pending_count']} / {r['line_count']}</td></tr>"
        for r in recent
    ) or '<tr><td colspan="7" class="c muted">성적서가 없습니다.</td></tr>'

    warn = ""
    if no_item:
        links = ", ".join(f'<a href="/models/{m["id"]}">{esc(m["model_name"])}</a>' for m in no_item[:10])
        warn = (f'<div class="msg err">검사항목이 등록되지 않은 모델 {len(no_item)}건: {links}'
                " — 성적서를 만들기 전에 기준서 항목을 먼저 등록해야 합니다.</div>")

    body = f"""
    <h1>대시보드</h1>
    <p class="sub">기준일 {date.today().isoformat()} · 출하검사 성적서 작성 자동화 (품질팀)
       · 기준서(Word) → 검사항목 추출 → 성적서(Excel) + 이력관리</p>
    {steps(-1)}
    {warn}
    <div class="stat">{cards}</div>

    <h2>최근 성적서</h2>
    <div class="scroll"><table>
      <tr><th>성적서번호</th><th>모델명</th><th>품번</th><th class="c">검사일</th>
          <th class="c">상태</th><th class="c">종합판정</th><th class="c">입력 진행</th></tr>
      {rows}
    </table></div>

    <h2>최근 변경 이력</h2>
    <div class="scroll"><table class="tight">
      <tr><th class="c">일시</th><th class="c">담당</th><th class="c">대상</th>
          <th class="c">행위</th><th>내용</th><th class="c">바로가기</th></tr>
      {audit_table(db.list_audit(limit=15))}
    </table></div>
    <p><a href="/history">전체 변경 이력 →</a></p>
    """
    return page("대시보드", body, "/", request)


# ------------------------------------------------------------------ 모델 마스터

@app.get("/models", response_class=HTMLResponse)
def models_page(request: Request, q: str = ""):
    rows = "".join(
        f"<tr><td><a href=\"/models/{m['id']}\">{esc(m['model_name'])}</a></td>"
        f"<td>{esc(m['part_no'])}</td><td class=\"c\">{esc(m['sec_code'] or '-')}</td>"
        f"<td class=\"c\">{esc(m['rev'] or '-')}</td><td>{esc(m['customer'] or '-')}</td>"
        f"<td>{esc(m['spec_doc_no'] or '-')}</td>"
        f"<td class=\"c\">{m['item_count']}</td><td class=\"c\">{m['report_count']}</td></tr>"
        for m in db.list_models(q)
    ) or '<tr><td colspan="8" class="c muted">등록된 모델이 없습니다.</td></tr>'

    body = f"""
    <h1>모델 · 검사항목 마스터</h1>
    <p class="sub">출하검사 기준서의 검사항목·규격을 모델별로 백데이터화하여 보관합니다.</p>

    <div class="card">
      <form method="get" class="row">
        {field("검색 (모델명 / 품번 / SEC CODE)", "q", q, width=280)}
        <button type="submit">검색</button>
        <a href="/models"><button type="button" class="alt">초기화</button></a>
      </form>
    </div>

    <div class="scroll"><table>
      <tr><th>모델명</th><th>품번</th><th class="c">SEC CODE</th><th class="c">REV</th>
          <th>고객사</th><th>기준서번호</th><th class="c">검사항목</th><th class="c">성적서</th></tr>
      {rows}
    </table></div>

    <h2>모델 신규 등록</h2>
    <div class="card"><form method="post" action="/models/new" class="row">
      {field("모델명 *", "model_name", width=240, required="required")}
      {field("품번 *", "part_no", width=170, required="required")}
      {field("SEC CODE", "sec_code", width=130)}
      {field("REV", "rev", width=80)}
      {field("고객사", "customer", width=130)}
      {field("기준서번호", "spec_doc_no", width=160)}
      {field("비고", "note", width=200)}
      <button type="submit">등록</button>
    </form></div>
    """
    return page("모델 마스터", body, "/models", request)


@app.post("/models/new")
async def model_new(request: Request):
    form = dict(await request.form())
    if not str(form.get("model_name", "")).strip() or not str(form.get("part_no", "")).strip():
        return see("/models", err="모델명과 품번은 필수입니다.")
    try:
        mid = db.create_model(form, actor=actor_of(request))
    except sqlite3.IntegrityError:
        return see("/models", err="같은 모델명·품번·REV 조합이 이미 등록되어 있습니다.")
    return see(f"/models/{mid}", msg="모델을 등록했습니다. 검사항목을 등록해 주십시오.")


@app.get("/models/{model_id}", response_class=HTMLResponse)
def model_detail(model_id: int, request: Request):
    m = db.get_model(model_id)
    if m is None:
        return see("/models", err="모델을 찾을 수 없습니다.")
    items = db.list_items(model_id)
    docs = db.list_spec_docs(model_id)
    cur_doc = db.current_spec_doc(model_id)
    others = [(str(o["id"]), f"{o['model_name']} / {o['part_no']} ({o['item_count']}항목)")
              for o in db.list_models() if o["id"] != model_id and o["item_count"]]

    rows, last_cat = [], None
    for it in items:
        if it["category"] != last_cat:
            last_cat = it["category"]
            n = sum(1 for x in items if x["category"] == last_cat)
            rows.append(f'<tr><td colspan="11" class="cat">{esc(last_cat)} · {n}항목</td></tr>')
        dim = "" if it["active"] else ' style="opacity:.45"'
        rows.append(
            f'<tr{dim}><form method="post" action="/items/{it["id"]}/edit" class="inline">'
            f'<td class="c">{select("", "category", [(c, c) for c in db.CATEGORIES], it["category"], width=78)}</td>'
            f'<td class="c">{field("", "seq", it["seq"], kind="number", width=58)}</td>'
            f'<td>{field("", "item_name", it["item_name"], width=260, required="required")}</td>'
            f'<td>{field("", "spec_text", it["spec_text"], width=135)}</td>'
            f'<td class="c">{limit_text(it)}</td>'
            f'<td class="c">{field("", "unit", it["unit"], width=58)}</td>'
            f'<td class="c">{select("", "judge_type", [("numeric", "수치"), ("ox", "OX")], it["judge_type"], width=78)}</td>'
            f'<td class="c">{field("", "sample_size", it["sample_size"], kind="number", width=56, min="1", max="50")}</td>'
            f'<td>{field("", "method", it["method"], width=120)}</td>'
            f'<td>{field("", "gauge", it["gauge"], width=110)}</td>'
            f'<td class="c" style="white-space:nowrap">'
            f'<label class="muted"><input type="checkbox" name="active" value="1" '
            f'{"checked" if it["active"] else ""}> 사용</label> '
            f'<button type="submit" class="alt">저장</button></form>'
            f'<form method="post" action="/items/{it["id"]}/delete" class="inline">'
            f'<button type="submit" class="danger">삭제</button></form></td></tr>'
        )
    item_rows = "".join(rows) or '<tr><td colspan="11" class="c muted">검사항목이 없습니다.</td></tr>'

    copy_block = ""
    if others:
        copy_block = f"""
        <form method="post" action="/models/{model_id}/items/copy" class="row" style="margin-top:12px">
          {select("다른 모델에서 복사", "src_model_id", others, width=330)}
          <button type="submit" class="alt">검사항목 복사</button>
        </form>"""

    doc_rows = "".join(
        f"<tr><td class=\"c muted\" style=\"white-space:nowrap\">{esc(d['uploaded_at'])}</td>"
        f"<td><a href=\"/specs/{d['id']}/file\">{esc(d['filename'])}</a>"
        + (' <span class="badge cur">현행</span>' if d["is_current"] else "")
        + f"</td><td class=\"c\">{esc(d['doc_no'] or '-')}</td>"
        f"<td class=\"c\">{esc(d['rev'] or '-')}</td>"
        f"<td class=\"c\">{d['item_count'] or 0}</td>"
        f"<td class=\"c\">{esc(d['applied_at'] or '미적용')}</td>"
        f"<td class=\"c\">{esc(d['uploaded_by'] or '-')}</td>"
        f"<td class=\"c\"><a href=\"/specs/{d['id']}/map\">열 매핑 · 재적용</a></td></tr>"
        for d in docs
    ) or '<tr><td colspan="8" class="c muted">업로드된 기준서가 없습니다.</td></tr>'

    cur_line = (f'현행 기준서 <b>{esc(cur_doc["filename"])}</b>'
                f' (문서번호 {esc(cur_doc["doc_no"] or "-")} / REV {esc(cur_doc["rev"] or "-")},'
                f' 적용 {esc(cur_doc["applied_at"] or "-")})'
                if cur_doc is not None else
                '<span class="muted">현행 기준서가 지정되지 않았습니다. 기준서를 올려 검사항목을 적용해 주십시오.</span>')

    body = f"""
    {steps(1 if not docs else (3 if items else 2))}
    <h1>{esc(m["model_name"])} <span class="muted">/ {esc(m["part_no"])}</span></h1>
    <p class="sub">SEC CODE {esc(m["sec_code"] or "-")} · REV {esc(m["rev"] or "-")}
       · 고객사 {esc(m["customer"] or "-")} · 기준서번호 {esc(m["spec_doc_no"] or "-")}<br>{cur_line}</p>

    <h2>출하검사 기준서(Word) 업로드</h2>
    <div class="card">
      <form method="post" action="/models/{model_id}/spec/upload"
            enctype="multipart/form-data" class="row">
        <div class="f"><label>기준서 파일 (.docx)</label>
          <input type="file" name="file" accept=".docx" required style="width:300px"></div>
        {field("문서번호", "doc_no", width=150, placeholder="비우면 문서에서 추출")}
        {field("REV", "rev", width=80, placeholder="자동")}
        {field("비고", "note", width=200)}
        <button type="submit">업로드 · 표 분석</button>
      </form>
      <p class="muted" style="margin:10px 0 0">
        Word 기준서의 검사항목 표를 읽어 <b>구분 · 검사항목 · 규격 · 단위 · 검사방법 · 측정기구 · 시료수</b>를
        추출합니다. 열 역할은 자동 추정 후 다음 화면에서 담당자가 확인·수정합니다.
        원본 파일은 개정 이력으로 보관됩니다. (최대 20MB, <code>.docx</code> 전용)</p>
    </div>

    <h2>기준서 개정 이력 · {len(docs)}건</h2>
    <div class="scroll"><table class="tight">
      <tr><th class="c">업로드</th><th>파일</th><th class="c">문서번호</th><th class="c">REV</th>
          <th class="c">적용항목</th><th class="c">적용일시</th><th class="c">담당</th><th class="c">작업</th></tr>
      {doc_rows}
    </table></div>

    <h2>모델정보</h2>
    <div class="card"><form method="post" action="/models/{model_id}/edit" class="row">
      {field("모델명 *", "model_name", m["model_name"], width=240, required="required")}
      {field("품번 *", "part_no", m["part_no"], width=170, required="required")}
      {field("SEC CODE", "sec_code", m["sec_code"], width=130)}
      {field("REV", "rev", m["rev"], width=80)}
      {field("고객사", "customer", m["customer"], width=130)}
      {field("기준서번호", "spec_doc_no", m["spec_doc_no"], width=160)}
      {field("비고", "note", m["note"], width=200)}
      <button type="submit">모델정보 저장</button>
    </form>
    <form method="post" action="/models/{model_id}/delete" style="margin-top:10px">
      <button type="submit" class="danger"
        onclick="return confirm('이 모델과 검사항목을 삭제합니다. 계속하시겠습니까?')">모델 삭제</button>
      <span class="muted">성적서가 연결된 모델은 삭제되지 않습니다.</span>
    </form></div>

    <h2>검사항목 마스터 · {len(items)}건</h2>
    <div class="scroll"><table class="tight">
      <tr><th class="c">구분</th><th class="c">순번</th><th>검사항목</th><th>규격</th>
          <th class="c">하한~상한</th><th class="c">단위</th><th class="c">판정</th>
          <th class="c">시료수</th><th>검사방법</th><th>측정기구</th><th class="c">작업</th></tr>
      {item_rows}
    </table></div>
    <p class="muted">규격에 <code>1.25±0.10</code>, <code>1.0~2.0</code>, <code>MAX 2.0</code>,
       <code>0.5 이하</code> 처럼 입력하면 하한·상한을 자동으로 계산해 측정값을 자동 판정합니다.
       수치 한계를 찾지 못하면 OX 판정으로 처리합니다.</p>

    <h2>검사항목 추가</h2>
    <div class="card"><form method="post" action="/models/{model_id}/items/new" class="row">
      {select("구분", "category", [(c, c) for c in db.CATEGORIES], width=100)}
      {field("순번", "seq", len(items) + 1, kind="number", width=70)}
      {field("검사항목 *", "item_name", width=260, required="required")}
      {field("규격", "spec_text", width=150)}
      {field("단위", "unit", width=70)}
      {field("시료수", "sample_size", 5, kind="number", width=70, min="1", max="50")}
      {field("검사방법", "method", width=140)}
      {field("측정기구", "gauge", width=130)}
      <input type="hidden" name="active" value="1">
      <button type="submit">추가</button>
    </form></div>

    <h2>CSV 일괄 등록 <span class="muted">(기준서가 Word 표가 아닌 경우의 대체 경로)</span></h2>
    <div class="card">
      <form method="post" action="/models/{model_id}/items/import" enctype="multipart/form-data" class="row">
        <div class="f"><label>CSV 파일</label>
          <input type="file" name="file" accept=".csv,text/csv" required></div>
        <div class="f"><label>기존 항목</label>
          <select name="mode" style="width:170px">
            <option value="append">추가로 등록</option>
            <option value="replace">전체 교체</option>
          </select></div>
        <button type="submit">업로드</button>
        <a href="/models/{model_id}/items/template.csv"><button type="button" class="alt">양식 CSV 받기</button></a>
      </form>
      <p class="muted">헤더: <code>{",".join(db.CSV_HEADER)}</code>
         · UTF-8 또는 CP949(한국어 Excel) 모두 읽습니다. <code>구분</code>, <code>검사항목</code>은 필수입니다.</p>
      {copy_block}
    </div>

    <h2>이 모델의 변경 이력 <span class="muted">(최근 30건)</span></h2>
    <div class="scroll"><table class="tight">
      <tr><th class="c">일시</th><th class="c">담당</th><th class="c">대상</th>
          <th class="c">행위</th><th>내용</th><th class="c">바로가기</th></tr>
      {audit_table(db.list_audit(model_id=model_id, limit=30))}
    </table></div>

    <p style="margin-top:14px">
      <a href="/reports?model_id={model_id}">이 모델의 성적서 보기 →</a> &nbsp;·&nbsp;
      <a href="/history?model_id={model_id}">전체 변경 이력 →</a>
    </p>
    """
    return page(m["model_name"], body, "/models", request)


@app.post("/models/{model_id}/edit")
async def model_edit(model_id: int, request: Request):
    form = dict(await request.form())
    try:
        db.update_model(model_id, form, actor=actor_of(request))
    except sqlite3.IntegrityError:
        return see(f"/models/{model_id}", err="같은 모델명·품번·REV 조합이 이미 있습니다.")
    return see(f"/models/{model_id}", msg="모델정보를 저장했습니다.")


@app.post("/models/{model_id}/delete")
def model_delete(model_id: int, request: Request):
    try:
        db.delete_model(model_id, actor=actor_of(request))
    except ValueError as e:
        return see(f"/models/{model_id}", err=str(e))
    return see("/models", msg="모델을 삭제했습니다.")


# ---------------------------------------------------------------- 검사항목

@app.post("/models/{model_id}/items/new")
async def item_new(model_id: int, request: Request):
    form = dict(await request.form())
    if not str(form.get("item_name", "")).strip():
        return see(f"/models/{model_id}", err="검사항목명은 필수입니다.")
    db.create_item(model_id, form, actor=actor_of(request))
    return see(f"/models/{model_id}", msg="검사항목을 추가했습니다.")


@app.post("/items/{item_id}/edit")
async def item_edit(item_id: int, request: Request):
    form = dict(await request.form())
    form.setdefault("active", "0")
    with db.tx() as con:
        row = con.execute("SELECT model_id FROM inspect_item WHERE id=?", (item_id,)).fetchone()
    if row is None:
        return see("/models", err="검사항목을 찾을 수 없습니다.")
    db.update_item(item_id, form, actor=actor_of(request))
    return see(f"/models/{row['model_id']}", msg="검사항목을 저장했습니다.")


@app.post("/items/{item_id}/delete")
def item_delete(item_id: int, request: Request):
    with db.tx() as con:
        row = con.execute("SELECT model_id FROM inspect_item WHERE id=?", (item_id,)).fetchone()
    if row is None:
        return see("/models", err="검사항목을 찾을 수 없습니다.")
    db.delete_item(item_id, actor=actor_of(request))
    return see(f"/models/{row['model_id']}", msg="검사항목을 삭제했습니다.")


@app.post("/models/{model_id}/items/import")
async def item_import(model_id: int, request: Request, mode: str = Form("append"),
                      file: UploadFile = File(...)):
    raw = await file.read()
    n, warns = db.import_items_csv(model_id, raw, replace=(mode == "replace"),
                                  actor=actor_of(request))
    if not n and warns:
        return see(f"/models/{model_id}", err=warns[0])
    msg = f"{n}건을 등록했습니다."
    if warns:
        msg += f" 확인 필요 {len(warns)}건: " + " / ".join(warns[:3])
    return see(f"/models/{model_id}", msg=msg)


@app.get("/models/{model_id}/items/template.csv")
def item_template(model_id: int):
    sample = [
        ["외관", "1", "외관 상태 (스크래치·이물)", "이상 없음", "", "육안검사", "-", "5"],
        ["치수", "1", "전장", "25.40±0.20", "mm", "측정", "캘리퍼스", "5"],
        ["치수", "2", "두께", "1.00~1.20", "mm", "측정", "마이크로미터", "5"],
        ["기능", "1", "접촉저항", "MAX 50", "mΩ", "측정", "LCR 미터", "5"],
        ["도금", "1", "Au 도금두께", "0.05 이상", "um", "측정", "형광X선", "3"],
        ["추가", "1", "납땜성", "이상 없음", "", "시험", "-", "3"],
    ]
    lines = [",".join(db.CSV_HEADER)] + [",".join(r) for r in sample]
    body = "﻿" + "\r\n".join(lines) + "\r\n"
    return Response(
        body.encode("utf-8"), media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="inspect_items_template.csv"'},
    )


@app.post("/models/{model_id}/items/copy")
def item_copy(model_id: int, request: Request, src_model_id: int = Form(...)):
    n = db.copy_items(int(src_model_id), model_id, actor=actor_of(request))
    return see(f"/models/{model_id}", msg=f"검사항목 {n}건을 복사했습니다.")


# ------------------------------------------------- 기준서(Word) 업로드 · 이력

def steps(active: int) -> str:
    names = ["① 모델 등록", "② 기준서(Word) 업로드", "③ 열 매핑 확인", "④ 검사항목 적용",
             "⑤ 성적서 작성", "⑥ Excel 출력"]
    return '<ul class="steps">' + "".join(
        f'<li class="{"on" if i == active else ("done" if i < active else "")}">{esc(n)}</li>'
        for i, n in enumerate(names)
    ) + "</ul>"


def doc_label(d) -> str:
    bits = [d["filename"]]
    if d["doc_no"]:
        bits.append(d["doc_no"])
    if d["rev"]:
        bits.append("REV " + d["rev"])
    return " · ".join(bits)


@app.get("/specs", response_class=HTMLResponse)
def specs_page(request: Request, model_id: int = 0):
    docs = db.list_spec_docs(model_id or None)
    models = db.list_models()
    rows = "".join(
        f"<tr><td class=\"c\">{esc(d['uploaded_at'])}</td>"
        f"<td><a href=\"/models/{d['model_id']}\">{esc(d['model_name'])}</a>"
        f"<div class=\"muted\">{esc(d['part_no'])}</div></td>"
        f"<td><a href=\"/specs/{d['id']}/file\">{esc(d['filename'])}</a>"
        + (' <span class="badge cur">현행</span>' if d["is_current"] else "")
        + f"</td><td class=\"c\">{esc(d['doc_no'] or '-')}</td>"
        f"<td class=\"c\">{esc(d['rev'] or '-')}</td>"
        f"<td class=\"c\">{d['item_count'] or 0}</td>"
        f"<td class=\"c\">{esc(d['applied_at'] or '미적용')}</td>"
        f"<td class=\"c\">{esc(d['uploaded_by'] or '-')}</td>"
        f"<td class=\"c\"><a href=\"/specs/{d['id']}/map\">열 매핑</a></td></tr>"
        for d in docs
    ) or '<tr><td colspan="9" class="c muted">업로드된 기준서가 없습니다.</td></tr>'

    opts = [("", "전체")] + [(str(m["id"]), f"{m['model_name']} / {m['part_no']}") for m in models]
    body = f"""
    <h1>기준서 개정 이력</h1>
    <p class="sub">업로드한 출하검사 기준서(Word) 원본과, 거기서 추출해 적용한 검사항목 이력입니다.
       같은 파일을 다시 올리면 해시로 동일본임을 알려 줍니다.</p>

    <div class="card"><form method="get" class="row">
      {select("모델", "model_id", opts, str(model_id or ""), width=320)}
      <button type="submit">조회</button>
      <a href="/specs"><button type="button" class="alt">초기화</button></a>
    </form></div>

    <div class="scroll"><table>
      <tr><th class="c">업로드</th><th>모델</th><th>파일</th><th class="c">문서번호</th>
          <th class="c">REV</th><th class="c">적용항목</th><th class="c">적용일시</th>
          <th class="c">담당</th><th class="c">작업</th></tr>
      {rows}
    </table></div>
    <p class="muted">모델 화면에서 기준서를 업로드합니다. 파일은 <code>data/spec_docs/</code> 에 원본 그대로 보관됩니다.</p>
    """
    return page("기준서 이력", body, "/specs", request)


@app.post("/models/{model_id}/spec/upload")
async def spec_upload(model_id: int, request: Request, file: UploadFile = File(...),
                      doc_no: str = Form(""), rev: str = Form(""), note: str = Form("")):
    m = db.get_model(model_id)
    if m is None:
        return see("/models", err="모델을 찾을 수 없습니다.")
    raw = await file.read()
    if not raw:
        return see(f"/models/{model_id}", err="빈 파일입니다.")
    if len(raw) > MAX_UPLOAD:
        return see(f"/models/{model_id}", err=f"파일이 너무 큽니다 ({len(raw) // 1024 // 1024}MB). 최대 20MB.")
    name = os.path.basename(file.filename or "spec.docx")
    if not name.lower().endswith(".docx"):
        return see(f"/models/{model_id}",
                   err="`.docx` 파일만 읽을 수 있습니다. 구형 `.doc` 는 Word 에서 .docx 로 저장해 주십시오.")

    try:
        paras, tables = docx_parse.read_docx(raw)
    except ValueError as e:
        return see(f"/models/{model_id}", err=str(e))
    if not tables:
        return see(f"/models/{model_id}", err="기준서에서 표를 찾지 못했습니다. 검사항목이 표로 작성된 기준서가 필요합니다.")

    digest = docx_parse.sha256(raw)
    dup = db.find_spec_doc_by_hash(model_id, digest)
    meta = docx_parse.guess_meta(paras, tables)
    pick = docx_parse.pick_item_table(tables)

    doc_id = db.create_spec_doc(model_id, {
        "doc_no": doc_no.strip() or meta.get("doc_no", ""),
        "rev": rev.strip() or meta.get("rev", ""),
        "title": meta.get("title", "") or (paras[0] if paras else ""),
        "filename": name, "sha256": digest, "size": len(raw),
        "table_index": pick if pick >= 0 else None, "note": note.strip(),
    }, actor=actor_of(request))

    # 원본 보관에 실패하면 이력만 남는 고아 레코드가 되므로 되돌린다.
    try:
        os.makedirs(SPEC_DIR, exist_ok=True)
        with open(spec_path(doc_id), "wb") as fh:
            fh.write(raw)
    except OSError as e:
        db.delete_spec_doc(doc_id, actor=actor_of(request))
        return see(f"/models/{model_id}", err=f"기준서 원본 보관에 실패했습니다: {e}")

    warn = f" 주의: 동일 파일이 {dup['uploaded_at']} 에 이미 업로드되었습니다." if dup else ""
    return see(f"/specs/{doc_id}/map",
               msg=f"기준서를 등록했습니다. 표 {len(tables)}개 중 검사항목 표를 추정했습니다.{warn}")


@app.get("/specs/{doc_id}/map", response_class=HTMLResponse)
def spec_map(doc_id: int, request: Request, table_index: int = -1):
    d = db.get_spec_doc(doc_id)
    if d is None:
        return see("/specs", err="기준서를 찾을 수 없습니다.")
    path = spec_path(doc_id)
    if not os.path.exists(path):
        return see("/specs", err="보관된 원본 파일이 없습니다. 기준서를 다시 업로드해 주십시오.")
    with open(path, "rb") as fh:
        raw = fh.read()
    try:
        paras, tables = docx_parse.read_docx(raw)
    except ValueError as e:
        return see("/specs", err=str(e))

    ti = table_index if 0 <= table_index < len(tables) else (
        d["table_index"] if d["table_index"] is not None and 0 <= d["table_index"] < len(tables) else 0)
    grid = tables[ti]
    a = docx_parse.analyse_table(grid)
    header_row = a["header_row"]
    mapping = a["mapping"]

    tabs = " ".join(
        f'<a href="/specs/{doc_id}/map?table_index={i}">'
        + (f"<b>[표 {i + 1}]</b>" if i == ti else f"표 {i + 1}")
        + f'<span class="muted">({len(t)}행'
        + (", 검사항목 추정" if docx_parse.analyse_table(t)["is_item_table"] else "")
        + ")</span></a>"
        for i, t in enumerate(tables)
    )

    role_opts = [("", "— 사용 안 함 —")] + [(k, docx_parse.ROLE_LABEL[k]) for k in docx_parse.ROLE_KEYS]
    map_cells = "".join(
        f'<td class="c">{select("", f"map_{i}", role_opts, mapping[i] if i < len(mapping) else "", width=108)}</td>'
        for i in range(a["width"])
    )
    hdr_cells = "".join(
        f'<td>{esc((grid[header_row][i] if i < len(grid[header_row]) else "").replace(chr(10), " / "))}</td>'
        for i in range(a["width"])
    )
    prev_rows = "".join(
        "<tr>" + "".join(
            f'<td>{esc((r[i] if i < len(r) else "").replace(chr(10), " / "))}</td>'
            for i in range(a["width"])
        ) + "</tr>"
        for r in grid[header_row + 1: header_row + 9]
    )

    items = docx_parse.rows_to_items(grid, header_row, mapping)
    by_cat: dict[str, int] = {}
    for it in items:
        by_cat[it["category"]] = by_cat.get(it["category"], 0) + 1
    item_prev = "".join(
        f"<tr><td class=\"c\">{esc(it['category'])}</td><td class=\"c\">{it['seq']}</td>"
        f"<td>{esc(it['item_name'])}</td><td>{esc(it['spec_text'] or '-')}</td>"
        f"<td class=\"c\">{esc(db.guess_judge_type(it['spec_text']) == 'numeric' and '수치' or 'OX')}</td>"
        f"<td class=\"c\">{esc(it['unit'] or '-')}</td>"
        f"<td class=\"c\">{it['sample_size']}</td>"
        f"<td>{esc(it['method'] or '-')}</td><td>{esc(it['gauge'] or '-')}</td></tr>"
        for it in items[:25]
    ) or '<tr><td colspan="9" class="c muted">추출된 항목이 없습니다. 열 매핑을 확인해 주십시오.</td></tr>'

    existing = len(db.list_items(d["model_id"]))
    summary = " · ".join(f"{k} {v}건" for k, v in by_cat.items()) or "없음"

    body = f"""
    {steps(2)}
    <h1>열 매핑 확인 <span class="muted">· {esc(d["filename"])}</span></h1>
    <p class="sub">모델 <a href="/models/{d['model_id']}">{esc(d["model_name"])} / {esc(d["part_no"])}</a>
       · 문서번호 {esc(d["doc_no"] or "-")} · REV {esc(d["rev"] or "-")}
       · 업로드 {esc(d["uploaded_at"])} · <a href="/specs/{doc_id}/file">원본 내려받기</a></p>

    <div class="card">
      <div><b>기준서 내 표</b> {tabs}</div>
      <p class="muted" style="margin:8px 0 0">
        머리글 행은 {header_row + 1}번째 행으로 인식했습니다. 자동 추정이 틀렸으면 아래 드롭다운을 직접 바꿉니다.
        <b>검사항목</b>과 <b>규격</b>은 반드시 지정해야 합니다.</p>
    </div>

    <form method="post" action="/specs/{doc_id}/apply">
      <input type="hidden" name="table_index" value="{ti}">
      <input type="hidden" name="header_row" value="{header_row}">
      <h2>표 {ti + 1} · 열 역할 지정</h2>
      <div class="scroll"><table class="prev">
        <tr>{map_cells}</tr>
        <tr class="hdr">{hdr_cells}</tr>
        {prev_rows}
      </table></div>
      <p class="muted">위 표는 기준서 원문 상위 8행 미리보기입니다. 세로병합으로 비어 있는 <b>구분</b> 값은 위 값을 이어받습니다.</p>

      <h2>추출 결과 미리보기 · 총 {len(items)}건 ({esc(summary)})</h2>
      <div class="scroll"><table class="tight">
        <tr><th class="c">구분</th><th class="c">순번</th><th>검사항목</th><th>규격</th>
            <th class="c">판정</th><th class="c">단위</th><th class="c">시료수</th>
            <th>검사방법</th><th>측정기구</th></tr>
        {item_prev}
      </table></div>
      {'<p class="muted">… 앞 25건만 표시했습니다.</p>' if len(items) > 25 else ''}

      <div class="card" style="margin-top:16px">
        <div class="row">
          {select("기존 검사항목", "mode", [("replace", f"전체 교체 (기존 {existing}건 삭제)"),
                                          ("append", "추가로 등록")], width=240)}
          <button type="submit">이 매핑으로 검사항목 적용</button>
          <a href="/specs/{doc_id}/map?table_index={ti}"><button type="button" class="alt">다시 분석</button></a>
        </div>
        <p class="muted" style="margin:10px 0 0">
          적용하면 이 기준서가 해당 모델의 <b>현행</b> 기준서가 되고, 이후 생성되는 성적서에 근거로 기록됩니다.
          기존 검사항목을 교체해도 <b>이미 발행된 성적서는 변하지 않습니다</b>(스냅샷).</p>
      </div>
    </form>

    <form method="post" action="/specs/{doc_id}/delete" style="margin-top:14px">
      <button type="submit" class="danger"
        onclick="return confirm('이 기준서 이력을 삭제합니다. 계속하시겠습니까?')">기준서 이력 삭제</button>
      <span class="muted">이 기준서를 근거로 발행된 성적서가 있으면 삭제되지 않습니다.</span>
    </form>
    """
    return page("열 매핑", body, "/specs", request)


@app.post("/specs/{doc_id}/apply")
async def spec_apply(doc_id: int, request: Request):
    d = db.get_spec_doc(doc_id)
    if d is None:
        return see("/specs", err="기준서를 찾을 수 없습니다.")
    path = spec_path(doc_id)
    if not os.path.exists(path):
        return see("/specs", err="보관된 원본 파일이 없습니다.")

    form = await request.form()
    with open(path, "rb") as fh:
        paras, tables = docx_parse.read_docx(fh.read())

    try:
        ti = int(form.get("table_index") or 0)
        header_row = int(form.get("header_row") or 0)
    except ValueError:
        return see(f"/specs/{doc_id}/map", err="표 정보를 읽을 수 없습니다.")
    if not 0 <= ti < len(tables):
        return see(f"/specs/{doc_id}/map", err="선택한 표를 찾을 수 없습니다.")

    grid = tables[ti]
    width = max((len(r) for r in grid), default=0)
    mapping = [str(form.get(f"map_{i}", "") or "") for i in range(width)]
    if "item_name" not in mapping:
        return see(f"/specs/{doc_id}/map?table_index={ti}", err="'검사항목' 열을 지정해 주십시오.")
    if "spec_text" not in mapping:
        return see(f"/specs/{doc_id}/map?table_index={ti}", err="'규격' 열을 지정해 주십시오.")

    items = docx_parse.rows_to_items(grid, header_row, mapping)
    if not items:
        return see(f"/specs/{doc_id}/map?table_index={ti}",
                   err="추출된 검사항목이 없습니다. 머리글 행과 열 매핑을 확인해 주십시오.")

    replace = str(form.get("mode", "replace")) == "replace"
    n = db.apply_items(d["model_id"], items, spec_doc_id=doc_id, replace=replace,
                       source=f"기준서 '{d['filename']}' 표{ti + 1} 추출"
                              f" ({'전체 교체' if replace else '추가'})",
                       actor=actor_of(request))
    ox = sum(1 for it in items if db.guess_judge_type(it["spec_text"]) == "ox")
    extra = f" 이 중 {ox}건은 규격에서 수치 한계를 찾지 못해 OX 판정입니다." if ox else ""
    return see(f"/models/{d['model_id']}",
               msg=f"기준서에서 검사항목 {n}건을 적용했습니다.{extra}")


@app.get("/specs/{doc_id}/file")
def spec_file(doc_id: int):
    from urllib.parse import quote
    d = db.get_spec_doc(doc_id)
    if d is None:
        return see("/specs", err="기준서를 찾을 수 없습니다.")
    path = spec_path(doc_id)
    if not os.path.exists(path):
        return see("/specs", err="보관된 원본 파일이 없습니다.")
    with open(path, "rb") as fh:
        raw = fh.read()
    return Response(
        raw,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition":
                 f"attachment; filename*=UTF-8''{quote(d['filename'] or 'spec.docx')}"},
    )


@app.post("/specs/{doc_id}/delete")
def spec_delete(doc_id: int, request: Request):
    d = db.get_spec_doc(doc_id)
    if d is None:
        return see("/specs", err="기준서를 찾을 수 없습니다.")
    try:
        db.delete_spec_doc(doc_id, actor=actor_of(request))
    except ValueError as e:
        return see(f"/specs/{doc_id}/map", err=str(e))
    path = spec_path(doc_id)
    if os.path.exists(path):
        os.remove(path)
    return see(f"/models/{d['model_id']}", msg="기준서 이력을 삭제했습니다.")


# -------------------------------------------------------------------- 변경 이력

_ACTION_LABEL = {"create": "생성", "update": "수정", "delete": "삭제", "import": "일괄적용",
                 "upload": "업로드", "judge": "판정", "status": "상태변경", "export": "Excel출력"}
_ENTITY_LABEL = {"model": "모델", "inspect_item": "검사항목", "spec_doc": "기준서",
                 "report": "성적서", "report_line": "측정값"}


def audit_table(rows, show_model: bool = True) -> str:
    out = []
    for r in rows:
        link = ""
        if r["report_id"]:
            link = f'<a href="/reports/{r["report_id"]}">성적서</a>'
        elif r["entity"] == "spec_doc" and r["entity_id"]:
            link = f'<a href="/specs/{r["entity_id"]}/map">기준서</a>'
        elif r["model_id"]:
            link = f'<a href="/models/{r["model_id"]}">모델</a>'
        detail = ""
        if r["detail_json"]:
            detail = (f'<details><summary class="muted">상세</summary>'
                      f'<pre style="font-size:11px;white-space:pre-wrap;margin:4px 0">'
                      f'{esc(r["detail_json"][:1500])}</pre></details>')
        out.append(
            f"<tr><td class=\"c muted\" style=\"white-space:nowrap\">{esc(r['at'])}</td>"
            f"<td class=\"c\">{esc(r['actor'] or '-')}</td>"
            f"<td class=\"c\">{esc(_ENTITY_LABEL.get(r['entity'], r['entity']))}</td>"
            f"<td class=\"c\">{esc(_ACTION_LABEL.get(r['action'], r['action']))}</td>"
            f"<td>{esc(r['summary'])}{detail}</td>"
            + (f'<td class="c">{link}</td>' if show_model else "")
            + "</tr>"
        )
    cols = 6 if show_model else 5
    return "".join(out) or f'<tr><td colspan="{cols}" class="c muted">이력이 없습니다.</td></tr>'


@app.get("/history", response_class=HTMLResponse)
def history_page(request: Request, model_id: int = 0, entity: str = "", limit: int = 200):
    rows = db.list_audit(model_id or None, None, entity, limit)
    models = db.list_models()
    opts = [("", "전체")] + [(str(m["id"]), f"{m['model_name']} / {m['part_no']}") for m in models]
    ent_opts = [("", "전체")] + [(k, v) for k, v in _ENTITY_LABEL.items()]

    body = f"""
    <h1>변경 이력</h1>
    <p class="sub">모델·기준서·검사항목·성적서에 일어난 모든 변경을 누가·언제·무엇을 기준으로 남깁니다.</p>

    <div class="card"><form method="get" class="row">
      {select("모델", "model_id", opts, str(model_id or ""), width=300)}
      {select("대상", "entity", ent_opts, entity, width=130)}
      {select("표시 건수", "limit", [("100", "100건"), ("200", "200건"), ("500", "500건"),
                                   ("2000", "2000건")], str(limit), width=110)}
      <button type="submit">조회</button>
      <a href="/history"><button type="button" class="alt">초기화</button></a>
    </form></div>

    <div class="scroll"><table class="tight">
      <tr><th class="c">일시</th><th class="c">담당</th><th class="c">대상</th>
          <th class="c">행위</th><th>내용</th><th class="c">바로가기</th></tr>
      {audit_table(rows)}
    </table></div>
    <p class="muted">{len(rows)}건 표시 · 이력은 삭제되지 않습니다.</p>
    """
    return page("변경 이력", body, "/history", request)


# -------------------------------------------------------------------- 성적서

@app.get("/reports", response_class=HTMLResponse)
def reports_page(request: Request, model_id: int = 0, status: str = "", q: str = ""):
    models = db.list_models()
    reports = db.list_reports(model_id or None, status, q)

    rows = "".join(
        f"<tr><td><a href=\"/reports/{r['id']}\">{esc(r['report_no'])}</a></td>"
        f"<td>{esc(r['model_name'])}</td><td>{esc(r['part_no'])}</td>"
        f"<td class=\"c\">{esc(r['inspect_date'] or '-')}</td>"
        f"<td class=\"c\">{esc(r['lot_no'] or '-')}</td>"
        f"<td class=\"c\">{esc(r['inspector'] or '-')}</td>"
        f"<td class=\"c\">{esc(r['status'])}</td>"
        f"<td class=\"c\">{judge_tag(r['overall'])}</td>"
        f"<td class=\"c\">{r['line_count'] - r['pending_count']} / {r['line_count']}</td>"
        f"<td class=\"c\"><a href=\"/reports/{r['id']}/excel\">Excel</a></td></tr>"
        for r in reports
    ) or '<tr><td colspan="10" class="c muted">조건에 맞는 성적서가 없습니다.</td></tr>'

    model_opts = [("", "전체")] + [(str(m["id"]), f"{m['model_name']} / {m['part_no']}") for m in models]
    new_opts = [(str(m["id"]), f"{m['model_name']} / {m['part_no']} ({m['item_count']}항목)")
                for m in models if m["item_count"]]

    new_form = (
        '<div class="msg err">검사항목이 등록된 모델이 없습니다. '
        '<a href="/models">모델 · 검사항목 마스터</a>에서 먼저 등록해 주십시오.</div>'
        if not new_opts else f"""
        <div class="card"><form method="post" action="/reports/new" class="row">
          {select("모델 *", "model_id", new_opts, width=330, required="required")}
          {field("검사일", "inspect_date", date.today().isoformat(), kind="date", width=140)}
          {field("성적서번호", "report_no", "", width=150, placeholder="비우면 자동 생성")}
          {field("LOT NO", "lot_no", width=140)}
          {field("LOT 수량", "lot_qty", kind="number", width=100, min="0")}
          {field("시료수", "sample_qty", kind="number", width=90, min="0")}
          {field("검사자", "inspector", width=110)}
          {field("승인자", "approver", width=110)}
          <button type="submit">성적서 생성</button>
        </form>
        <p class="muted">생성 시점의 검사항목 마스터가 성적서에 복사(스냅샷)됩니다.
           이후 마스터 규격이 바뀌어도 발행된 성적서는 변하지 않습니다.</p></div>"""
    )

    body = f"""
    <h1>성적서</h1>
    <p class="sub">모델별 검사항목을 자동 반영한 성적서를 작성하고 Excel 로 출력합니다.</p>

    <div class="card"><form method="get" class="row">
      {select("모델", "model_id", model_opts, str(model_id or ""), width=300)}
      {select("상태", "status", [("", "전체")] + [(s, s) for s in db.REPORT_STATUS], status, width=120)}
      {field("검색 (성적서번호 / LOT / 모델)", "q", q, width=240)}
      <button type="submit">조회</button>
      <a href="/reports"><button type="button" class="alt">초기화</button></a>
    </form></div>

    <div class="scroll"><table>
      <tr><th>성적서번호</th><th>모델명</th><th>품번</th><th class="c">검사일</th><th class="c">LOT</th>
          <th class="c">검사자</th><th class="c">상태</th><th class="c">종합판정</th>
          <th class="c">입력 진행</th><th class="c">출력</th></tr>
      {rows}
    </table></div>

    <h2>성적서 신규 생성</h2>
    {new_form}
    """
    return page("성적서", body, "/reports", request)


@app.post("/reports/new")
async def report_new(request: Request):
    form = dict(await request.form())
    try:
        rid = db.create_report(form, actor=actor_of(request))
    except ValueError as e:
        return see("/reports", err=str(e))
    except sqlite3.IntegrityError:
        return see("/reports", err="같은 성적서번호가 이미 있습니다.")
    return see(f"/reports/{rid}", msg="성적서를 생성했습니다. 측정값을 입력해 주십시오.")


@app.get("/reports/{report_id}", response_class=HTMLResponse)
def report_detail(report_id: int, request: Request):
    r = db.get_report(report_id)
    if r is None:
        return see("/reports", err="성적서를 찾을 수 없습니다.")
    lines = db.list_lines(report_id)
    max_cols = max([ln["sample_size"] for ln in lines] or [1])
    max_cols = min(max_cols, report_excel.MAX_SAMPLE_COLS)

    head = ("".join(f'<th class="c">#{i + 1}</th>' for i in range(max_cols)))
    rows, last_cat = [], None
    for ln in lines:
        if ln["category"] != last_cat:
            last_cat = ln["category"]
            n = sum(1 for x in lines if x["category"] == last_cat)
            rows.append(f'<tr><td colspan="{7 + max_cols}" class="cat">{esc(last_cat)} · {n}항목</td></tr>')
        vals = db.line_values(ln)
        cells = []
        for i in range(max_cols):
            if i >= ln["sample_size"]:
                cells.append('<td class="c muted">-</td>')
                continue
            cls = "val ng" if ln["judge"] == "NG" else "val"
            cells.append(f'<td class="c"><input class="{cls}" name="v_{ln["id"]}_{i}" '
                         f'value="{esc(vals[i])}"></td>')
        manual = select("", f'j_{ln["id"]}', [("", "자동"), ("OK", "OK"), ("NG", "NG")], "", width=78)
        rows.append(
            f'<tr><td class="c muted">{ln["seq"]}</td>'
            f'<td>{esc(ln["item_name"])}</td>'
            f'<td class="c">{esc(ln["spec_text"] or "-")}</td>'
            f'<td class="c">{limit_text(ln)}</td>'
            f'<td class="c muted">{esc(ln["unit"] or "-")}</td>'
            + "".join(cells)
            + f'<td class="c">{judge_tag(ln["judge"])}</td>'
            f'<td class="c">{manual}</td></tr>'
        )
    line_rows = "".join(rows) or f'<tr><td colspan="{7 + max_cols}" class="c muted">검사항목이 없습니다.</td></tr>'

    pending = sum(1 for ln in lines if not ln["judge"])
    ng = sum(1 for ln in lines if ln["judge"] == "NG")

    if r["spec_doc_id"]:
        basis = (f'근거 기준서 <a href="/specs/{r["spec_doc_id"]}/file">{esc(r["doc_filename"] or "-")}</a>'
                 f' (문서번호 {esc(r["doc_no"] or "-")} / REV {esc(r["doc_rev"] or "-")},'
                 f' 업로드 {esc(r["doc_uploaded_at"] or "-")})'
                 f' · <a href="/specs/{r["spec_doc_id"]}/map">기준서 보기</a>')
    else:
        basis = ('<span class="muted">근거 기준서가 기록되지 않았습니다 '
                 '(기준서 업로드 전에 등록된 검사항목으로 작성됨).</span>')

    body = f"""
    {steps(4)}
    <h1>{esc(r["report_no"])} <span class="muted">· {esc(r["model_name"])} / {esc(r["part_no"])}</span></h1>
    <p class="sub">SEC CODE {esc(r["sec_code"] or "-")} · REV {esc(r["rev"] or "-")}
       · 최종수정 {esc(r["updated_at"])}<br>{basis}</p>

    <div class="stat">
      <div><b>{len(lines)}</b><span>검사항목</span></div>
      <div><b>{len(lines) - pending}</b><span>판정완료</span></div>
      <div><b>{ng}</b><span>NG</span></div>
      <div><b>{esc(r["overall"] or "미판정")}</b><span>종합판정</span></div>
    </div>

    <div class="card"><form method="post" action="/reports/{report_id}/header" class="row">
      {field("검사일", "inspect_date", r["inspect_date"], kind="date", width=140)}
      {field("LOT NO", "lot_no", r["lot_no"], width=140)}
      {field("LOT 수량", "lot_qty", r["lot_qty"], kind="number", width=100, min="0")}
      {field("시료수", "sample_qty", r["sample_qty"], kind="number", width=90, min="0")}
      {field("검사자", "inspector", r["inspector"], width=110)}
      {field("승인자", "approver", r["approver"], width=110)}
      {select("상태", "status", [(s, s) for s in db.REPORT_STATUS], r["status"], width=120)}
      {field("비고", "remark", r["remark"], width=240)}
      <button type="submit">머리정보 저장</button>
    </form></div>

    <h2>측정값 입력 · 자동 판정</h2>
    <form method="post" action="/reports/{report_id}/values">
      <div class="scroll"><table class="tight">
        <tr><th class="c">순번</th><th>검사항목</th><th class="c">규격</th><th class="c">하한~상한</th>
            <th class="c">단위</th>{head}<th class="c">판정</th><th class="c">수동지정</th></tr>
        {line_rows}
      </table></div>
      <p style="margin-top:12px">
        <button type="submit">측정값 저장 · 판정</button>
        <a href="/reports/{report_id}/excel"><button type="button" class="alt">Excel 다운로드</button></a>
      </p>
      <p class="muted">빈 칸은 미입력으로 보고 판정하지 않습니다. 하한·상한이 없는 항목과
         OX 항목 중 자동판정이 되지 않는 경우에는 <b>수동지정</b>으로 OK/NG 를 직접 넣습니다.</p>
    </form>

    <h2>이 성적서의 이력</h2>
    <div class="scroll"><table class="tight">
      <tr><th class="c">일시</th><th class="c">담당</th><th class="c">대상</th>
          <th class="c">행위</th><th>내용</th></tr>
      {audit_table(db.list_audit(report_id=report_id, limit=50), show_model=False)}
    </table></div>

    <form method="post" action="/reports/{report_id}/delete" style="margin-top:18px">
      <button type="submit" class="danger"
        onclick="return confirm('이 성적서를 삭제합니다. 계속하시겠습니까?')">성적서 삭제</button>
      <span class="muted">승인 상태의 성적서는 삭제되지 않습니다.</span>
    </form>
    """
    return page(r["report_no"], body, "/reports", request)


@app.post("/reports/{report_id}/header")
async def report_header(report_id: int, request: Request):
    db.update_report(report_id, dict(await request.form()), actor=actor_of(request))
    return see(f"/reports/{report_id}", msg="머리정보를 저장했습니다.")


@app.post("/reports/{report_id}/values")
async def report_values(report_id: int, request: Request):
    form = await request.form()
    values: dict[int, list[str]] = {}
    manual: dict[int, str] = {}
    for key, raw in form.multi_items():
        if key.startswith("v_"):
            _, lid, idx = key.split("_", 2)
            if lid.isdigit() and idx.isdigit():
                slot = values.setdefault(int(lid), [])
                while len(slot) <= int(idx):
                    slot.append("")
                slot[int(idx)] = str(raw)
        elif key.startswith("j_"):
            lid = key[2:]
            if lid.isdigit() and str(raw).strip():
                manual[int(lid)] = str(raw)
    result = db.save_values(report_id, values, manual, actor=actor_of(request))
    return see(f"/reports/{report_id}", msg=f"측정값을 저장했습니다. 종합판정: {result or '미판정'}")


@app.post("/reports/{report_id}/delete")
def report_delete(report_id: int, request: Request):
    try:
        db.delete_report(report_id, actor=actor_of(request))
    except ValueError as e:
        return see(f"/reports/{report_id}", err=str(e))
    return see("/reports", msg="성적서를 삭제했습니다.")


@app.get("/reports/{report_id}/excel")
def report_xlsx(report_id: int, request: Request):
    from urllib.parse import quote
    try:
        filename, content = report_excel.export_report(report_id)
    except ValueError as e:
        return see("/reports", err=str(e))
    r = db.get_report(report_id)
    if r is not None:
        db.log("export", "report", entity_id=report_id, model_id=r["model_id"],
               report_id=report_id, actor=actor_of(request),
               summary=f"Excel 성적서 출력 {filename} (판정 {r['overall'] or '미판정'})")
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )
