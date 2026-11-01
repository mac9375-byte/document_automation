"""검사성적서 관리 시스템 - 데이터 계층.

운영 프로세스: 출하검사 기준서(Word) → 검사항목 추출 → 검사성적서(Excel)
시스템의 목적은 그 과정과 결과의 이력관리다.

  model        모델 마스터
  spec_doc     기준서(Word) 업로드 이력 = 개정 이력. 파일 해시로 동일본 재업로드를 구분한다.
  inspect_item 검사항목 마스터(기준서에서 추출한 백데이터). 출처 기준서를 spec_doc_id 로 기록한다.
  report       성적서. 근거 기준서 개정판을 spec_doc_id 로 고정한다.
  report_line  성적서 작성 시점의 검사항목 스냅샷 + 측정값.
  audit_log    모든 변경 이력(누가·언제·무엇을).

성적서를 스냅샷으로 두는 이유: 마스터 규격이 뒤에 바뀌어도 이미 발행된
품질기록이 소급 변경되면 안 된다.
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "inspection.db")

CATEGORIES = ["외관", "치수", "기능", "도금", "추가"]
REPORT_STATUS = ["작성중", "검토요청", "승인"]

SCHEMA = """
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS model (
    id          INTEGER PRIMARY KEY,
    model_name  TEXT NOT NULL,
    part_no     TEXT NOT NULL,
    sec_code    TEXT DEFAULT '',
    rev         TEXT DEFAULT '',
    customer    TEXT DEFAULT '',
    spec_doc_no TEXT DEFAULT '',
    note        TEXT DEFAULT '',
    created_at  TEXT NOT NULL,
    updated_at  TEXT NOT NULL,
    UNIQUE (model_name, part_no, rev)
);

CREATE TABLE IF NOT EXISTS spec_doc (
    id           INTEGER PRIMARY KEY,
    model_id     INTEGER NOT NULL REFERENCES model(id) ON DELETE CASCADE,
    doc_no       TEXT DEFAULT '',
    rev          TEXT DEFAULT '',
    title        TEXT DEFAULT '',
    filename     TEXT NOT NULL,
    sha256       TEXT NOT NULL,
    size         INTEGER NOT NULL DEFAULT 0,
    item_count   INTEGER NOT NULL DEFAULT 0,
    table_index  INTEGER,
    uploaded_by  TEXT DEFAULT '',
    uploaded_at  TEXT NOT NULL,
    applied_at   TEXT DEFAULT '',
    note         TEXT DEFAULT '',
    is_current   INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS ix_doc_model ON spec_doc(model_id, uploaded_at DESC);

CREATE TABLE IF NOT EXISTS audit_log (
    id          INTEGER PRIMARY KEY,
    at          TEXT NOT NULL,
    actor       TEXT DEFAULT '',
    entity      TEXT NOT NULL,
    entity_id   INTEGER,
    model_id    INTEGER,
    report_id   INTEGER,
    action      TEXT NOT NULL,
    summary     TEXT DEFAULT '',
    detail_json TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_audit_at ON audit_log(at DESC);
CREATE INDEX IF NOT EXISTS ix_audit_model ON audit_log(model_id, at DESC);
CREATE INDEX IF NOT EXISTS ix_audit_report ON audit_log(report_id, at DESC);

CREATE TABLE IF NOT EXISTS inspect_item (
    id          INTEGER PRIMARY KEY,
    model_id    INTEGER NOT NULL REFERENCES model(id) ON DELETE CASCADE,
    spec_doc_id INTEGER REFERENCES spec_doc(id) ON DELETE SET NULL,
    category    TEXT NOT NULL,
    seq         INTEGER NOT NULL DEFAULT 0,
    item_name   TEXT NOT NULL,
    spec_text   TEXT DEFAULT '',
    lsl         REAL,
    usl         REAL,
    unit        TEXT DEFAULT '',
    method      TEXT DEFAULT '',
    gauge       TEXT DEFAULT '',
    judge_type  TEXT NOT NULL DEFAULT 'numeric',
    sample_size INTEGER NOT NULL DEFAULT 5,
    active      INTEGER NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS ix_item_model ON inspect_item(model_id, category, seq);

CREATE TABLE IF NOT EXISTS report (
    id           INTEGER PRIMARY KEY,
    report_no    TEXT NOT NULL UNIQUE,
    model_id     INTEGER NOT NULL REFERENCES model(id),
    spec_doc_id  INTEGER REFERENCES spec_doc(id) ON DELETE SET NULL,
    inspect_date TEXT DEFAULT '',
    lot_no       TEXT DEFAULT '',
    lot_qty      INTEGER,
    sample_qty   INTEGER,
    inspector    TEXT DEFAULT '',
    approver     TEXT DEFAULT '',
    status       TEXT NOT NULL DEFAULT '작성중',
    overall      TEXT DEFAULT '',
    remark       TEXT DEFAULT '',
    created_at   TEXT NOT NULL,
    updated_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_report_model ON report(model_id, inspect_date);

CREATE TABLE IF NOT EXISTS report_line (
    id          INTEGER PRIMARY KEY,
    report_id   INTEGER NOT NULL REFERENCES report(id) ON DELETE CASCADE,
    item_id     INTEGER,
    category    TEXT NOT NULL,
    seq         INTEGER NOT NULL DEFAULT 0,
    item_name   TEXT NOT NULL,
    spec_text   TEXT DEFAULT '',
    lsl         REAL,
    usl         REAL,
    unit        TEXT DEFAULT '',
    method      TEXT DEFAULT '',
    gauge       TEXT DEFAULT '',
    judge_type  TEXT NOT NULL DEFAULT 'numeric',
    sample_size INTEGER NOT NULL DEFAULT 5,
    values_json TEXT NOT NULL DEFAULT '[]',
    judge       TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_line_report ON report_line(report_id, category, seq);
"""


def connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


@contextmanager
def tx():
    """커밋·롤백하고 반드시 닫는다.

    sqlite3.Connection 을 그대로 with 에 쓰면 커밋은 되지만 연결이 닫히지 않아
    장시간 돌리는 서버에서 파일 핸들이 쌓인다.
    """
    con = connect()
    try:
        yield con
        con.commit()
    except BaseException:
        con.rollback()
        raise
    finally:
        con.close()


def init_db() -> None:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with tx() as con:
        con.executescript(SCHEMA)


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ------------------------------------------------------------------ 변경 이력

def log(action: str, entity: str, *, entity_id: int | None = None, model_id: int | None = None,
        report_id: int | None = None, actor: str = "", summary: str = "",
        detail: dict | None = None, con: sqlite3.Connection | None = None) -> None:
    """모든 변경을 audit_log 에 남긴다. 이력관리가 이 시스템의 존재 이유다."""
    args = (now(), actor, entity, entity_id, model_id, report_id, action, summary,
            json.dumps(detail, ensure_ascii=False) if detail else "")
    sql = ("INSERT INTO audit_log (at, actor, entity, entity_id, model_id, report_id,"
           " action, summary, detail_json) VALUES (?,?,?,?,?,?,?,?,?)")
    if con is not None:
        con.execute(sql, args)
    else:
        with tx() as c:
            c.execute(sql, args)


def list_audit(model_id: int | None = None, report_id: int | None = None,
               entity: str = "", limit: int = 200) -> list[sqlite3.Row]:
    sql = "SELECT * FROM audit_log WHERE 1=1"
    args: list = []
    if model_id:
        sql += " AND model_id=?"
        args.append(model_id)
    if report_id:
        sql += " AND report_id=?"
        args.append(report_id)
    if entity:
        sql += " AND entity=?"
        args.append(entity)
    sql += " ORDER BY at DESC, id DESC LIMIT ?"
    args.append(int(limit))
    with tx() as con:
        return con.execute(sql, args).fetchall()


# ---------------------------------------------------------------- 규격 파서

_NUM = r"[-+]?\d+(?:\.\d+)?"
_TRANS = str.maketrans({
    "±": "±", "～": "~", "〜": "~", "∼": "~", "≤": "<=", "≦": "<=",
    "≥": ">=", "≧": ">=", "＋": "+", "－": "-", "／": "/", "　": " ",
})


def _round(x: float | None) -> float | None:
    """공차 덧·뺄셈에서 생기는 부동소수점 오차를 털어낸다.

    25.40+0.20 이 25.599999999999998 로 남으면 상한에 정확히 걸린 측정값
    25.60 이 NG 로 오판된다. 실제 치수 규격의 유효자리는 9자리를 넘지 않는다.
    """
    return None if x is None else round(x, 9)


def parse_spec(spec_text: str) -> tuple[float | None, float | None]:
    """기준서 규격 표기를 하한(LSL)·상한(USL) 절대값으로 변환한다.

    지원 표기: `1.25±0.10`, `1.25+0.10/-0.05`, `1.0~2.0`, `MAX 2.0`,
    `2.0 이하`, `MIN 1.0`, `1.0 이상`, `<=2.0`, `>=1.0`.
    판정 가능한 수치가 없으면 (None, None) 을 돌려준다.
    """
    if not spec_text:
        return None, None
    s = spec_text.translate(_TRANS).strip().upper().replace(",", "")

    # 1.25 +0.10 / -0.05  (비대칭 공차)
    m = re.search(rf"({_NUM})\s*\+\s*({_NUM})\s*/?\s*-\s*({_NUM})", s)
    if m:
        base, up, dn = (float(m.group(i)) for i in (1, 2, 3))
        return _round(base - abs(dn)), _round(base + abs(up))

    # 1.25 ± 0.10  (대칭 공차)
    m = re.search(rf"({_NUM})\s*±\s*({_NUM})", s)
    if m:
        base, tol = float(m.group(1)), abs(float(m.group(2)))
        return _round(base - tol), _round(base + tol)

    # 1.0 ~ 2.0  (범위)
    m = re.search(rf"({_NUM})\s*~\s*({_NUM})", s)
    if m:
        lo, hi = float(m.group(1)), float(m.group(2))
        return min(lo, hi), max(lo, hi)

    lsl = usl = None
    # 상한만: MAX 2.0 / 2.0 MAX / 2.0 이하 / <=2.0
    m = (re.search(rf"MAX\.?\s*({_NUM})", s)
         or re.search(rf"({_NUM})\s*MAX", s)
         or re.search(rf"({_NUM})\s*이하", s)
         or re.search(rf"<=\s*({_NUM})", s))
    if m:
        usl = float(m.group(1))
    # 하한만: MIN 1.0 / 1.0 MIN / 1.0 이상 / >=1.0
    m = (re.search(rf"MIN\.?\s*({_NUM})", s)
         or re.search(rf"({_NUM})\s*MIN", s)
         or re.search(rf"({_NUM})\s*이상", s)
         or re.search(rf">=\s*({_NUM})", s))
    if m:
        lsl = float(m.group(1))
    return lsl, usl


def guess_judge_type(spec_text: str) -> str:
    """규격 표기로 판정방식을 추정한다. 수치 한계가 잡히면 numeric, 아니면 ox."""
    lsl, usl = parse_spec(spec_text)
    return "numeric" if (lsl is not None or usl is not None) else "ox"


# ---------------------------------------------------------------- 판정 로직

_OK_WORDS = {"OK", "O", "양호", "합격", "PASS", "정상"}
_NG_WORDS = {"NG", "X", "불량", "불합격", "FAIL"}


def judge_values(judge_type: str, lsl, usl, values: list[str]) -> str:
    """측정값 목록을 상·하한과 비교해 OK / NG / '' (미판정) 을 돌려준다."""
    entered = [str(v).strip() for v in values if str(v).strip()]
    if not entered:
        return ""

    if judge_type == "ox":
        upper = [v.upper() for v in entered]
        if any(v in _NG_WORDS for v in upper):
            return "NG"
        if all(v in _OK_WORDS for v in upper):
            return "OK"
        return ""

    if lsl is None and usl is None:
        return ""  # 한계가 없어 자동판정 불가 - 담당자 수동 판정
    for v in entered:
        try:
            x = float(v.replace(",", ""))
        except ValueError:
            return "NG"  # 수치항목에 해석 불가한 값 - 안전측으로 NG
        if lsl is not None and x < lsl:
            return "NG"
        if usl is not None and x > usl:
            return "NG"
    return "OK"


def overall_judge(lines: list[sqlite3.Row]) -> str:
    """성적서 전체 판정. 하나라도 NG 면 불합격, 전 항목 OK 면 합격."""
    if not lines:
        return ""
    judges = [(r["judge"] or "") for r in lines]
    if "NG" in judges:
        return "불합격"
    if all(j == "OK" for j in judges):
        return "합격"
    return "판정보류"


# ---------------------------------------------------------------- 모델 마스터

MODEL_FIELDS = ("model_name", "part_no", "sec_code", "rev", "customer", "spec_doc_no", "note")


def list_models(q: str = "") -> list[sqlite3.Row]:
    sql = """
        SELECT m.*,
               (SELECT COUNT(*) FROM inspect_item i WHERE i.model_id = m.id AND i.active = 1) AS item_count,
               (SELECT COUNT(*) FROM report r WHERE r.model_id = m.id) AS report_count
        FROM model m
    """
    args: list = []
    if q:
        sql += " WHERE m.model_name LIKE ? OR m.part_no LIKE ? OR m.sec_code LIKE ?"
        args = [f"%{q}%"] * 3
    sql += " ORDER BY m.model_name, m.part_no"
    with tx() as con:
        return con.execute(sql, args).fetchall()


def get_model(model_id: int) -> sqlite3.Row | None:
    with tx() as con:
        return con.execute("SELECT * FROM model WHERE id = ?", (model_id,)).fetchone()


def create_model(data: dict, actor: str = "") -> int:
    vals = [str(data.get(f, "") or "").strip() for f in MODEL_FIELDS]
    ts = now()
    with tx() as con:
        cur = con.execute(
            f"INSERT INTO model ({','.join(MODEL_FIELDS)}, created_at, updated_at) "
            f"VALUES ({','.join('?' * len(MODEL_FIELDS))}, ?, ?)",
            (*vals, ts, ts),
        )
        model_id = int(cur.lastrowid)
        log("create", "model", entity_id=model_id, model_id=model_id, actor=actor,
            summary=f"모델 등록 {vals[0]} / {vals[1]}", con=con)
    return model_id


def update_model(model_id: int, data: dict, actor: str = "") -> None:
    vals = [str(data.get(f, "") or "").strip() for f in MODEL_FIELDS]
    with tx() as con:
        before = con.execute("SELECT * FROM model WHERE id=?", (model_id,)).fetchone()
        con.execute(
            f"UPDATE model SET {'=?, '.join(MODEL_FIELDS)}=?, updated_at=? WHERE id=?",
            (*vals, now(), model_id),
        )
        if before is not None:
            diff = {f: [before[f], v] for f, v in zip(MODEL_FIELDS, vals) if before[f] != v}
            if diff:
                log("update", "model", entity_id=model_id, model_id=model_id, actor=actor,
                    summary=f"모델정보 수정 {vals[0]} ({', '.join(diff)})", detail=diff, con=con)


def delete_model(model_id: int, actor: str = "") -> None:
    """성적서가 한 건이라도 있으면 삭제를 거부한다 (품질기록 보존)."""
    with tx() as con:
        n = con.execute("SELECT COUNT(*) FROM report WHERE model_id=?", (model_id,)).fetchone()[0]
        if n:
            raise ValueError(f"성적서 {n}건이 연결되어 삭제할 수 없습니다.")
        row = con.execute("SELECT * FROM model WHERE id=?", (model_id,)).fetchone()
        con.execute("DELETE FROM model WHERE id=?", (model_id,))
        if row is not None:
            log("delete", "model", entity_id=model_id, model_id=model_id, actor=actor,
                summary=f"모델 삭제 {row['model_name']} / {row['part_no']}", con=con)


# ------------------------------------------------------------ 검사항목 마스터

ITEM_FIELDS = ("category", "seq", "item_name", "spec_text", "unit",
               "method", "gauge", "judge_type", "sample_size", "active")


def _cat_order(category: str) -> int:
    return CATEGORIES.index(category) if category in CATEGORIES else len(CATEGORIES)


def list_items(model_id: int, only_active: bool = False) -> list[sqlite3.Row]:
    sql = "SELECT * FROM inspect_item WHERE model_id=?"
    if only_active:
        sql += " AND active=1"
    with tx() as con:
        rows = con.execute(sql, (model_id,)).fetchall()
    return sorted(rows, key=lambda r: (_cat_order(r["category"]), r["seq"], r["id"]))


def _item_values(data: dict) -> tuple:
    spec = str(data.get("spec_text", "") or "").strip()
    lsl, usl = parse_spec(spec)
    jt = str(data.get("judge_type", "") or "").strip() or guess_judge_type(spec)
    try:
        seq = int(data.get("seq") or 0)
    except ValueError:
        seq = 0
    try:
        ss = max(1, int(data.get("sample_size") or 5))
    except ValueError:
        ss = 5
    cat = str(data.get("category", "") or "").strip() or CATEGORIES[0]
    return (cat, seq, str(data.get("item_name", "") or "").strip(), spec, lsl, usl,
            str(data.get("unit", "") or "").strip(), str(data.get("method", "") or "").strip(),
            str(data.get("gauge", "") or "").strip(), jt, ss,
            1 if str(data.get("active", "1")) in ("1", "on", "true", "True") else 0)


_ITEM_COLS = ("category", "seq", "item_name", "spec_text", "lsl", "usl",
              "unit", "method", "gauge", "judge_type", "sample_size", "active")

_INSERT_ITEM = (f"INSERT INTO inspect_item (model_id, spec_doc_id, {','.join(_ITEM_COLS)}) "
                f"VALUES (?, ?, {','.join('?' * len(_ITEM_COLS))})")


def create_item(model_id: int, data: dict, spec_doc_id: int | None = None, actor: str = "") -> int:
    with tx() as con:
        cur = con.execute(_INSERT_ITEM, (model_id, spec_doc_id, *_item_values(data)))
        item_id = int(cur.lastrowid)
        log("create", "inspect_item", entity_id=item_id, model_id=model_id, actor=actor,
            summary=f"검사항목 추가 [{data.get('category', '')}] {data.get('item_name', '')}",
            con=con)
    return item_id


def update_item(item_id: int, data: dict, actor: str = "") -> None:
    with tx() as con:
        before = con.execute("SELECT * FROM inspect_item WHERE id=?", (item_id,)).fetchone()
        con.execute(
            f"UPDATE inspect_item SET {'=?, '.join(_ITEM_COLS)}=? WHERE id=?",
            (*_item_values(data), item_id),
        )
        after = con.execute("SELECT * FROM inspect_item WHERE id=?", (item_id,)).fetchone()
        diff = {}
        if before is not None and after is not None:
            diff = {k: [before[k], after[k]] for k in _ITEM_COLS if before[k] != after[k]}
        if diff:
            log("update", "inspect_item", entity_id=item_id,
                model_id=(after["model_id"] if after is not None else None), actor=actor,
                summary=f"검사항목 수정 {after['item_name'] if after is not None else ''}"
                        f" ({', '.join(diff)})",
                detail=diff, con=con)


def delete_item(item_id: int, actor: str = "") -> None:
    with tx() as con:
        row = con.execute("SELECT * FROM inspect_item WHERE id=?", (item_id,)).fetchone()
        con.execute("DELETE FROM inspect_item WHERE id=?", (item_id,))
        if row is not None:
            log("delete", "inspect_item", entity_id=item_id, model_id=row["model_id"], actor=actor,
                summary=f"검사항목 삭제 [{row['category']}] {row['item_name']}",
                detail={k: row[k] for k in _ITEM_COLS}, con=con)


CSV_HEADER = ["구분", "순번", "검사항목", "규격", "단위", "검사방법", "측정기구", "시료수"]


def import_items_csv(model_id: int, raw: bytes, replace: bool = False,
                     actor: str = "") -> tuple[int, list[str]]:
    """검사항목 CSV 일괄 등록. (등록건수, 경고목록) 을 돌려준다."""
    import csv
    import io

    text = None
    for enc in ("utf-8-sig", "cp949", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return 0, ["파일 인코딩을 해석할 수 없습니다. UTF-8 또는 CP949(한국어 Excel)로 저장해 주십시오."]

    reader = csv.DictReader(io.StringIO(text))
    cols = [c.strip() for c in (reader.fieldnames or [])]
    missing = [c for c in ("구분", "검사항목") if c not in cols]
    if missing:
        return 0, [f"필수 열이 없습니다: {', '.join(missing)} (헤더 예시: {','.join(CSV_HEADER)})"]

    rows, warns = [], []
    for n, row in enumerate(reader, start=2):
        clean = {(k or "").strip(): (v or "").strip() for k, v in row.items()}
        name = clean.get("검사항목", "")
        if not name:
            continue
        cat = clean.get("구분", "")
        if cat not in CATEGORIES:
            warns.append(f"{n}행: 구분 '{cat}' 은 표준 구분이 아닙니다 (허용: {'/'.join(CATEGORIES)}). 그대로 등록합니다.")
        spec = clean.get("규격", "")
        if spec and guess_judge_type(spec) == "ox":
            warns.append(f"{n}행 '{name}': 규격 '{spec}' 에서 수치 한계를 찾지 못해 OX 판정으로 등록합니다.")
        rows.append({
            "category": cat, "seq": clean.get("순번", n - 1), "item_name": name,
            "spec_text": spec, "unit": clean.get("단위", ""), "method": clean.get("검사방법", ""),
            "gauge": clean.get("측정기구", ""), "sample_size": clean.get("시료수", 5), "active": "1",
        })

    n = apply_items(model_id, rows, spec_doc_id=None, replace=replace, actor=actor,
                    source=f"CSV 업로드 ({'전체 교체' if replace else '추가'})")
    return n, warns


def apply_items(model_id: int, rows: list[dict], spec_doc_id: int | None = None,
                replace: bool = False, source: str = "", actor: str = "") -> int:
    """검사항목을 일괄 등록한다. 교체 시 삭제된 항목 수까지 이력에 남긴다."""
    with tx() as con:
        removed = 0
        if replace:
            removed = con.execute(
                "SELECT COUNT(*) FROM inspect_item WHERE model_id=?", (model_id,)).fetchone()[0]
            con.execute("DELETE FROM inspect_item WHERE model_id=?", (model_id,))
        con.executemany(_INSERT_ITEM,
                        [(model_id, spec_doc_id, *_item_values(r)) for r in rows])
        if spec_doc_id:
            con.execute("UPDATE spec_doc SET item_count=?, applied_at=?, is_current=0 WHERE id=?",
                        (len(rows), now(), spec_doc_id))
            con.execute("UPDATE spec_doc SET is_current=0 WHERE model_id=?", (model_id,))
            con.execute("UPDATE spec_doc SET is_current=1 WHERE id=?", (spec_doc_id,))
        log("import", "inspect_item", entity_id=spec_doc_id, model_id=model_id, actor=actor,
            summary=f"{source or '일괄 등록'} - 등록 {len(rows)}건"
                    + (f", 기존 {removed}건 삭제" if removed else ""),
            detail={"added": len(rows), "removed": removed,
                    "by_category": {c: sum(1 for r in rows if r.get("category") == c)
                                    for c in CATEGORIES}},
            con=con)
    return len(rows)


def copy_items(src_model_id: int, dst_model_id: int, actor: str = "") -> int:
    """다른 모델의 검사항목을 복사한다 (유사 모델 신규 등록 시)."""
    src = list_items(src_model_id, only_active=True)
    with tx() as con:
        con.executemany(_INSERT_ITEM,
                        [(dst_model_id, None, *(r[c] for c in _ITEM_COLS)) for r in src])
        srcm = con.execute("SELECT model_name FROM model WHERE id=?", (src_model_id,)).fetchone()
        log("import", "inspect_item", model_id=dst_model_id, actor=actor,
            summary=f"'{srcm['model_name'] if srcm else src_model_id}' 모델에서 검사항목 {len(src)}건 복사",
            con=con)
    return len(src)


# ------------------------------------------------------------ 기준서 개정 이력

def create_spec_doc(model_id: int, data: dict, actor: str = "") -> int:
    """업로드한 기준서 파일을 개정 이력으로 남긴다 (항목 적용은 별도 단계)."""
    ts = now()
    with tx() as con:
        cur = con.execute(
            "INSERT INTO spec_doc (model_id, doc_no, rev, title, filename, sha256, size,"
            " table_index, uploaded_by, uploaded_at, note) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (model_id, str(data.get("doc_no", "") or "").strip(), str(data.get("rev", "") or "").strip(),
             str(data.get("title", "") or "").strip(), str(data.get("filename", "") or "").strip(),
             str(data.get("sha256", "") or ""), int(data.get("size") or 0),
             data.get("table_index"), str(data.get("uploaded_by", "") or actor).strip(), ts,
             str(data.get("note", "") or "").strip()),
        )
        doc_id = int(cur.lastrowid)
        log("upload", "spec_doc", entity_id=doc_id, model_id=model_id, actor=actor,
            summary=f"기준서 업로드 {data.get('filename', '')}"
                    f" (문서번호 {data.get('doc_no') or '-'} / REV {data.get('rev') or '-'})",
            detail={"sha256": data.get("sha256", ""), "size": data.get("size", 0)}, con=con)
    return doc_id


def list_spec_docs(model_id: int | None = None, limit: int = 300) -> list[sqlite3.Row]:
    sql = ("SELECT d.*, m.model_name, m.part_no FROM spec_doc d"
           " JOIN model m ON m.id = d.model_id WHERE 1=1")
    args: list = []
    if model_id:
        sql += " AND d.model_id=?"
        args.append(model_id)
    sql += " ORDER BY d.uploaded_at DESC, d.id DESC LIMIT ?"
    args.append(int(limit))
    with tx() as con:
        return con.execute(sql, args).fetchall()


def get_spec_doc(doc_id: int) -> sqlite3.Row | None:
    with tx() as con:
        return con.execute(
            "SELECT d.*, m.model_name, m.part_no FROM spec_doc d"
            " JOIN model m ON m.id = d.model_id WHERE d.id=?", (doc_id,)).fetchone()


def current_spec_doc(model_id: int) -> sqlite3.Row | None:
    with tx() as con:
        return con.execute(
            "SELECT * FROM spec_doc WHERE model_id=? AND is_current=1"
            " ORDER BY uploaded_at DESC LIMIT 1", (model_id,)).fetchone()


def find_spec_doc_by_hash(model_id: int, digest: str) -> sqlite3.Row | None:
    with tx() as con:
        return con.execute(
            "SELECT * FROM spec_doc WHERE model_id=? AND sha256=? ORDER BY id DESC LIMIT 1",
            (model_id, digest)).fetchone()


def delete_spec_doc(doc_id: int, actor: str = "") -> None:
    with tx() as con:
        row = con.execute("SELECT * FROM spec_doc WHERE id=?", (doc_id,)).fetchone()
        if row is None:
            return
        n = con.execute("SELECT COUNT(*) FROM report WHERE spec_doc_id=?", (doc_id,)).fetchone()[0]
        if n:
            raise ValueError(f"이 기준서를 근거로 발행된 성적서 {n}건이 있어 삭제할 수 없습니다.")
        con.execute("DELETE FROM spec_doc WHERE id=?", (doc_id,))
        log("delete", "spec_doc", entity_id=doc_id, model_id=row["model_id"], actor=actor,
            summary=f"기준서 이력 삭제 {row['filename']}", con=con)


# ---------------------------------------------------------------------- 성적서

_SNAP_COLS = ("category", "seq", "item_name", "spec_text", "lsl", "usl",
              "unit", "method", "gauge", "judge_type", "sample_size")


def next_report_no(inspect_date: str = "") -> str:
    day = (inspect_date or date.today().isoformat()).replace("-", "")[:8] or date.today().strftime("%Y%m%d")
    prefix = f"QR-{day}-"
    with tx() as con:
        rows = con.execute("SELECT report_no FROM report WHERE report_no LIKE ?", (prefix + "%",)).fetchall()
    used = {int(r["report_no"][len(prefix):]) for r in rows if r["report_no"][len(prefix):].isdigit()}
    n = 1
    while n in used:
        n += 1
    return f"{prefix}{n:03d}"


def create_report(data: dict, actor: str = "") -> int:
    """성적서를 만들고 해당 모델의 활성 검사항목을 스냅샷으로 복사한다.

    근거 기준서 개정판(spec_doc_id)을 함께 고정해 성적서 → 기준서 추적이 되게 한다.
    """
    model_id = int(data["model_id"])
    items = list_items(model_id, only_active=True)
    if not items:
        raise ValueError("이 모델에 등록된 검사항목이 없습니다. 먼저 기준서를 올려 검사항목을 등록해 주십시오.")

    doc = current_spec_doc(model_id)
    doc_id = doc["id"] if doc is not None else None
    inspect_date = str(data.get("inspect_date", "") or "").strip()
    report_no = str(data.get("report_no", "") or "").strip() or next_report_no(inspect_date)
    ts = now()

    def _int(key):
        try:
            return int(data.get(key) or 0) or None
        except ValueError:
            return None

    with tx() as con:
        cur = con.execute(
            "INSERT INTO report (report_no, model_id, spec_doc_id, inspect_date, lot_no, lot_qty,"
            " sample_qty, inspector, approver, status, overall, remark, created_at, updated_at)"
            " VALUES (?,?,?,?,?,?,?,?,?,?,'',?,?,?)",
            (report_no, model_id, doc_id, inspect_date, str(data.get("lot_no", "") or "").strip(),
             _int("lot_qty"), _int("sample_qty"), str(data.get("inspector", "") or "").strip(),
             str(data.get("approver", "") or "").strip(), REPORT_STATUS[0],
             str(data.get("remark", "") or "").strip(), ts, ts),
        )
        report_id = int(cur.lastrowid)
        con.executemany(
            f"INSERT INTO report_line (report_id, item_id, {','.join(_SNAP_COLS)}) "
            f"VALUES (?, ?, {','.join('?' * len(_SNAP_COLS))})",
            [(report_id, it["id"], *(it[c] for c in _SNAP_COLS)) for it in items],
        )
        log("create", "report", entity_id=report_id, model_id=model_id, report_id=report_id,
            actor=actor,
            summary=f"성적서 생성 {report_no} - 검사항목 {len(items)}건 스냅샷"
                    + (f", 근거 기준서 {doc['filename']}" if doc is not None else ", 근거 기준서 없음"),
            detail={"spec_doc_id": doc_id, "lines": len(items)}, con=con)
    return report_id


def list_reports(model_id: int | None = None, status: str = "", q: str = "") -> list[sqlite3.Row]:
    sql = ("SELECT r.*, m.model_name, m.part_no, m.sec_code, m.rev,"
           " (SELECT COUNT(*) FROM report_line l WHERE l.report_id=r.id) AS line_count,"
           " (SELECT COUNT(*) FROM report_line l WHERE l.report_id=r.id AND l.judge='') AS pending_count"
           " FROM report r JOIN model m ON m.id = r.model_id WHERE 1=1")
    args: list = []
    if model_id:
        sql += " AND r.model_id=?"
        args.append(model_id)
    if status:
        sql += " AND r.status=?"
        args.append(status)
    if q:
        sql += " AND (r.report_no LIKE ? OR r.lot_no LIKE ? OR m.model_name LIKE ? OR m.part_no LIKE ?)"
        args += [f"%{q}%"] * 4
    sql += " ORDER BY r.inspect_date DESC, r.id DESC"
    with tx() as con:
        return con.execute(sql, args).fetchall()


def get_report(report_id: int) -> sqlite3.Row | None:
    with tx() as con:
        return con.execute(
            "SELECT r.*, m.model_name, m.part_no, m.sec_code, m.rev, m.customer, m.spec_doc_no,"
            " d.filename AS doc_filename, d.doc_no AS doc_no, d.rev AS doc_rev,"
            " d.uploaded_at AS doc_uploaded_at"
            " FROM report r JOIN model m ON m.id = r.model_id"
            " LEFT JOIN spec_doc d ON d.id = r.spec_doc_id WHERE r.id=?", (report_id,)
        ).fetchone()


def list_lines(report_id: int) -> list[sqlite3.Row]:
    with tx() as con:
        rows = con.execute("SELECT * FROM report_line WHERE report_id=?", (report_id,)).fetchall()
    return sorted(rows, key=lambda r: (_cat_order(r["category"]), r["seq"], r["id"]))


_HEADER_COLS = ("inspect_date", "lot_no", "lot_qty", "sample_qty", "inspector",
                "approver", "status", "remark")


def update_report(report_id: int, data: dict, actor: str = "") -> None:
    def _int(key):
        try:
            return int(data.get(key) or 0) or None
        except ValueError:
            return None

    status = str(data.get("status", "") or "").strip()
    if status not in REPORT_STATUS:
        status = REPORT_STATUS[0]
    with tx() as con:
        before = con.execute("SELECT * FROM report WHERE id=?", (report_id,)).fetchone()
        con.execute(
            "UPDATE report SET inspect_date=?, lot_no=?, lot_qty=?, sample_qty=?, inspector=?,"
            " approver=?, status=?, remark=?, updated_at=? WHERE id=?",
            (str(data.get("inspect_date", "") or "").strip(), str(data.get("lot_no", "") or "").strip(),
             _int("lot_qty"), _int("sample_qty"), str(data.get("inspector", "") or "").strip(),
             str(data.get("approver", "") or "").strip(), status,
             str(data.get("remark", "") or "").strip(), now(), report_id),
        )
        after = con.execute("SELECT * FROM report WHERE id=?", (report_id,)).fetchone()
        if before is None or after is None:
            return
        diff = {k: [before[k], after[k]] for k in _HEADER_COLS if before[k] != after[k]}
        if diff:
            action = "status" if "status" in diff else "update"
            summary = f"성적서 {after['report_no']} 머리정보 수정 ({', '.join(diff)})"
            if "status" in diff:
                summary = (f"성적서 {after['report_no']} 상태 변경"
                           f" {diff['status'][0]} → {diff['status'][1]}")
            log(action, "report", entity_id=report_id, model_id=after["model_id"],
                report_id=report_id, actor=actor, summary=summary, detail=diff, con=con)


def save_values(report_id: int, values_by_line: dict[int, list[str]],
                manual_judge: dict[int, str] | None = None, actor: str = "") -> str:
    """측정값을 저장하고 항목별·전체 판정을 갱신한다. 전체 판정 결과를 돌려준다."""
    manual_judge = manual_judge or {}
    lines = list_lines(report_id)
    updates, changed, new_ng = [], [], []
    for ln in lines:
        vals = values_by_line.get(ln["id"])
        if vals is None:
            continue
        vals = [str(v).strip() for v in vals][: ln["sample_size"]]
        auto = judge_values(ln["judge_type"], ln["lsl"], ln["usl"], vals)
        forced = (manual_judge.get(ln["id"]) or "").strip().upper()
        judge = forced if forced in ("OK", "NG") else auto
        payload = json.dumps(vals, ensure_ascii=False)
        if payload != (ln["values_json"] or "[]") or judge != (ln["judge"] or ""):
            changed.append(ln["item_name"])
            if judge == "NG":
                new_ng.append(f"{ln['item_name']} (규격 {ln['spec_text'] or '-'}, 측정 "
                              f"{', '.join(v for v in vals if v) or '-'})")
        updates.append((payload, judge, ln["id"]))

    with tx() as con:
        con.executemany("UPDATE report_line SET values_json=?, judge=? WHERE id=?", updates)
        rows = con.execute("SELECT judge FROM report_line WHERE report_id=?", (report_id,)).fetchall()
        result = overall_judge(rows)
        before = con.execute("SELECT report_no, model_id, overall FROM report WHERE id=?",
                             (report_id,)).fetchone()
        con.execute("UPDATE report SET overall=?, updated_at=? WHERE id=?", (result, now(), report_id))
        if changed and before is not None:
            log("judge", "report_line", model_id=before["model_id"], report_id=report_id, actor=actor,
                summary=f"측정값 저장 {len(changed)}항목, 종합판정 "
                        f"{before['overall'] or '미판정'} → {result or '미판정'}"
                        + (f" / NG {len(new_ng)}건" if new_ng else ""),
                detail={"changed_items": changed[:50], "ng_items": new_ng[:50]}, con=con)
    return result


def delete_report(report_id: int, actor: str = "") -> None:
    with tx() as con:
        r = con.execute("SELECT * FROM report WHERE id=?", (report_id,)).fetchone()
        if r is None:
            return
        if r["status"] == "승인":
            raise ValueError("승인된 성적서는 삭제할 수 없습니다.")
        con.execute("DELETE FROM report_line WHERE report_id=?", (report_id,))
        con.execute("DELETE FROM report WHERE id=?", (report_id,))
        log("delete", "report", entity_id=report_id, model_id=r["model_id"], actor=actor,
            summary=f"성적서 삭제 {r['report_no']} (상태 {r['status']}, 판정 {r['overall'] or '미판정'})",
            con=con)


def line_values(line: sqlite3.Row) -> list[str]:
    try:
        vals = json.loads(line["values_json"] or "[]")
    except (ValueError, TypeError):
        vals = []
    vals = [str(v) for v in vals]
    return (vals + [""] * line["sample_size"])[: line["sample_size"]]


def stats() -> dict:
    with tx() as con:
        g = lambda sql, a=(): con.execute(sql, a).fetchone()[0]
        out = {
            "models": g("SELECT COUNT(*) FROM model"),
            "items": g("SELECT COUNT(*) FROM inspect_item WHERE active=1"),
            "reports": g("SELECT COUNT(*) FROM report"),
            "pass": g("SELECT COUNT(*) FROM report WHERE overall='합격'"),
            "fail": g("SELECT COUNT(*) FROM report WHERE overall='불합격'"),
            "docs": g("SELECT COUNT(*) FROM spec_doc"),
            "events": g("SELECT COUNT(*) FROM audit_log"),
        }
        out["by_status"] = {s: g("SELECT COUNT(*) FROM report WHERE status=?", (s,)) for s in REPORT_STATUS}
    return out
