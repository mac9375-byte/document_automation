"""출하검사 기준서(Word .docx)에서 검사항목을 추출한다.

운영 프로세스: 개발부서 의뢰 → Word 기준서 → (여기) → Excel 성적서

ponytail: python-docx 를 설치하지 않는다. .docx 는 XML 묶음(ZIP)이고
표 구조(w:tbl / w:tr / w:tc)는 표준이 고정되어 있어 표준 라이브러리로 충분하다.

열 역할 추정은 자동으로 하되 최종 확정은 담당자가 화면에서 확인한다.
기준서 양식이 모델·작성자마다 조금씩 달라 완전 자동 매핑은 오판 위험이 있다.
"""

from __future__ import annotations

import hashlib
import io
import re
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 검사항목 표의 열 역할과, 기준서에서 쓰이는 머리글 표기들
ROLES: list[tuple[str, str, tuple[str, ...]]] = [
    ("category", "구분", ("구분", "분류", "검사구분", "항목구분", "대분류", "category", "class")),
    ("item_name", "검사항목", ("검사항목", "점검항목", "측정항목", "검사내용", "항목", "내용", "item", "검사부위")),
    ("spec_text", "규격", ("규격", "기준", "기준치", "규격치", "판정기준", "공차", "허용치", "관리기준",
                          "spec", "criteria", "standard", "치수")),
    ("unit", "단위", ("단위", "unit")),
    ("method", "검사방법", ("검사방법", "측정방법", "방법", "시험방법", "method")),
    ("gauge", "측정기구", ("측정기구", "측정기", "계측기", "사용기기", "측정구", "검사기구", "gauge", "기구")),
    ("sample_size", "시료수", ("시료수", "샘플수", "샘플", "n수", "시료", "sample", "검사수량")),
]
ROLE_KEYS = [r[0] for r in ROLES]
ROLE_LABEL = {r[0]: r[1] for r in ROLES}

_WS = re.compile(r"[\s ]+")


def _norm(s: str) -> str:
    return _WS.sub("", (s or "")).lower()


# -------------------------------------------------------------- docx 읽기

def _cell_text(tc: ET.Element) -> str:
    """셀 안의 문단을 줄바꿈으로 이어 붙인다. 중첩 표는 무시한다."""
    parts: list[str] = []
    for p in tc.findall(f"{W}p"):
        buf: list[str] = []
        for node in p.iter():
            if node.tag == f"{W}t":
                buf.append(node.text or "")
            elif node.tag == f"{W}tab":
                buf.append(" ")
            elif node.tag == f"{W}br":
                buf.append(" ")
        line = "".join(buf).strip()
        if line:
            parts.append(line)
    return "\n".join(parts)


def _grid_span(tc: ET.Element) -> int:
    pr = tc.find(f"{W}tcPr")
    if pr is None:
        return 1
    gs = pr.find(f"{W}gridSpan")
    if gs is None:
        return 1
    try:
        return max(1, int(gs.get(f"{W}val", "1")))
    except ValueError:
        return 1


def _is_vmerge_continue(tc: ET.Element) -> bool:
    pr = tc.find(f"{W}tcPr")
    if pr is None:
        return False
    vm = pr.find(f"{W}vMerge")
    if vm is None:
        return False
    return (vm.get(f"{W}val") or "continue") != "restart"


def _table_to_grid(tbl: ET.Element) -> list[list[str]]:
    """표를 문자열 2차원 배열로 바꾼다. 가로병합은 열을 채우고 세로병합은 위 값을 잇는다."""
    grid: list[list[str]] = []
    for tr in tbl.findall(f"{W}tr"):
        row: list[str] = []
        for tc in tr.findall(f"{W}tc"):
            text = "" if _is_vmerge_continue(tc) else _cell_text(tc)
            row.extend([text] + [""] * (_grid_span(tc) - 1))
        if any(c.strip() for c in row):
            grid.append(row)
    width = max((len(r) for r in grid), default=0)
    for r in grid:
        r.extend([""] * (width - len(r)))
    return grid


def read_docx(raw: bytes) -> tuple[list[str], list[list[list[str]]]]:
    """(문단 목록, 표 목록) 을 돌려준다. 표는 본문 등장 순서를 유지한다."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile as e:
        raise ValueError("Word 파일(.docx)이 아닙니다. 구형 .doc 는 Word 에서 .docx 로 저장해 주십시오.") from e
    try:
        xml = zf.read("word/document.xml")
    except KeyError as e:
        raise ValueError("word/document.xml 을 찾을 수 없습니다. 손상된 .docx 로 보입니다.") from e

    root = ET.fromstring(xml)
    body = root.find(f"{W}body")
    if body is None:
        return [], []

    paras: list[str] = []
    tables: list[list[list[str]]] = []
    for child in body:
        if child.tag == f"{W}p":
            text = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
            if text:
                paras.append(text)
        elif child.tag == f"{W}tbl":
            grid = _table_to_grid(child)
            if grid:
                tables.append(grid)
    return paras, tables


def sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


# ------------------------------------------------------- 머리정보 / 열 매핑 추정

_META_PATTERNS = {
    "doc_no": ("문서번호", "문서No", "기준서번호", "관리번호", "document no", "doc no"),
    "model_name": ("모델명", "모델", "model name", "model"),
    "part_no": ("품번", "part no", "part number", "부품번호"),
    "sec_code": ("sec code", "seccode", "sec 코드"),
    "rev": ("rev", "개정", "개정번호", "revision", "차수"),
    "title": ("제목", "문서명", "title"),
}


def guess_meta(paras: list[str], tables: list[list[list[str]]]) -> dict[str, str]:
    """기준서 상단의 '라벨: 값' 또는 2열 표에서 머리정보를 추정한다."""
    found: dict[str, str] = {}

    def take(label: str, value: str) -> None:
        key_norm = _norm(label)
        value = (value or "").strip()
        if not value or len(value) > 80:
            return
        for key, words in _META_PATTERNS.items():
            if key in found:
                continue
            if any(_norm(w) == key_norm or _norm(w) in key_norm for w in words):
                found[key] = value
                return

    for line in paras[:40]:
        for part in re.split(r"\s{2,}|\t", line):
            m = re.match(r"\s*([^:：]{1,20})\s*[:：]\s*(.+)$", part)
            if m:
                take(m.group(1), m.group(2))

    for grid in tables[:4]:
        for row in grid:
            cells = [c.strip() for c in row]
            for i in range(0, len(cells) - 1):
                if cells[i] and cells[i + 1]:
                    take(cells[i], cells[i + 1])
    return found


def _role_of(header: str) -> str:
    """머리글 한 칸이 어떤 열 역할인지 추정한다."""
    h = _norm(header)
    if not h:
        return ""
    best, best_len = "", 0
    for key, _label, words in ROLES:
        for w in words:
            wn = _norm(w)
            if wn and wn in h and len(wn) > best_len:
                best, best_len = key, len(wn)
    return best


def analyse_table(grid: list[list[str]]) -> dict:
    """표 하나를 분석해 머리글 행 위치, 열 역할 추정, 검사항목 표 점수를 돌려준다."""
    width = max((len(r) for r in grid), default=0)
    best = {"score": 0, "header_row": 0, "mapping": [""] * width}

    for hr in range(min(4, len(grid))):
        mapping = [_role_of(c) for c in grid[hr]]
        mapping += [""] * (width - len(mapping))
        # 같은 역할이 여러 열에 겹치면 앞쪽만 남긴다
        seen: set[str] = set()
        for i, role in enumerate(mapping):
            if role and role in seen:
                mapping[i] = ""
            elif role:
                seen.add(role)
        score = len(seen) * 10
        if "item_name" in seen:
            score += 25
        if "spec_text" in seen:
            score += 20
        if len(grid) - hr - 1 >= 2:
            score += 5
        if score > best["score"]:
            best = {"score": score, "header_row": hr, "mapping": mapping}

    best["width"] = width
    best["data_rows"] = max(0, len(grid) - best["header_row"] - 1)
    best["is_item_table"] = bool(
        best["score"] >= 45 and best["data_rows"] >= 1
        and "item_name" in best["mapping"] and "spec_text" in best["mapping"]
    )
    return best


def pick_item_table(tables: list[list[list[str]]]) -> int:
    """검사항목 표로 가장 유력한 표의 인덱스. 없으면 -1."""
    best_i, best_score = -1, 0
    for i, grid in enumerate(tables):
        a = analyse_table(grid)
        total = a["score"] + min(a["data_rows"], 40)
        if a["is_item_table"] and total > best_score:
            best_i, best_score = i, total
    return best_i


# ------------------------------------------------------------- 행 → 검사항목

_CATEGORY_ALIAS = {
    "외관": ("외관", "육안", "appearance", "visual", "표면"),
    "치수": ("치수", "dimension", "size", "측정", "길이", "두께"),
    "기능": ("기능", "전기", "특성", "function", "electrical", "성능"),
    "도금": ("도금", "plating", "코팅", "표면처리", "두께도금"),
    "추가": ("추가", "기타", "신뢰성", "etc", "other", "별도"),
}


def normalise_category(text: str, fallback: str = "추가") -> str:
    t = _norm(text)
    if not t:
        return fallback
    for std, words in _CATEGORY_ALIAS.items():
        if any(_norm(w) in t for w in words):
            return std
    return fallback


def rows_to_items(grid: list[list[str]], header_row: int, mapping: list[str]) -> list[dict]:
    """매핑된 열 역할에 따라 데이터 행을 검사항목 dict 목록으로 바꾼다.

    구분 열이 세로병합으로 비어 있는 경우 위 값을 이어서 채운다(forward fill).
    """
    col: dict[str, int] = {}
    for i, role in enumerate(mapping):
        if role in ROLE_KEYS and role not in col:
            col[role] = i

    if "item_name" not in col:
        return []

    items: list[dict] = []
    last_cat = ""
    seq_by_cat: dict[str, int] = {}

    for row in grid[header_row + 1:]:
        get = lambda role: (row[col[role]].strip() if role in col and col[role] < len(row) else "")
        name = " ".join(get("item_name").split())
        if not name:
            continue
        # 표 안에 다시 나타난 머리글 행(쪽 반복 머리글)은 건너뛴다
        if _role_of(name) == "item_name":
            continue

        raw_cat = get("category")
        if raw_cat:
            last_cat = raw_cat
        category = normalise_category(last_cat)
        seq_by_cat[category] = seq_by_cat.get(category, 0) + 1

        sample = re.sub(r"[^\d]", "", get("sample_size"))
        items.append({
            "category": category,
            "seq": seq_by_cat[category],
            "item_name": name,
            "spec_text": " ".join(get("spec_text").split()),
            "unit": get("unit")[:20],
            "method": " ".join(get("method").split())[:60],
            "gauge": " ".join(get("gauge").split())[:60],
            "sample_size": int(sample) if sample and int(sample) > 0 else 5,
            "active": "1",
            "source_category": last_cat,
        })
    return items
