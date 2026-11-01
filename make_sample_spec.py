"""테스트·시연용 출하검사 기준서(Word .docx) 생성기.

실제 기준서를 받기 전에 업로드 → 표 분석 → 열 매핑 → 적용 흐름을 확인하기 위한 것이다.
구분 열을 세로병합(vMerge)으로 만들어 실제 기준서와 같은 조건을 재현한다.

실행:  python make_sample_spec.py            -> sample/출하검사기준서_샘플.docx
"""

from __future__ import annotations

import os
import zipfile

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def xq(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def para(text: str, bold: bool = False) -> str:
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{xq(text)}</w:t></w:r></w:p>'


def cell(text: str, vmerge: str = "") -> str:
    """vmerge: '' | 'restart' | 'continue'"""
    pr = ""
    if vmerge == "restart":
        pr = '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
    elif vmerge == "continue":
        pr = "<w:tcPr><w:vMerge/></w:tcPr>"
    body = para(text) if (text or vmerge != "continue") else "<w:p/>"
    return f"<w:tc>{pr}{body}</w:tc>"


def table(rows: list[list[tuple[str, str]]]) -> str:
    """rows: 셀마다 (텍스트, vmerge) 튜플."""
    trs = []
    for row in rows:
        tcs = "".join(cell(t, v) for t, v in row)
        trs.append(f"<w:tr>{tcs}</w:tr>")
    return "<w:tbl>" + "".join(trs) + "</w:tbl>"


# 모델 머리정보 표 (2열 라벨:값 형태 - guess_meta 가 읽는다)
META_ROWS = [
    [("문서번호", ""), ("QI-IF-2026-014", ""), ("REV", ""), ("3", "")],
    [("모델명", ""), ("IF SUB PBA A-TYPE", ""), ("품번", ""), ("GH97-28431A", "")],
    [("SEC CODE", ""), ("SEC-4412-09", ""), ("고객사", ""), ("삼성전자", "")],
]

# 검사항목 표: 구분 열은 세로병합으로 첫 행에만 값이 있다
ITEM_ROWS: list[list[tuple[str, str]]] = [
    [("구분", ""), ("검사항목", ""), ("규격", ""), ("단위", ""),
     ("검사방법", ""), ("측정기구", ""), ("시료수", "")],

    [("외관", "restart"), ("외관 상태 (스크래치·이물·오염)", ""), ("이상 없음", ""), ("", ""),
     ("육안검사", ""), ("-", ""), ("5", "")],
    [("", "continue"), ("인쇄 상태 (모델명·로트표기)", ""), ("판독 가능, 누락 없음", ""), ("", ""),
     ("육안검사", ""), ("-", ""), ("5", "")],
    [("", "continue"), ("단자 변형·휨", ""), ("이상 없음", ""), ("", ""),
     ("육안검사", ""), ("확대경 10배", ""), ("5", "")],

    [("치수", "restart"), ("전장 (L)", ""), ("25.40±0.20", ""), ("mm", ""),
     ("측정", ""), ("버니어캘리퍼스", ""), ("5", "")],
    [("", "continue"), ("전폭 (W)", ""), ("12.70±0.15", ""), ("mm", ""),
     ("측정", ""), ("버니어캘리퍼스", ""), ("5", "")],
    [("", "continue"), ("두께 (T)", ""), ("1.00~1.20", ""), ("mm", ""),
     ("측정", ""), ("마이크로미터", ""), ("5", "")],
    [("", "continue"), ("단자 피치", ""), ("0.50+0.05/-0.03", ""), ("mm", ""),
     ("측정", ""), ("투영기", ""), ("3", "")],
    [("", "continue"), ("휨량 (Warpage)", ""), ("0.10 이하", ""), ("mm", ""),
     ("측정", ""), ("3차원측정기", ""), ("3", "")],

    [("기능", "restart"), ("접촉저항", ""), ("MAX 50", ""), ("mΩ", ""),
     ("측정", ""), ("LCR 미터", ""), ("5", "")],
    [("", "continue"), ("절연저항", ""), ("100 이상", ""), ("MΩ", ""),
     ("측정", ""), ("절연저항계", ""), ("5", "")],
    [("", "continue"), ("삽입력", ""), ("5.0~15.0", ""), ("N", ""),
     ("측정", ""), ("푸시풀게이지", ""), ("5", "")],
    [("", "continue"), ("도통 검사", ""), ("전핀 도통", ""), ("", ""),
     ("검사", ""), ("도통검사기", ""), ("5", "")],

    [("도금", "restart"), ("Au 도금두께 (접점부)", ""), ("0.05 이상", ""), ("um", ""),
     ("측정", ""), ("형광X선", ""), ("3", "")],
    [("", "continue"), ("Ni 하지도금 두께", ""), ("1.0~3.0", ""), ("um", ""),
     ("측정", ""), ("형광X선", ""), ("3", "")],

    [("추가검사", "restart"), ("납땜성", ""), ("95% 이상 젖음", ""), ("", ""),
     ("시험", ""), ("솔더체커", ""), ("3", "")],
    [("", "continue"), ("내열성 (260℃ 10초)", ""), ("이상 없음", ""), ("", ""),
     ("시험", ""), ("리플로우", ""), ("3", "")],
]


def build() -> bytes:
    body = (
        para("출하검사 기준서", bold=True)
        + para("")
        + table(META_ROWS)
        + para("")
        + para("1. 검사 항목 및 규격", bold=True)
        + table(ITEM_ROWS)
        + para("")
        + para("2. 판정 기준: 전 항목 규격 만족 시 합격. 1항목이라도 규격 이탈 시 불합격 처리한다.")
        + para("3. 본 기준서는 개발부서 의뢰에 따라 작성되며, 개정 시 REV 를 상향한다.")
    )
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f"<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>")

    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", doc)
    return buf.getvalue()


def main() -> str:
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "출하검사기준서_샘플.docx")
    with open(path, "wb") as fh:
        fh.write(build())
    return path


if __name__ == "__main__":
    print(main())
