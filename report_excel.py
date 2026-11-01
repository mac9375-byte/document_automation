"""검사성적서 Excel 출력.

과제 원안의 요구사항을 그대로 구현한다.
  - 기준서에서 백데이터화한 검사항목·규격을 성적서 양식에 자동 반영
  - 외관 / 치수 / 기능 / 도금 / 추가 구분별 시트 구성
  - 검사 항목 수에 따라 시트를 자동 분할 (구분당 ROWS_PER_SHEET 초과 시)

ponytail: 사내 기존 성적서 양식(.xlsx) 원본을 받지 못해 표준 레이아웃으로 생성한다.
원본 양식을 템플릿으로 채워 넣어야 하면 build_report_workbook 만 교체하면 된다.
"""

from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import db

ROWS_PER_SHEET = 30          # 구분당 한 시트에 담는 최대 검사항목 수
MAX_SAMPLE_COLS = 10         # 시트에 표시할 최대 측정값 열 수

_THIN = Side(style="thin", color="999999")
_BOX = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
_TITLE_FILL = PatternFill("solid", fgColor="F2F2F2")
_NG_FILL = PatternFill("solid", fgColor="FFC7CE")
_OK_FONT = Font(size=10, color="1F6F3D")
_NG_FONT = Font(size=10, bold=True, color="9C0006")
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

_META = [
    ("성적서번호", "report_no"), ("검사일", "inspect_date"),
    ("모델명", "model_name"), ("품번", "part_no"),
    ("SEC CODE", "sec_code"), ("REV", "rev"),
    ("고객사", "customer"), ("LOT NO", "lot_no"),
    ("LOT 수량", "lot_qty"), ("시료수", "sample_qty"),
    ("검사자", "inspector"), ("승인자", "approver"),
    ("진행상태", "status"), ("근거 기준서", "doc_filename"),
    ("기준서 문서번호", "doc_no"), ("기준서 REV", "doc_rev"),
]


def _safe_sheet_name(name: str) -> str:
    for ch in '[]:*?/\\':
        name = name.replace(ch, "-")
    return name[:31] or "Sheet"


def _write_meta_block(ws, report, start_row: int = 1) -> int:
    """제목 + 성적서 머리정보를 쓰고 표가 시작될 행 번호를 돌려준다."""
    title = ws.cell(start_row, 1, "출하검사 성적서")
    title.font = Font(size=16, bold=True)
    title.alignment = _CENTER
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
    ws.row_dimensions[start_row].height = 26

    keys = report.keys()
    row = start_row + 1
    for i, (label, key) in enumerate(_META):
        if i % 3 == 0:
            row += 1
        col = 1 + (i % 3) * 2
        lc = ws.cell(row, col, label)
        lc.font = Font(size=10, bold=True)
        lc.fill = _TITLE_FILL
        lc.alignment = _CENTER
        lc.border = _BOX
        val = report[key] if key in keys else ""
        vc = ws.cell(row, col + 1, "" if val is None else val)
        vc.font = Font(size=10)
        vc.alignment = _LEFT
        vc.border = _BOX

    row += 1
    overall = report["overall"] or "미판정"
    for col, label in ((1, "종합판정"), (3, "비고")):
        lc = ws.cell(row, col, label)
        lc.font = Font(size=10, bold=True)
        lc.fill = _TITLE_FILL
        lc.alignment = _CENTER
        lc.border = _BOX

    jc = ws.cell(row, 2, overall)
    jc.alignment = _CENTER
    jc.border = _BOX
    jc.font = _NG_FONT if overall == "불합격" else Font(size=11, bold=True)
    if overall == "불합격":
        jc.fill = _NG_FILL

    rc = ws.cell(row, 4, report["remark"] or "")
    rc.alignment = _LEFT
    rc.border = _BOX
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
    return row + 2


def _write_item_table(ws, lines, sample_cols: int, start_row: int) -> None:
    head = ["No.", "검사항목", "규격", "단위"]
    head += ["#%d" % (i + 1) for i in range(sample_cols)]
    head += ["판정", "검사방법", "측정기구"]

    for c, text in enumerate(head, start=1):
        cell = ws.cell(start_row, c, text)
        cell.font = Font(size=10, bold=True)
        cell.fill = _HDR_FILL
        cell.alignment = _CENTER
        cell.border = _BOX
    ws.row_dimensions[start_row].height = 20

    jcol = 4 + sample_cols + 1
    for r, ln in enumerate(lines, start=start_row + 1):
        vals = db.line_values(ln)
        judge = ln["judge"] or ""
        cells = [r - start_row, ln["item_name"], ln["spec_text"] or "", ln["unit"] or ""]
        cells += [vals[i] if i < len(vals) else "" for i in range(sample_cols)]
        cells += [judge or "-", ln["method"] or "", ln["gauge"] or ""]
        for c, v in enumerate(cells, start=1):
            cell = ws.cell(r, c, v)
            cell.border = _BOX
            cell.font = Font(size=10)
            cell.alignment = _LEFT if c == 2 else _CENTER
        jc = ws.cell(r, jcol)
        if judge == "NG":
            jc.font = _NG_FONT
            jc.fill = _NG_FILL
        elif judge == "OK":
            jc.font = _OK_FONT

    widths = [5, 30, 18, 7] + [9] * sample_cols + [7, 16, 14]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.freeze_panes = ws.cell(start_row + 1, 3)
    ws.print_title_rows = "%d:%d" % (start_row, start_row)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _plan_sheets(lines) -> list[tuple[str, str, list]]:
    """구분별로 묶고 항목 수에 따라 시트를 분할한다. (시트명, 구분, 항목들) 목록."""
    groups: dict[str, list] = {}
    for ln in lines:
        groups.setdefault(ln["category"], []).append(ln)

    def order(cat: str) -> int:
        return db.CATEGORIES.index(cat) if cat in db.CATEGORIES else len(db.CATEGORIES)

    plan = []
    for cat in sorted(groups, key=order):
        items = groups[cat]
        chunks = [items[i:i + ROWS_PER_SHEET] for i in range(0, len(items), ROWS_PER_SHEET)]
        for n, chunk in enumerate(chunks, start=1):
            name = cat if len(chunks) == 1 else "%s(%d)" % (cat, n)
            plan.append((_safe_sheet_name(name), cat, chunk))
    return plan


def build_report_workbook(report, lines) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    plan = _plan_sheets(lines)

    summary = wb.create_sheet("성적서")
    row = _write_meta_block(summary, report)
    for c, text in enumerate(["구분", "검사항목 수", "OK", "NG", "미판정", "시트"], start=1):
        cell = summary.cell(row, c, text)
        cell.font = Font(size=10, bold=True)
        cell.fill = _HDR_FILL
        cell.alignment = _CENTER
        cell.border = _BOX

    seen: list[str] = []
    for _, cat, _items in plan:
        if cat not in seen:
            seen.append(cat)
    for cat in seen:
        row += 1
        names = [n for n, c, _ in plan if c == cat]
        items = [i for _, c, chunk in plan if c == cat for i in chunk]
        judges = [i["judge"] or "" for i in items]
        cells = [cat, len(items), judges.count("OK"), judges.count("NG"),
                 judges.count(""), ", ".join(names)]
        for c, v in enumerate(cells, start=1):
            cell = summary.cell(row, c, v)
            cell.border = _BOX
            cell.font = Font(size=10)
            cell.alignment = _LEFT if c in (1, 6) else _CENTER

    for c, w in enumerate([10, 13, 7, 7, 9, 28], start=1):
        summary.column_dimensions[get_column_letter(c)].width = w
    summary.page_setup.orientation = "landscape"
    summary.page_setup.paperSize = summary.PAPERSIZE_A4

    for name, _cat, items in plan:
        ws = wb.create_sheet(name)
        sizes = [i["sample_size"] for i in items] or [1]
        sample_cols = max(1, min(MAX_SAMPLE_COLS, max(sizes)))
        start = _write_meta_block(ws, report)
        ws.cell(start, 1, "■ " + name).font = Font(size=12, bold=True)
        _write_item_table(ws, items, sample_cols, start + 1)

    _write_history_sheet(wb, report)
    return wb


def _write_history_sheet(wb: Workbook, report) -> None:
    """성적서 작성 이력을 시트로 함께 내보낸다 (이력관리 목적)."""
    rows = db.list_audit(report_id=report["id"], limit=500)
    ws = wb.create_sheet("이력")
    ws.cell(1, 1, "성적서 작성 이력").font = Font(size=13, bold=True)
    ws.cell(2, 1, f"성적서 {report['report_no']} · {report['model_name']} / {report['part_no']}").font = \
        Font(size=10, color="6B7280")

    head = ["일시", "담당", "대상", "행위", "내용"]
    for c, text in enumerate(head, start=1):
        cell = ws.cell(4, c, text)
        cell.font = Font(size=10, bold=True)
        cell.fill = _HDR_FILL
        cell.alignment = _CENTER
        cell.border = _BOX

    labels = {"model": "모델", "inspect_item": "검사항목", "spec_doc": "기준서",
              "report": "성적서", "report_line": "측정값"}
    actions = {"create": "생성", "update": "수정", "delete": "삭제", "import": "일괄적용",
               "upload": "업로드", "judge": "판정", "status": "상태변경", "export": "Excel출력"}
    for r, row in enumerate(rows, start=5):
        cells = [row["at"], row["actor"] or "-", labels.get(row["entity"], row["entity"]),
                 actions.get(row["action"], row["action"]), row["summary"]]
        for c, v in enumerate(cells, start=1):
            cell = ws.cell(r, c, v)
            cell.border = _BOX
            cell.font = Font(size=10)
            cell.alignment = _LEFT if c == 5 else _CENTER
    if not rows:
        ws.cell(5, 1, "이력이 없습니다.").font = Font(size=10, color="6B7280")

    for c, w in enumerate([19, 10, 10, 10, 86], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(5, 1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4


def export_report(report_id: int) -> tuple[str, bytes]:
    """성적서 하나를 Excel 바이트로 만든다. (파일명, 내용) 을 돌려준다."""
    report = db.get_report(report_id)
    if report is None:
        raise ValueError("성적서를 찾을 수 없습니다.")
    lines = db.list_lines(report_id)
    wb = build_report_workbook(report, lines)
    buf = io.BytesIO()
    wb.save(buf)
    model = _safe_sheet_name(str(report["model_name"] or "")).replace(" ", "_")
    return "출하검사성적서_%s_%s.xlsx" % (report["report_no"], model), buf.getvalue()
