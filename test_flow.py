"""전 구간 통합 테스트 (서버 기동 없이 ASGI 앱 직접 호출).

기준서(Word) 업로드 -> 표 분석 -> 열 매핑 적용 -> 성적서 생성 ->
측정값 입력·자동판정 -> Excel 출력 -> 이력 확인 까지 한 번에 확인한다.

실행:  python test_flow.py
"""

from __future__ import annotations

import io
import os
import re
import sqlite3
import sys
import tempfile
import unittest
from urllib.parse import parse_qs, urlparse


def _fresh_env() -> str:
    """테스트용 임시 DB·보관폴더를 쓰도록 모듈을 다시 불러온다."""
    tmp = tempfile.mkdtemp(prefix="qr_test_")
    for mod in ("db", "report_excel", "docx_parse", "app"):
        sys.modules.pop(mod, None)
    import db
    db.DB_PATH = os.path.join(tmp, "test.db")
    import app
    app.SPEC_DIR = os.path.join(tmp, "spec_docs")
    os.makedirs(app.SPEC_DIR, exist_ok=True)
    db.init_db()
    return tmp


class Client:
    """의존성 없이 ASGI 앱을 호출하는 최소 테스트 클라이언트."""

    def __init__(self, app):
        self.app = app
        self.cookies: dict[str, str] = {}

    def _call(self, method: str, path: str, body: bytes = b"", content_type: str = "") -> dict:
        import asyncio
        url = urlparse(path)
        headers = [(b"host", b"testserver")]
        if content_type:
            headers.append((b"content-type", content_type.encode()))
        if body:
            headers.append((b"content-length", str(len(body)).encode()))
        if self.cookies:
            jar = "; ".join(f"{k}={v}" for k, v in self.cookies.items())
            headers.append((b"cookie", jar.encode()))

        scope = {"type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
                 "method": method, "scheme": "http", "path": url.path,
                 "raw_path": url.path.encode(), "query_string": url.query.encode(),
                 "root_path": "", "headers": headers, "client": ("127.0.0.1", 1234),
                 "server": ("testserver", 80)}

        sent = {"status": 0, "headers": [], "body": b""}
        sent_body = {"done": False}

        async def receive():
            if sent_body["done"]:
                return {"type": "http.disconnect"}
            sent_body["done"] = True
            return {"type": "http.request", "body": body, "more_body": False}

        async def send(msg):
            if msg["type"] == "http.response.start":
                sent["status"] = msg["status"]
                sent["headers"] = msg.get("headers", [])
            elif msg["type"] == "http.response.body":
                sent["body"] += msg.get("body", b"")

        asyncio.run(self.app(scope, receive, send))

        for k, v in sent["headers"]:
            if k.decode().lower() == "set-cookie":
                piece = v.decode().split(";")[0]
                if "=" in piece:
                    name, val = piece.split("=", 1)
                    self.cookies[name] = val
        loc = next((v.decode() for k, v in sent["headers"] if k.decode().lower() == "location"), "")
        ctype = next((v.decode() for k, v in sent["headers"]
                      if k.decode().lower() == "content-type"), "")
        return {"status": sent["status"], "body": sent["body"], "location": loc, "ctype": ctype}

    def get(self, path: str) -> dict:
        return self._call("GET", path)

    def post(self, path: str, data: dict) -> dict:
        from urllib.parse import urlencode
        return self._call("POST", path, urlencode(data, doseq=True).encode(),
                          "application/x-www-form-urlencoded")

    def upload(self, path: str, field: str, filename: str, content: bytes, extra: dict) -> dict:
        b = "----QRTEST1234"
        parts = []
        for k, v in extra.items():
            parts.append(f"--{b}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n")
        head = (f"--{b}\r\nContent-Disposition: form-data; name=\"{field}\"; "
                f"filename=\"{filename}\"\r\nContent-Type: application/octet-stream\r\n\r\n")
        body = ("".join(parts) + head).encode() + content + f"\r\n--{b}--\r\n".encode()
        return self._call("POST", path, body, f"multipart/form-data; boundary={b}")


def flash(resp: dict) -> str:
    """리다이렉트 URL 의 msg / err 를 꺼낸다."""
    q = parse_qs(urlparse(resp["location"]).query)
    return (q.get("msg", [""])[0] or "") + (q.get("err", [""])[0] or "")


class TestFullFlow(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = _fresh_env()
        import app as app_mod
        import db as db_mod
        import make_sample_spec
        cls.app_mod, cls.db = app_mod, db_mod
        cls.c = Client(app_mod.app)
        cls.spec_bytes = make_sample_spec.build()

    # ---------------------------------------------------------------- 1. 기본

    def test_01_pages_render(self):
        self.c.post("/actor", {"actor": "이재우"})
        for path in ("/", "/models", "/specs", "/reports", "/history"):
            r = self.c.get(path)
            self.assertEqual(r["status"], 200, f"{path} -> {r['status']}")
            self.assertIn("검사성적서 관리 시스템", r["body"].decode("utf-8"))

    def test_02_create_model(self):
        r = self.c.post("/models/new", {
            "model_name": "IF SUB PBA A-TYPE", "part_no": "GH97-28431A",
            "sec_code": "SEC-4412-09", "rev": "3", "customer": "삼성전자",
        })
        self.assertEqual(r["status"], 303)
        self.assertIn("/models/1", r["location"])
        self.assertEqual(self.c.get("/models/1")["status"], 200)

    def test_03_reject_non_docx(self):
        r = self.c.upload("/models/1/spec/upload", "file", "spec.doc", b"not a docx",
                          {"doc_no": "", "rev": "", "note": ""})
        self.assertIn(".docx", flash(r))

    def test_04_upload_spec_docx(self):
        r = self.c.upload("/models/1/spec/upload", "file", "출하검사기준서_IFSUB.docx",
                          self.spec_bytes, {"doc_no": "", "rev": "", "note": "초기 등록"})
        self.assertEqual(r["status"], 303)
        self.assertIn("/specs/1/map", r["location"])

        doc = self.db.get_spec_doc(1)
        self.assertEqual(doc["doc_no"], "QI-IF-2026-014", "문서번호가 Word 에서 추출되어야 한다")
        self.assertEqual(doc["rev"], "3", "REV 가 Word 에서 추출되어야 한다")
        self.assertTrue(os.path.exists(self.app_mod.spec_path(1)), "원본이 보관되어야 한다")

    def test_05_duplicate_upload_is_flagged(self):
        r = self.c.upload("/models/1/spec/upload", "file", "같은파일.docx",
                          self.spec_bytes, {"doc_no": "", "rev": "", "note": ""})
        self.assertIn("동일 파일", flash(r), "해시로 동일본을 알려야 한다")
        self.db.delete_spec_doc(2, actor="이재우")

    def test_05b_upload_recreates_missing_storage_dir(self):
        """보관 폴더가 없어도 업로드가 죽지 않고, 실패 시 고아 레코드를 남기지 않아야 한다."""
        import shutil
        existing = {d["id"] for d in self.db.list_spec_docs(1)}
        shutil.rmtree(self.app_mod.SPEC_DIR, ignore_errors=True)
        try:
            r = self.c.upload("/models/1/spec/upload", "file", "폴더없음.docx",
                              self.spec_bytes, {"doc_no": "", "rev": "", "note": ""})
            self.assertEqual(r["status"], 303)
            self.assertNotIn("err=", r["location"], flash(r))

            new = [d for d in self.db.list_spec_docs(1) if d["id"] not in existing]
            self.assertEqual(len(new), 1, "고아 레코드 없이 한 건만 추가되어야 한다")
            self.assertTrue(os.path.exists(self.app_mod.spec_path(new[0]["id"])),
                            "폴더가 없어도 원본이 보관되어야 한다")
            self.db.delete_spec_doc(new[0]["id"], actor="이재우")
            os.remove(self.app_mod.spec_path(new[0]["id"]))
        finally:
            # 뒤 테스트가 쓰는 기존 기준서 원본을 되돌려 놓는다
            os.makedirs(self.app_mod.SPEC_DIR, exist_ok=True)
            for doc_id in existing:
                with open(self.app_mod.spec_path(doc_id), "wb") as fh:
                    fh.write(self.spec_bytes)

    def test_06_mapping_page(self):
        r = self.c.get("/specs/1/map")
        self.assertEqual(r["status"], 200)
        html = r["body"].decode("utf-8")
        self.assertIn("열 매핑 확인", html)
        self.assertIn("총 16건", html, "16개 검사항목이 추출되어야 한다")
        for cat in ("외관", "치수", "기능", "도금"):
            self.assertIn(cat, html)

    def test_07_apply_requires_spec_column(self):
        r = self.c.post("/specs/1/apply", {
            "table_index": 1, "header_row": 0,
            "map_0": "category", "map_1": "item_name", "map_2": "",
            "map_3": "unit", "map_4": "method", "map_5": "gauge", "map_6": "sample_size",
            "mode": "replace",
        })
        self.assertIn("규격", flash(r), "규격 열 미지정은 거부되어야 한다")

    def test_08_apply_mapping(self):
        r = self.c.post("/specs/1/apply", {
            "table_index": 1, "header_row": 0,
            "map_0": "category", "map_1": "item_name", "map_2": "spec_text",
            "map_3": "unit", "map_4": "method", "map_5": "gauge", "map_6": "sample_size",
            "mode": "replace",
        })
        self.assertEqual(r["status"], 303)
        self.assertIn("16건을 적용", flash(r))

        items = self.db.list_items(1, only_active=True)
        self.assertEqual(len(items), 16)
        cats = {}
        for it in items:
            cats[it["category"]] = cats.get(it["category"], 0) + 1
        self.assertEqual(cats, {"외관": 3, "치수": 5, "기능": 4, "도금": 2, "추가": 2},
                         "세로병합 구분 열이 이어 채워져야 한다")

        length = next(i for i in items if i["item_name"].startswith("전장"))
        self.assertEqual((length["lsl"], length["usl"]), (25.2, 25.6))
        self.assertEqual(length["judge_type"], "numeric")
        self.assertEqual(length["unit"], "mm")
        self.assertEqual(length["gauge"], "버니어캘리퍼스")

        appearance = next(i for i in items if i["category"] == "외관")
        self.assertEqual(appearance["judge_type"], "ox", "수치 한계가 없으면 OX 판정")

        self.assertTrue(self.db.current_spec_doc(1)["id"] == 1, "현행 기준서로 지정되어야 한다")

    # ------------------------------------------------------------- 2. 성적서

    def test_09_create_report(self):
        r = self.c.post("/reports/new", {
            "model_id": 1, "inspect_date": "2026-11-01", "lot_no": "L2611-001",
            "lot_qty": 3000, "sample_qty": 5, "inspector": "이재우", "approver": "김과장",
        })
        self.assertEqual(r["status"], 303)
        rep = self.db.get_report(1)
        self.assertEqual(rep["report_no"], "QR-20261101-001")
        self.assertEqual(rep["spec_doc_id"], 1, "근거 기준서가 고정되어야 한다")
        self.assertEqual(len(self.db.list_lines(1)), 16, "검사항목이 스냅샷되어야 한다")
        self.assertEqual(self.c.get("/reports/1")["status"], 200)

    def test_10_enter_values_and_judge(self):
        lines = self.db.list_lines(1)
        data: dict = {}
        for ln in lines:
            if ln["judge_type"] == "ox":
                for i in range(ln["sample_size"]):
                    data[f'v_{ln["id"]}_{i}'] = "OK"
            else:
                lo = ln["lsl"] if ln["lsl"] is not None else (ln["usl"] - 1)
                hi = ln["usl"] if ln["usl"] is not None else (ln["lsl"] + 1)
                mid = (lo + hi) / 2
                for i in range(ln["sample_size"]):
                    data[f'v_{ln["id"]}_{i}'] = f"{mid:.3f}"
        r = self.c.post("/reports/1/values", data)
        self.assertIn("합격", flash(r))
        self.assertEqual(self.db.get_report(1)["overall"], "합격")

    def test_11_out_of_spec_becomes_ng(self):
        length = next(l for l in self.db.list_lines(1) if l["item_name"].startswith("전장"))
        self.c.post("/reports/1/values", {f'v_{length["id"]}_0': "25.75"})
        self.assertEqual(self.db.get_report(1)["overall"], "불합격")

        # 상한에 정확히 걸린 값은 합격 (부동소수점 오차 회귀 방지)
        self.c.post("/reports/1/values", {f'v_{length["id"]}_0': "25.600"})
        row = next(l for l in self.db.list_lines(1) if l["id"] == length["id"])
        self.assertEqual(row["judge"], "OK")
        self.assertEqual(self.db.get_report(1)["overall"], "합격")

    def test_12_manual_judge_override(self):
        ln = self.db.list_lines(1)[0]
        self.c.post("/reports/1/values", {f'v_{ln["id"]}_0': "OK", f'j_{ln["id"]}': "NG"})
        row = next(l for l in self.db.list_lines(1) if l["id"] == ln["id"])
        self.assertEqual(row["judge"], "NG", "수동지정이 자동판정을 덮어써야 한다")
        self.c.post("/reports/1/values", {f'v_{ln["id"]}_0': "OK"})

    def test_13_status_change_logged(self):
        self.c.post("/reports/1/header", {
            "inspect_date": "2026-11-01", "lot_no": "L2611-001", "lot_qty": 3000,
            "sample_qty": 5, "inspector": "이재우", "approver": "김과장",
            "status": "승인", "remark": "출하 승인",
        })
        self.assertEqual(self.db.get_report(1)["status"], "승인")
        logs = [r["summary"] for r in self.db.list_audit(report_id=1)]
        self.assertTrue(any("상태 변경" in s and "승인" in s for s in logs), logs)

    def test_14_approved_report_not_deletable(self):
        r = self.c.post("/reports/1/delete", {})
        self.assertIn("승인된 성적서는 삭제할 수 없습니다", flash(r))

    # ----------------------------------------------------------- 3. Excel 출력

    def test_15_excel_export(self):
        r = self.c.get("/reports/1/excel")
        self.assertEqual(r["status"], 200)
        self.assertIn("spreadsheetml", r["ctype"])

        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(r["body"]))
        self.assertEqual(wb.sheetnames, ["성적서", "외관", "치수", "기능", "도금", "추가", "이력"],
                         "구분별 시트가 자동 구성되어야 한다")

        cover = "\n".join(str(c.value) for row in wb["성적서"].iter_rows()
                          for c in row if c.value is not None)
        self.assertIn("QR-20261101-001", cover)
        self.assertIn("출하검사기준서_IFSUB.docx", cover, "근거 기준서가 표기되어야 한다")
        self.assertIn("QI-IF-2026-014", cover)

        dim = wb["치수"]
        text = "\n".join(str(c.value) for row in dim.iter_rows()
                         for c in row if c.value is not None)
        self.assertIn("전장 (L)", text)
        self.assertIn("25.40±0.20", text)
        self.assertIn("버니어캘리퍼스", text)
        self.assertEqual(sum(1 for row in dim.iter_rows() for c in row
                             if c.value == "OK"), 5, "치수 5항목 모두 OK 여야 한다")

        hist = "\n".join(str(c.value) for row in wb["이력"].iter_rows()
                         for c in row if c.value is not None)
        self.assertIn("성적서 생성", hist)
        self.assertIn("측정값 저장", hist)

    def test_16_sheet_split_on_many_items(self):
        """검사 항목 수가 많으면 시트를 자동 분할한다 (과제 원안 요구사항)."""
        import report_excel
        for n in range(40):
            self.db.create_item(1, {"category": "치수", "seq": 100 + n,
                                    "item_name": f"추가치수 {n + 1}", "spec_text": "1.0±0.1",
                                    "unit": "mm", "sample_size": 3, "active": "1"})
        self.c.post("/reports/new", {"model_id": 1, "inspect_date": "2026-11-02",
                                     "inspector": "이재우"})
        r = self.c.get("/reports/2/excel")
        from openpyxl import load_workbook
        names = load_workbook(io.BytesIO(r["body"])).sheetnames
        self.assertIn("치수(1)", names)
        self.assertIn("치수(2)", names)
        self.assertEqual(report_excel.ROWS_PER_SHEET, 30)

    # --------------------------------------------------------------- 4. 이력

    def test_17_audit_trail_complete(self):
        logs = self.db.list_audit(limit=500)
        actions = {r["action"] for r in logs}
        for expected in ("create", "upload", "import", "judge", "status", "export"):
            self.assertIn(expected, actions, f"'{expected}' 이력이 남아야 한다")

        # 화면(HTTP)에서 일어난 동작은 담당자가 반드시 기록되어야 한다.
        # 테스트가 db 함수를 직접 부른 건들(test_16/18)은 담당자가 없는 것이 정상이다.
        web = [r for r in logs if r["entity"] in ("spec_doc", "report")]
        self.assertTrue(web, "기준서·성적서 이력이 있어야 한다")
        for r in web:
            self.assertEqual(r["actor"], "이재우", f"담당자 누락: {r['summary']}")

        self.assertEqual(self.c.get("/history")["status"], 200)
        self.assertEqual(self.c.get("/history?model_id=1&entity=spec_doc")["status"], 200)

    def test_18_snapshot_survives_master_change(self):
        """마스터 규격을 바꿔도 이미 발행된 성적서는 변하지 않아야 한다."""
        before = next(l for l in self.db.list_lines(1) if l["item_name"].startswith("전장"))
        item = next(i for i in self.db.list_items(1) if i["item_name"].startswith("전장"))
        self.db.update_item(item["id"], {
            "category": "치수", "seq": item["seq"], "item_name": item["item_name"],
            "spec_text": "99.00±0.01", "unit": "mm", "sample_size": 5, "active": "1",
        })
        after = next(l for l in self.db.list_lines(1) if l["id"] == before["id"])
        self.assertEqual((after["lsl"], after["usl"]), (25.2, 25.6))
        self.assertEqual(after["spec_text"], "25.40±0.20")

    def test_19_model_with_reports_not_deletable(self):
        r = self.c.post("/models/1/delete", {})
        self.assertIn("삭제할 수 없습니다", flash(r))

    def test_20_spec_doc_with_reports_not_deletable(self):
        r = self.c.post("/specs/1/delete", {})
        self.assertIn("삭제할 수 없습니다", flash(r))

    def test_21_csv_path_still_works(self):
        self.c.post("/models/new", {"model_name": "TEST-CSV", "part_no": "P-001"})
        csv = ("구분,순번,검사항목,규격,단위,검사방법,측정기구,시료수\r\n"
               "치수,1,폭,10.0±0.5,mm,측정,캘리퍼스,5\r\n"
               "외관,1,외관,이상 없음,,육안,-,5\r\n").encode("utf-8-sig")
        r = self.c.upload("/models/2/items/import", "file", "items.csv", csv, {"mode": "replace"})
        self.assertIn("2건을 등록", flash(r))
        self.assertEqual(len(self.db.list_items(2)), 2)

    def test_22_html_is_escaped(self):
        self.c.post("/models/new", {"model_name": '<script>alert(1)</script>', "part_no": "XSS-1"})
        html = self.c.get("/models")["body"].decode("utf-8")
        self.assertNotIn("<script>alert(1)</script>", html)
        self.assertIn("&lt;script&gt;", html)

    def test_23_duplicate_model_rejected(self):
        r = self.c.post("/models/new", {"model_name": "TEST-CSV", "part_no": "P-001"})
        self.assertIn("이미 등록", flash(r))


if __name__ == "__main__":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    unittest.main(verbosity=2)
