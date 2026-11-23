"""Offline tests for rfi_review_agent.py (no network, no API key)."""
import io
import os
import sys
import tempfile
import types
import unittest
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import rfi_review_agent as agent  # noqa: E402


def make_docx(text: str) -> bytes:
    import docx
    d = docx.Document()
    for line in text.splitlines():
        d.add_paragraph(line)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


class ClassifyTests(unittest.TestCase):
    def test_kizimkazi_by_archive(self):
        self.assertEqual(agent.classify_project(
            "KIZIMKAZI ROAD NETWORK RFI's FOR SITE 21.11.2026.zip", "RFI-001.pdf", ""), "kizimkazi")

    def test_kizimkazi_by_text_beats_km(self):
        txt = "DESIGN AND BUILD OF THE KIZIMKAZI ROAD NETWORK ... THE 113.50 KM ROADS"
        self.assertEqual(agent.classify_project("RFIs.zip", "a.pdf", txt), "kizimkazi")

    def test_kzp_prefix(self):
        self.assertEqual(agent.classify_project("x.zip", "KZP-IRIS-QAQC-RFI-0014.pdf", ""), "kizimkazi")

    def test_road_code_series(self):
        self.assertEqual(agent.classify_project("x.zip", "WA-104 subbase.pdf", ""), "65.5")
        self.assertEqual(agent.classify_project("x.zip", "NA-253 subgrade.pdf", ""), "163.22")

    def test_rfi_prefix_and_text(self):
        self.assertEqual(agent.classify_project("x.zip", "IRIS-QAQC-0012.pdf", ""), "163.22")
        self.assertEqual(agent.classify_project("x.zip", "IR-QAQC-0012.pdf", ""), "65.5")
        self.assertEqual(agent.classify_project("x.zip", "a.pdf", "THE 65.5 KM ROADS"), "65.5")
        self.assertEqual(agent.classify_project("x.zip", "a.pdf", "nothing here"), "")


class ParseTests(unittest.TestCase):
    def test_parse_and_row(self):
        reply = 'Here:\n{"rfi_no": "IRIS-Q-RFI-U-26-0101", "road_code": "SO-301", ' \
                '"result": "REJECTED", "flag": true, "severity": "high", ' \
                '"chainage": "0+500 - 0+750", "issue": "MDD 92% below 95%"}'
        r = agent.parse_review(reply)
        row = agent.review_to_row(r, {"archive": "a.zip", "inner": "b.pdf"}, "=DATE(2026,11,23)")
        self.assertEqual(len(row), len(agent.REVIEW_HEADER))
        self.assertEqual(row[0], "=DATE(2026,11,23)")
        self.assertEqual(row[1], "'IRIS-Q-RFI-U-26-0101")
        self.assertEqual(row[11], "'YES")
        self.assertEqual(row[12], "'HIGH")
        self.assertEqual(row[15], "'a.zip > b.pdf")
        self.assertEqual(row[2], "")  # missing value stays empty

    def test_bad_reply(self):
        self.assertEqual(agent.parse_review("sorry")["severity"], "LOW")

    def test_cell_date(self):
        self.assertEqual(str(agent.cell_date(46349)), "2026-11-23")
        self.assertEqual(str(agent.cell_date("23/11/2026")), "2026-11-23")
        self.assertIsNone(agent.cell_date("Reviewed On"))


class FolderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name
        today = datetime.now().strftime("%d.%m.%Y")
        old = (datetime.now() - timedelta(days=60)).strftime("%d.%m.%Y")
        body = ("REQUEST FOR INSPECTION\nRFI No: IRIS-Q-RFI-U-26-0101\nRoad: SO-301 Kizimkazi Mkunguni\n"
                "Chainage 0+500 - 0+750\nActivity: Subbase compaction\nResult: Field density 92% MDD, "
                "required 95% MDD.\n" * 3)
        with zipfile.ZipFile(os.path.join(self.dir, f"KIZIMKAZI ROAD NETWORK RFI's FOR LAB {today}.zip"), "w") as z:
            z.writestr("RFI-0101.docx", make_docx(body))
            inner = io.BytesIO()
            with zipfile.ZipFile(inner, "w") as z2:
                z2.writestr("WA-104 RFI-0007.docx", make_docx("Gando Wete Konde\nTack coat\n" * 30))
            z.writestr("nested.zip", inner.getvalue())
        with zipfile.ZipFile(os.path.join(self.dir, f"163.22 RFIs {old}.zip"), "w") as z:
            z.writestr("NA-253.docx", make_docx("old RFI\n" * 50))

    def tearDown(self):
        self.tmp.cleanup()

    def test_collect_and_filter(self):
        docs = agent.collect_documents(self.dir)
        self.assertEqual(len(docs), 3)
        since = (datetime.now() - timedelta(days=14)).date()
        kept = [d for d in docs if agent.include_by_date(d, since)]
        self.assertEqual(len(kept), 2)  # the 60-day-old archive is left out

    def test_request_shape(self):
        docs = agent.collect_documents(self.dir)
        d = next(x for x in docs if x["inner"] == "RFI-0101.docx")
        req = agent.build_request(d, "kizimkazi", "claude-sonnet-5-5")
        self.assertTrue(req["custom_id"].startswith("rfi-"))
        self.assertLessEqual(len(req["custom_id"]), 64)
        self.assertIn("Kizimkazi", req["params"]["messages"][0]["content"][0]["text"])

    def test_scanned_pdf_sent_as_document(self):
        d = {"archive": "", "inner": "scan.pdf", "data": b"%PDF-1.4 fake", "text": "", "sha": "0" * 64}
        req = agent.build_request(d, "65.5", "m")
        self.assertEqual(req["params"]["messages"][0]["content"][0]["type"], "document")


class FakeWS:
    def __init__(self):
        self.rows = [list(agent.REVIEW_HEADER)]

    def append_rows(self, rows, **kw):
        self.rows.extend(rows)


class CollectBatchTests(unittest.TestCase):
    def test_collect_writes_and_marks(self):
        sha_ok, sha_bad = "a" * 64, "b" * 64
        metas = {
            agent.custom_id_for(sha_ok): {"sha": sha_ok, "project": "kizimkazi", "archive": "z", "inner": "1.pdf"},
            agent.custom_id_for(sha_bad): {"sha": sha_bad, "project": "65.5", "archive": "z", "inner": "2.pdf"},
        }
        msg = types.SimpleNamespace(content=[types.SimpleNamespace(
            type="text", text='{"rfi_no":"R1","severity":"NONE","flag":false}')])
        results = [
            types.SimpleNamespace(custom_id=agent.custom_id_for(sha_ok),
                                  result=types.SimpleNamespace(type="succeeded", message=msg)),
            types.SimpleNamespace(custom_id=agent.custom_id_for(sha_bad),
                                  result=types.SimpleNamespace(type="errored")),
        ]
        batches = types.SimpleNamespace(
            retrieve=lambda bid: types.SimpleNamespace(processing_status="ended"),
            results=lambda bid: iter(results))
        client = types.SimpleNamespace(messages=types.SimpleNamespace(batches=batches))
        ws = FakeWS()
        orig = agent.get_tab
        agent.get_tab = lambda gc, sid, title, header: ws
        try:
            state = {"reviewed": {}, "pending_batches": {}}
            done = agent.collect_batch(client, "msgbatch_1", metas, state, None, False)
        finally:
            agent.get_tab = orig
        self.assertTrue(done)
        self.assertIn(sha_ok, state["reviewed"])
        self.assertNotIn(sha_bad, state["reviewed"])  # retried next run
        self.assertEqual(len(ws.rows), 2)
        self.assertEqual(ws.rows[1][1], "'R1")


if __name__ == "__main__":
    unittest.main()
