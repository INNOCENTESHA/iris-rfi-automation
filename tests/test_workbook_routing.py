import importlib.util
import os
from pathlib import Path
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
PROJECT_SHEETS = ROOT / "project_sheets.py"


def load_project_sheets(env=None):
    spec = importlib.util.spec_from_file_location("project_sheets_under_test", PROJECT_SHEETS)
    module = importlib.util.module_from_spec(spec)
    with patch.dict(os.environ, env or {}, clear=True):
        spec.loader.exec_module(module)
    return module


class WorkbookRoutingTests(unittest.TestCase):
    def test_65_5_physical_progress_is_in_rfi_workbook(self):
        routes = load_project_sheets()
        self.assertEqual(
            routes.RFI_ID["65.5"],
            "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU",
        )
        self.assertEqual(routes.PP_ID["65.5"], routes.RFI_ID["65.5"])

    def test_163_22_stays_on_standalone_workbook_by_default(self):
        routes = load_project_sheets()
        self.assertEqual(
            routes.PP_ID["163.22"],
            "1YbDCzePl76CqtO9dLl2YZjIZtVmJKSgVkmU6LKqXKUc",
        )
        self.assertNotEqual(routes.PP_ID["163.22"], routes.RFI_ID["163.22"])

    def test_163_22_route_can_be_overridden_after_migration(self):
        destination = "future-163-22-combined-workbook"
        routes = load_project_sheets({"IRIS_PP_ID_163_22": destination})
        self.assertEqual(routes.PP_ID["163.22"], destination)

    def test_physical_progress_tab_must_match_exactly(self):
        routes = load_project_sheets()
        self.assertEqual(
            routes.physical_progress_tab(["RFI_DATA", "PHYSICAL PROGRESS"]),
            "PHYSICAL PROGRESS",
        )
        with self.assertRaises(ValueError):
            routes.physical_progress_tab(["RFI_DATA", "Physical Progress"])

    def test_retired_65_5_workbook_is_absent_from_python_scripts(self):
        retired_id = "16RE08Ivv2nSltqxVk1X9G8c" + "BYcZP82kIGhzGAwRtZ7s"
        offenders = [
            path.name
            for path in ROOT.glob("*.py")
            if retired_id in path.read_text(encoding="utf-8")
        ]
        self.assertEqual(offenders, [])

    def test_scripts_use_shared_explicit_routes(self):
        matcher = (ROOT / "match_road_codes.py").read_text(encoding="utf-8")
        importer = (ROOT / "import_rfi_from_pdf.py").read_text(encoding="utf-8")
        status_push = (ROOT / "push_rfi_status_from_pdfs.py").read_text(encoding="utf-8")
        rfi_update = (ROOT / "update_rfi_google_sheets.py").read_text(encoding="utf-8")

        self.assertIn("from project_sheets import PP_ID, RFI_ID, physical_progress_tab", matcher)
        self.assertIn("pp_tab = physical_progress_tab(tabs)", matcher)
        self.assertIn("ss = _api(gc.open_by_key, RFI_ID[km_key])", importer)
        self.assertIn("from project_sheets import RFI_ID_BY_NAME", status_push)
        self.assertLess(
            status_push.index("if name in RFI_ID_BY_NAME:"),
            status_push.index("drive.files().list"),
        )
        self.assertIn("candidates = [physical_progress_tab(sheet_names)]", rfi_update)
        self.assertIn("if sheet_title == PP_SHEET_NAME:", rfi_update)
        self.assertLess(
            rfi_update.index("if name in RFI_ID_BY_NAME:"),
            rfi_update.index("if os.path.exists(id_file):"),
        )


if __name__ == "__main__":
    unittest.main()
