"""Offline routing checks; no credentials or Google writes are required."""
import ast
import importlib
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import project_sheets as config

ROOT = Path(__file__).resolve().parents[1]


def load_function(filename, name, namespace):
    """Load production function without running unrelated script startup code."""
    tree = ast.parse((ROOT / filename).read_text())
    function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == name)
    exec(compile(ast.Module(body=[function], type_ignores=[]), filename, 'exec'), namespace)
    return namespace[name]


class SheetRoutingTests(unittest.TestCase):
    def test_defaults_and_163_migration(self):
        with patch.dict(os.environ, {'IRIS_PP_ID_163_22': ''}):
            importlib.reload(config)
            self.assertEqual(config.PP_ID['65.5'], config.RFI_ID['65.5'])
            self.assertNotEqual(config.PP_ID['163.22'], config.RFI_ID['163.22'])
        try:
            with patch.dict(os.environ, {'IRIS_PP_ID_163_22': config.RFI_ID['163.22']}):
                importlib.reload(config)
                self.assertEqual(config.PP_ID['163.22'], config.RFI_ID['163.22'])
                self.assertEqual(config.PP_ID['65.5'], config.RFI_ID['65.5'])
        finally:
            importlib.reload(config)

    def test_explicit_id_overrides_stale_cache_and_drive_search(self):
        lookup = load_function('update_rfi_google_sheets.py', 'get_spreadsheet_id',
                               {'RFI_ID_BY_NAME': config.RFI_ID_BY_NAME, 'os': os})
        drive = Mock()
        with tempfile.NamedTemporaryFile(mode='w') as cache:
            cache.write('obsolete-workbook-id')
            cache.flush()
            self.assertEqual(lookup(drive, config.RFI_NAMES['65.5'], cache.name),
                             config.RFI_ID['65.5'])
        drive.files.assert_not_called()

    def test_status_lookup_uses_same_workbook(self):
        lookup = load_function('push_rfi_status_from_pdfs.py', '_find_spreadsheet_id',
                               {'RFI_ID_BY_NAME': config.RFI_ID_BY_NAME})
        drive = Mock()
        self.assertEqual(lookup(drive, config.RFI_NAMES['65.5']), config.RFI_ID['65.5'])
        drive.files.assert_not_called()

    def test_rfi_maintenance_excludes_progress_tables(self):
        lookup = load_function('update_rfi_google_sheets.py', 'get_all_named_tables', {
            'PP_SHEET_NAME': config.PP_SHEET_NAME, 'DATA_COLUMNS': ['RFI No.'],
            'col_letter': lambda i: chr(65+i),
        })
        service = Mock()
        service.spreadsheets().get().execute.return_value = {'sheets': [
            {'properties': {'title': '65.5 UNGUJA RFI_DATA', 'sheetId': 1},
             'tables': [{'name': 'RFI_TABLE', 'range': {'endColumnIndex': 2}}]},
            {'properties': {'title': config.PP_SHEET_NAME, 'sheetId': 2133406991},
             'tables': [{'name': 'Progress', 'range': {'endColumnIndex': 2}}]},
        ]}
        self.assertEqual(set(lookup(service, config.RFI_ID['65.5'])), {'RFI_TABLE'})

    def test_progress_reader_targets_exact_tab_in_rfi_workbook(self):
        read = load_function('match_road_codes.py', 'read_from_physical_progress', {
            'physical_progress_tab': config.physical_progress_tab, 're': __import__('re'),
        })
        service = Mock()
        service.spreadsheets().get().execute.return_value = {'sheets': [
            {'properties': {'title': '65.5 UNGUJA RFI_DATA'}},
            {'properties': {'title': config.PP_SHEET_NAME}},
        ]}
        service.spreadsheets().values().get().execute.return_value = {'values': []}
        read(service, config.PP_ID['65.5'], '65.5')
        service.spreadsheets().values().get.assert_called_with(
            spreadsheetId=config.RFI_ID['65.5'], range="'PHYSICAL PROGRESS'!A:F")
        service.spreadsheets().values().get.reset_mock()
        service.spreadsheets().get().execute.return_value = {'sheets': [
            {'properties': {'title': '65.5 UNGUJA RFI_DATA'}},
        ]}
        read(service, config.PP_ID['65.5'], '65.5')
        service.spreadsheets().values().get.assert_not_called()

    def test_missing_progress_tab_fails_closed(self):
        with self.assertRaisesRegex(ValueError, 'PHYSICAL PROGRESS'):
            config.physical_progress_tab(['65.5 UNGUJA RFI_DATA'])


if __name__ == '__main__':
    unittest.main()
