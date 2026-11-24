import os
import tempfile
import unittest
from unittest.mock import Mock, patch

import pandas as pd
import requests
from openpyxl import load_workbook

import EntityFireActivityReport as fire
import FireCasualtyReport as casualty
import StateActivityReport as state
import StateActivityReport_w_IncTypeCount as state_types
import StateActivityReport_w_IncTypeCount_by_County as county


class ReportFetchIssuesTests(unittest.TestCase):
    def test_entity_fire_main_continues_after_failure(self):
        names = {"FD1": "Failed", "FD2": "Complete"}
        with patch.object(fire, "ensure_dependencies"), \
                patch.object(fire, "prompt_config", return_value=("MA", "statewide", None, None, [], "All Records", None, None)), \
                patch.object(fire, "authenticate"), \
                patch.object(fire, "fetch_entities", return_value=names), \
                patch.object(fire, "fetch_incidents", side_effect=[RuntimeError("HTTP 504"), []]), \
                patch.object(fire, "export_excel", return_value="report.xlsx") as export:
            fire.main()
            self.assertEqual(export.call_args.args[-1], {"FD1": "HTTP 504"})

    def test_state_fetches_reject_http_errors(self):
        response = requests.Response()
        response.status_code = 403
        response._content = b'{"detail": "forbidden"}'
        for module in (state, state_types, county):
            with self.subTest(module=module.__name__):
                client = Mock()
                client.list_incidents.return_value = response
                with self.assertRaises(requests.HTTPError):
                    module.fetch_incidents_for_entity(client, "MA", "FD1")
                client._session.get.return_value = response
                with self.assertRaises(requests.HTTPError):
                    module.fetch_nars_for_entity(client, "MA", "FD1")

    def test_fire_and_casualty_reports_mark_failures(self):
        names = {"FD1": "Failed", "FD2": "Complete"}
        errors = {"FD1": "HTTP 504"}
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(os.path, "expanduser", return_value=directory):
            paths = [
                fire.export_excel({}, {}, {}, {}, names, "MA", fetch_errors=errors),
                casualty.export_excel({}, {}, names, "state", "MA", "period", "period", errors),
            ]
            for path, summary_name, header_row in (
                    (paths[0], "Summary", 1), (paths[1], "Entity Summary", 2)):
                with self.subTest(path=path):
                    workbook = load_workbook(path)
                    sheet = workbook[summary_name]
                    failed = sheet[header_row + 1]
                    success = sheet[header_row + 2]
                    self.assertEqual(failed[0].value, "FD1")
                    self.assertEqual(failed[0].fill.fgColor.rgb, "00FFC7CE")
                    self.assertIsNone(failed[2].value)
                    self.assertEqual(success[2].value, 0)
                    self.assertNotEqual(success[0].fill.fgColor.rgb, "00FFC7CE")
                    self.assertEqual(workbook["Fetch Issues"]["A2"].value, "FD1")
                    self.assertIn("504", workbook["Fetch Issues"]["C2"].value)
                    workbook.close()

    def test_state_and_county_workbooks_identify_failed_entities(self):
        with tempfile.TemporaryDirectory() as directory:
            previous = os.getcwd()
            try:
                os.chdir(directory)
                for module, is_county in ((state, False), (state_types, False), (county, True)):
                    with self.subTest(module=module.__name__):
                        eid = "FD25001" if is_county else "FD2500101"
                        names = {"FD2500101": "Failed"}
                        counts = {
                            eid: {"name": "Name", "types": {"FIRE": 2},
                                  "total": 2, "multiple_types": 0}
                        }
                        columns = (["County NERIS ID", "County Name"] if is_county else
                                   ["NERIS Entity ID", "Department Name"]) + ["Jan-2026"]
                        pivot = pd.DataFrame([[eid, "Name", 2], ["TOTAL", "", 2]], columns=columns)
                        path = module.export_to_excel(
                            pivot, ["Jan-2026"], "MA", {}, {}, counts,
                            {"FD2500101": "HTTP 504"}, names,
                        )
                        workbook = load_workbook(path)
                        for sheet_name in ("Incident Type Summary", "Monthly Incident Counts"):
                            sheet = workbook[sheet_name]
                            self.assertEqual(sheet["A2"].fill.fgColor.rgb, "00FFC7CE")
                            self.assertEqual(sheet["C2"].value, 2 if is_county else None)
                            self.assertIn("Partial" if is_county else "unavailable",
                                          sheet.cell(2, sheet.max_column).value)
                            self.assertIn("Incomplete", sheet.cell(3, sheet.max_column).value)
                        self.assertEqual(workbook["Fetch Issues"]["B2"].value, "Failed")
                        workbook.close()
            finally:
                os.chdir(previous)


if __name__ == "__main__":
    unittest.main()
