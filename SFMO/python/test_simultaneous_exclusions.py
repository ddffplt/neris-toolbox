from contextlib import redirect_stdout
from copy import deepcopy
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import load_workbook

import IncidentPullSimultaneousCalls as report


class SimultaneousExclusionsTests(unittest.TestCase):
    def setUp(self):
        self.entities = [
            {"neris_id": "FD25025316", "name": "Boston Fire Department"},
            {"neris_id": "FD25017760", "name": "City Of Lowell Fire Department"},
            {"neris_id": "FD25017814", "name": "City Of Somerville Fire Deparment"},
            {"neris_id": "FD25009953", "name": "Lynnfield Fire Department"},
            {"neris_id": "FD25013306", "name": "West Springfield Fire Department"},
        ]
        self.client = Mock()
        self.client.list_entities.return_value = {"entities": self.entities, "total_count": len(self.entities)}

    def test_exact_list_contains_all_requested_departments(self):
        self.assertEqual(set(report.EXCLUDED_DEPARTMENTS.values()), {
            "Boston", "Brockton", "Cambridge", "Chicopee", "Fall River", "Framingham",
            "Lawrence", "Lowell", "Lynn", "New Bedford", "Newton", "Quincy",
            "Somerville", "Springfield", "Worcester", "Haverhill", "Malden", "Waltham",
        })
        self.assertEqual(len(report.EXCLUDED_DEPARTMENTS), 18)

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_statewide_excludes_ids_but_not_similar_names(self, fetch):
        report.get_scoped_incidents(self.client, "MA")
        self.assertEqual([call.args[1] for call in fetch.call_args_list],
                         ["FD25009953", "FD25013306"])

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_county_scope_applies_exclusion(self, fetch):
        report.get_scoped_incidents(self.client, "MA", entity_prefix="FD25017", entity_scope="fd25_prefix")
        fetch.assert_not_called()

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_explicit_id_is_excluded_without_a_fetch(self, fetch):
        for entity_id in report.EXCLUDED_DEPARTMENTS:
            with self.subTest(entity_id=entity_id):
                self.assertEqual(report.get_scoped_incidents(
                    self.client, "MA", entity_id=entity_id.lower(), entity_scope="entity"
                ), [])
        fetch.assert_not_called()
        self.client.list_entities.assert_not_called()

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_multiple_names_and_ids_apply_exclusions(self, fetch):
        report.get_scoped_incidents(
            self.client, "MA", entity_scope="multiple_entities",
            entity_names=["Boston", "FD25017814", "Lynnfield"],
        )
        self.assertEqual([call.args[1] for call in fetch.call_args_list], ["FD25009953"])

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_explicit_name_applies_exclusion(self, fetch):
        report.get_scoped_incidents(self.client, "MA", entity_id="Boston", entity_scope="entity")
        fetch.assert_not_called()

    @patch.object(report, "get_entity_incidents", return_value=[])
    def test_same_city_name_in_another_state_is_not_excluded(self, fetch):
        self.client.list_entities.return_value = {
            "entities": [{"neris_id": "FD17001001", "name": "Springfield Fire Department"}],
            "total_count": 1,
        }
        report.get_scoped_incidents(self.client, "IL")
        self.assertEqual(fetch.call_args.args[1], "FD17001001")


class SimultaneousExportTests(unittest.TestCase):
    def test_raw_data_fields_are_removed_by_name_not_excel_position(self):
        labels = ["retained_before", *sorted(report.RAW_DATA_EXCLUDED_FIELDS), "retained_after"]
        frame = pd.DataFrame([range(len(labels))], columns=labels)

        result = report.drop_raw_data_columns(frame)

        self.assertEqual(list(result.columns), ["retained_before", "retained_after"])

    def test_multiple_entities_get_one_summary_sheet_with_entity_sections(self):
        incidents = [
            {
                "neris_id": "INC1",
                "base": {
                    "department_neris_id": "FD1",
                    "incident_number": "1",
                    "people_present": 2,
                    "location_use": {
                        "use_type": "RESIDENTIAL",
                        "vacancy_cause": "UNKNOWN",
                        "location_modifiers": ["MULTIFAMILY"],
                    },
                },
                "dispatch": {"call_create": "2026-01-01T10:00:00Z"},
                "incident_types": [{"type": "FIRE", "primary": True}],
                "unit_responses": [{"reported_unit_id": "U1"}],
                "unmapped_field": [],
                "fire_detail": {"location_detail": {"acres_burned": 1.5}},
            },
            {
                "neris_id": "INC3",
                "base": {
                    "department_neris_id": "FD1",
                    "incident_number": "3",
                },
                "dispatch": {"call_create": "2025-12-15T10:00:00Z"},
                "incident_types": [{"type": "FIRE", "primary": True}],
            },
            {
                "neris_id": "INC2",
                "base": {
                    "department_neris_id": "FD2",
                    "incident_number": "2",
                    "location_use": {"use_type": "COMMERCIAL"},
                },
                "dispatch": {"call_create": "2026-01-02T10:00:00Z"},
                "incident_types": [{"type": "RESCUE", "primary": True}],
                "unmapped_field": [{"key": "value"}],
                "smoke_alarm": {"presence": {"alarm_types": ["SMOKE"]}},
                "fire_alarm": {"presence": {"alarm_types": ["FIRE"]}},
            },
        ]
        client = Mock()
        client.get_entity.side_effect = lambda entity_id: {
            "name": {"FD1": "Alpha Fire", "FD2": "Beta Fire"}[entity_id]
        }

        with TemporaryDirectory() as temp_dir:
            with patch("pathlib.Path.home", return_value=Path(temp_dir)):
                with redirect_stdout(StringIO()):
                    output_path = report.export_to_excel(
                        incidents,
                        client,
                        "MA",
                        entity_scope="fd25_prefix",
                        entity_prefix="FD25001",
                    )

            workbook = load_workbook(output_path, read_only=True, data_only=True)
            try:
                self.assertEqual(
                    workbook.sheetnames,
                    ["Overview", "Entity Summary", "Raw Data", "Report Sources"],
                )
                summary = workbook["Entity Summary"]
                rows = list(summary.iter_rows(values_only=True))
                self.assertIn(("FD1", "Alpha Fire", None, None, None, None, None), rows)
                self.assertIn(("December", 2025, 1, 0, None, None, None), rows)
                self.assertIn(("January", 2026, 1, 1, None, None, None), rows)
                self.assertIn(("FD2", "Beta Fire", None, None, None, None, None), rows)
                self.assertIn(("January", 2026, 1, 0, None, None, None), rows)
                fd1_heading = rows.index(("FD1", "Alpha Fire", None, None, None, None, None))
                fd1_months = [
                    row[:2] for row in rows[fd1_heading + 2:rows.index(
                        ("FD2", "Beta Fire", None, None, None, None, None)
                    )]
                    if row and row[0] in {"December", "January"}
                ]
                self.assertEqual(fd1_months, [("December", 2025), ("January", 2026)])
                self.assertNotIn("First Incident Type", [
                    cell for row in rows for cell in row if cell is not None
                ])
                raw_data = workbook["Raw Data"]
                raw_data.calculate_dimension(force=True)
                raw_headers = {
                    cell.value: cell.column for cell in raw_data[1]
                }
                self.assertIn("unmapped_field", raw_headers)
                incident_id_column = raw_headers["neris_id"]
                inc1_row = next(
                    row for row in range(2, raw_data.max_row + 1)
                    if raw_data.cell(row, incident_id_column).value == "INC1"
                )
                inc2_row = next(
                    row for row in range(2, raw_data.max_row + 1)
                    if raw_data.cell(row, incident_id_column).value == "INC2"
                )
                self.assertIsNone(raw_data.cell(inc1_row, raw_headers["unmapped_field"]).value)
                self.assertEqual(
                    raw_data.cell(inc2_row, raw_headers["unmapped_field"]).value,
                    '[{"key": "value"}]',
                )
                self.assertTrue({
                    "fire_detail_location_detail_acres_burned",
                    "base_location_use_use_type",
                    "base_location_use_vacancy_cause",
                    "base_location_use_location_modifiers",
                    "smoke_alarm_presence_alarm_types",
                    "fire_alarm_presence_alarm_types",
                    "smoke_alarm_presence_operation_alerted_failed_other_failure_reason",
                }.isdisjoint(raw_headers))
            finally:
                workbook.close()

    def test_flattening_matches_pandas_order_and_collisions(self):
        record = {
            "nested": {"value": 1, "inner": {"item": 2}},
            "nested_value": 3,
            "empty": {},
            "items": [{"a": 1}],
        }
        expected = pd.json_normalize(record, sep="_").iloc[0].to_dict()
        self.assertEqual(list(report.flatten_raw_record(record).items()), list(expected.items()))

    def test_streamed_export_keeps_late_fields_and_formulas(self):
        incidents = [
            {
                "neris_id": f"INC{i}",
                "base": {"department_neris_id": "FD1", "incident_number": str(i)},
                "dispatch": {"call_create": "2026-01-01T10:00:00Z"},
                "all_empty": [] if i % 2 else None,
                "late_field": {"answer": 0} if i == 54 else {},
                "weather": {"humidity": 75},
            }
            for i in range(55)
        ]
        original = deepcopy(incidents)
        client = Mock()
        client.get_entity.return_value = {"name": "Alpha Fire"}
        with TemporaryDirectory() as temp_dir:
            with patch("pathlib.Path.home", return_value=Path(temp_dir)):
                with redirect_stdout(StringIO()):
                    path = report.export_to_excel(incidents, client, "MA")
            self.assertEqual(incidents, original)
            workbook = load_workbook(path, data_only=False)
            try:
                self.assertEqual(workbook.sheetnames, ["Overview", "Raw Data", "Report Sources"])
                raw = workbook["Raw Data"]
                headers = [cell.value for cell in raw[1]]
                self.assertNotIn("all_empty", headers)
                self.assertNotIn("weather_humidity", headers)
                self.assertEqual(raw.cell(56, headers.index("late_field_answer") + 1).value, 0)
                overview = workbook["Overview"]
                overview_headers = [cell.value for cell in overview[1]]
                eastern = overview_headers.index("Call Date (Eastern Time Zone)") + 1
                self.assertTrue(overview.cell(2, eastern).value.startswith("=IF("))
                self.assertEqual(overview.cell(2, eastern).number_format, "mm/dd/yy hh:mm")
                self.assertEqual(overview.freeze_panes, "A2")
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
