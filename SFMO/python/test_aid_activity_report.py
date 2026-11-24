import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, call, patch

import requests
from openpyxl import load_workbook

import AidActivityReport as report


def http_response(status, body):
    response = requests.Response()
    response.status_code = status
    response._content = body.encode("utf-8")
    return response


class FetchEntitiesTests(unittest.TestCase):
    @patch.object(report.time, "sleep")
    def test_temporary_errors_retry_same_page(self, sleep):
        client = Mock()
        first_page = [{"neris_id": f"FD{i}", "name": f"Department {i}"} for i in range(100)]
        client.list_entities.side_effect = [
            {"entities": first_page, "total_count": 101},
            http_response(503, "<html>Unavailable</html>"),
            requests.HTTPError(response=http_response(504, "")),
            {"entities": [{"neris_id": "FD100", "name": "Last"}], "total_count": 101},
        ]
        entities = report.fetch_entities(client, "MA", None)
        self.assertEqual(len(entities), 101)
        self.assertEqual(entities["FD100"], "Last")
        self.assertEqual(client.list_entities.call_args_list, [
            call(state="MA", page_size=100, page_number=1),
            call(state="MA", page_size=100, page_number=2),
            call(state="MA", page_size=100, page_number=2),
            call(state="MA", page_size=100, page_number=2),
        ])
        self.assertEqual(sleep.call_args_list, [call(5), call(10)])

    @patch.object(report.time, "sleep")
    def test_503_exhaustion_reports_status_without_parsing_html(self, sleep):
        client = Mock()
        client.list_entities.return_value = http_response(503, "<html>Unavailable</html>")
        with self.assertRaises(report.EntityFetchError) as caught:
            report.fetch_entities(client, "MA", None)
        self.assertIn("HTTP 503", str(caught.exception))
        self.assertIn("after 5 retries", str(caught.exception))
        self.assertEqual(client.list_entities.call_count, 6)
        self.assertEqual(sleep.call_args_list, [call(5), call(10), call(20), call(40), call(80)])

    def test_other_http_errors_fail_without_retries(self):
        for raised in (False, True):
            with self.subTest(raised=raised):
                client = Mock()
                response = http_response(403, "")
                if raised:
                    client.list_entities.side_effect = requests.HTTPError(response=response)
                else:
                    client.list_entities.return_value = response
                with self.assertRaisesRegex(report.EntityFetchError, "HTTP 403"):
                    report.fetch_entities(client, "MA", None)
                client.list_entities.assert_called_once()

    def test_connection_error_fails_clearly(self):
        client = Mock()
        client.list_entities.side_effect = requests.ConnectionError("connection failed")
        with self.assertRaisesRegex(report.EntityFetchError, "MA.*page 1.*connection failed"):
            report.fetch_entities(client, "MA", None)

    def test_invalid_successful_responses_do_not_produce_empty_results(self):
        for response in (http_response(200, "<html>Unavailable</html>"), {"error": "unexpected"}):
            with self.subTest(response=response):
                client = Mock()
                client.list_entities.return_value = response
                with self.assertRaises(report.EntityFetchError):
                    report.fetch_entities(client, "MA", None)

    def test_successful_response_object(self):
        client = Mock()
        client.list_entities.return_value = http_response(
            200, '{"entities": [{"neris_id": "FD1", "name": "Department"}], "total_count": 1}'
        )
        self.assertEqual(report.fetch_entities(client, "MA", None), {"FD1": "Department"})


class FetchIncidentsTests(unittest.TestCase):
    def test_paginated_dictionary_responses(self):
        client = Mock()
        client.list_incidents.side_effect = [
            {"incidents": [{"id": 1}], "next_cursor": "next"},
            {"incidents": [{"id": 2}], "next_cursor": None},
        ]
        result = report.fetch_incidents(client, "MA", "FD1", page_size=5)
        self.assertEqual(result, [{"id": 1}, {"id": 2}])
        self.assertEqual(client.list_incidents.call_args_list, [
            call(neris_id_entity="FD1", page_size=5),
            call(neris_id_entity="FD1", page_size=5, cursor="next"),
        ])

    def test_non_json_http_error_reports_status(self):
        client = Mock()
        client.list_incidents.return_value = http_response(500, "<html>Server error</html>")
        with self.assertRaises(report.IncidentFetchError) as caught:
            report.fetch_incidents(client, "MA", "FD1")
        self.assertEqual(caught.exception.status_code, 500)
        self.assertIn("HTTP 500", str(caught.exception))
        self.assertIn("FD1", str(caught.exception))

    @patch.object(report.time, "sleep")
    def test_gateway_timeout_retries_the_same_cursor_without_duplicate_incidents(self, sleep):
        client = Mock()
        client.list_incidents.side_effect = [
            {"incidents": [{"id": 1}], "next_cursor": "next"},
            http_response(504, "Gateway timeout"),
            requests.HTTPError(response=http_response(504, "")),
            {"incidents": [{"id": 2}]},
        ]
        self.assertEqual(report.fetch_incidents(client, "MA", "FD1"), [{"id": 1}, {"id": 2}])
        self.assertEqual(client.list_incidents.call_args_list, [
            call(neris_id_entity="FD1", page_size=100),
            call(neris_id_entity="FD1", page_size=100, cursor="next"),
            call(neris_id_entity="FD1", page_size=100, cursor="next"),
            call(neris_id_entity="FD1", page_size=100, cursor="next"),
        ])
        self.assertEqual(sleep.call_args_list, [call(5), call(10)])

    @patch.object(report.time, "sleep")
    def test_gateway_timeout_fails_after_five_retries(self, sleep):
        for raised in (False, True):
            with self.subTest(raised=raised):
                client = Mock()
                response = http_response(504, "")
                client.list_incidents.side_effect = (
                    requests.HTTPError(response=response) if raised else None
                )
                client.list_incidents.return_value = response
                sleep.reset_mock()
                with self.assertRaises(report.IncidentFetchError) as caught:
                    report.fetch_incidents(client, "MA", "FD1")
                self.assertEqual(caught.exception.status_code, 504)
                self.assertIn("after 5 retries", str(caught.exception))
                self.assertEqual(client.list_incidents.call_count, 6)
                self.assertEqual(sleep.call_args_list, [call(5), call(10), call(20), call(40), call(80)])

    @patch.object(report.time, "sleep")
    def test_retry_budget_resets_for_each_page(self, sleep):
        client = Mock()
        client.list_incidents.side_effect = [
            http_response(504, ""),
            {"incidents": [{"id": 1}], "next_cursor": "next"},
            http_response(504, ""),
            {"incidents": [{"id": 2}]},
        ]
        self.assertEqual(report.fetch_incidents(client, "MA", "FD1"), [{"id": 1}, {"id": 2}])
        self.assertEqual(sleep.call_args_list, [call(5), call(5)])

    def test_successful_response_without_json_reports_status(self):
        client = Mock()
        client.list_incidents.return_value = http_response(200, "")
        with self.assertRaisesRegex(report.IncidentFetchError, "not valid JSON.*HTTP 200"):
            report.fetch_incidents(client, "MA", "FD1")

    def test_successful_response_object(self):
        client = Mock()
        client.list_incidents.return_value = http_response(200, '{"incidents": []}')
        self.assertEqual(report.fetch_incidents(client, "MA", "FD1"), [])

    def test_client_returned_text_reports_unavailable_status(self):
        client = Mock()
        client.list_incidents.return_value = "<html>Not JSON</html>"
        with self.assertRaisesRegex(report.IncidentFetchError, "not valid JSON.*status unavailable"):
            report.fetch_incidents(client, "MA", "FD1")

    def test_invalid_response_shape_is_not_treated_as_empty(self):
        client = Mock()
        client.list_incidents.return_value = {"error": "unexpected"}
        with self.assertRaisesRegex(report.IncidentFetchError, "expected an object containing an incidents list"):
            report.fetch_incidents(client, "MA", "FD1")

    def test_authorization_failure_on_cursor_retries_without_filters(self):
        client = Mock()
        client.list_incidents.side_effect = [
            {"incidents": [{"id": 1}], "next_cursor": "next"},
            http_response(403, ""),
            {"incidents": [{"id": 2}]},
        ]
        self.assertEqual(report.fetch_incidents(client, "MA", "FD1"), [{"id": 1}, {"id": 2}])
        self.assertEqual(client.list_incidents.call_args_list[-1], call(cursor="next", page_size=100))

    def test_raised_http_error_retains_status(self):
        client = Mock()
        client.list_incidents.side_effect = requests.HTTPError(response=http_response(403, ""))
        with self.assertRaises(report.IncidentFetchError) as caught:
            report.fetch_incidents(client, "MA", "FD1")
        self.assertEqual(caught.exception.status_code, 403)

    def test_non_http_error_is_not_mistaken_for_403(self):
        client = Mock()
        client.list_incidents.side_effect = ValueError("unrelated HTTP 403 text")
        with self.assertRaises(report.IncidentFetchError) as caught:
            report.fetch_incidents(client, "MA", "FD1")
        self.assertIsNone(caught.exception.status_code)


class AidInLieuTests(unittest.TestCase):
    def test_only_selected_recipients_are_grouped_with_the_reporting_department(self):
        client = Mock()
        names = {"FD25000001": "Recipient"}
        incidents = [
            {
                "neris_id": "incident-1",
                "base": {"department_neris_id": "FD25000002", "incident_number": "123"},
                "aids": [
                    {"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                     "department_neris_id": "FD25000001"},
                    {"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                     "department_neris_id": "FD26000001"},
                    {"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                     "department_neris_id": "FD25000001"},
                ],
            },
            {
                "neris_id": "incident-2",
                "base": {"department_neris_id": "FD25000001"},
                "aids": [
                    {"aid_direction": "RECEIVED", "aid_type": "IN_LIEU_AID",
                     "department_neris_id": "FD25000002"},
                ],
            },
        ]
        client.get_entity.return_value = {"name": "Giving Department"}
        result = report.collect_aid_in_lieu(client, incidents, names, dict(names))
        self.assertEqual(list(result), ["FD25000001"])
        self.assertEqual(len(result["FD25000001"]), 1)
        self.assertEqual(result["FD25000001"][0]["Incident NERIS ID"], "incident-1")
        self.assertEqual(result["FD25000001"][0]["Aid-Giving Department ID"], "FD25000002")
        self.assertEqual(result["FD25000001"][0]["Aid-Giving Department Name"], "Giving Department")
        self.assertEqual(result["FD25000001"][0]["Incident Number"], "123")


class MainTests(unittest.TestCase):
    def test_entity_list_failure_stops_before_creating_report(self):
        config = ("MA", "statewide", None, None, [], "All Records", None, None)
        with patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=Mock()), \
                patch.object(report, "fetch_entities", side_effect=report.EntityFetchError("HTTP 503")), \
                patch.object(report, "export_excel") as export:
            with self.assertRaisesRegex(SystemExit, "HTTP 503"):
                report.main()
            export.assert_not_called()

    def test_selected_entity_aggregates_all_aid_categories_without_fetching_other_entities(self):
        selected = "FD25000001"
        other = "FD25000002"
        config = ("MA", "entity", selected, None, [], "All Records", None, None)
        entities = {selected: "Selected Department", other: "Other Department"}
        incident = {
            "base": {"department_neris_id": selected},
            "incident_types": [{"type": "FIRE"}],
            "aids": [
                {"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID", "department_neris_id": other},
                {"aid_direction": "GIVEN", "aid_type": "SUPPORT_AID"},
                {"aid_direction": "RECEIVED", "aid_type": "ACTING_AS_AID"},
            ],
        }
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=Mock()), \
                patch.object(report, "fetch_entities", return_value=entities), \
                patch.object(report, "fetch_incidents", return_value=[incident]) as fetch, \
                patch.object(report, "collect_aid_in_lieu") as collect, \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            report.main()
            fetch.assert_called_once_with(unittest.mock.ANY, "MA", selected, None, None)
            collect.assert_not_called()
            workbook = load_workbook(next((Path(temp_dir) / "Documents").glob("*.xlsx")))
            self.addCleanup(workbook.close)
            self.assertEqual(list(workbook["Summary"].values)[1][:10],
                             (selected, "Selected Department", 1, 1, 0, 2, 0, 0, 1, 1))
            detail_rows = list(workbook["Entity Aid Reports"].values)
            self.assertEqual(detail_rows[2][:9], ("FIRE", 1, 1, 0, 2, 0, 0, 1, 1))
            self.assertEqual(detail_rows[3][:9], ("Total", 1, 1, 0, 2, 0, 0, 1, 1))
            self.assertIn("temporarily disabled", workbook["Aid Given In Lieu Of"]["A1"].value)
            self.assertTrue(any("temporarily disabled" in str(row[1])
                                for row in workbook["Report Sources"].values))

    def test_individual_report_includes_aid_from_other_state_departments(self):
        recipient = "FD25000001"
        giver = "FD25000002"
        config = ("MA", "entity", recipient, None, [], "All Records", None, None)
        entities = {recipient: "Recipient", giver: "Giving Department"}
        incident = {
            "neris_id": "incident-1",
            "base": {"department_neris_id": giver, "incident_number": "123"},
            "aids": [{"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                      "department_neris_id": recipient}],
        }
        client = Mock()
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report, "INCLUDE_AID_IN_LIEU_DETAILS", True), \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=client), \
                patch.object(report, "fetch_entities", return_value=entities), \
                patch.object(report, "fetch_incidents", side_effect=[[], [incident]]) as fetch, \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            report.main()
            workbook = load_workbook(next((Path(temp_dir) / "Documents").glob("*.xlsx")))
            self.addCleanup(workbook.close)
            rows = list(workbook["Aid Given In Lieu Of"].values)
            self.assertEqual(rows[0][:2], (recipient, "Recipient"))
            self.assertEqual(rows[2][:4], ("incident-1", giver, "Giving Department", "123"))
            self.assertEqual(list(workbook["Summary"].values)[1][:2], (recipient, "Recipient"))
            self.assertEqual(fetch.call_count, 2)

    def test_county_scope_uses_statewide_givers_and_records_their_fetch_failures(self):
        recipient = "FD25025001"
        giver = "FD25027001"
        failed = "FD25027002"
        entities = {recipient: "Recipient", giver: "Giving Department", failed: "Failed Department"}
        incident = {
            "neris_id": "incident-1",
            "base": {"department_neris_id": giver},
            "aids": [{"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                      "department_neris_id": recipient}],
        }
        config = ("MA", "fd25_prefix", None, "FD25025", [], "All Records", None, None)
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report, "INCLUDE_AID_IN_LIEU_DETAILS", True), \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=Mock()), \
                patch.object(report, "fetch_entities", return_value=entities), \
                patch.object(report, "fetch_incidents",
                             side_effect=[[], [incident], report.IncidentFetchError("unavailable")]), \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            report.main()
            workbook = load_workbook(next((Path(temp_dir) / "Documents").glob("*.xlsx")))
            self.addCleanup(workbook.close)
            rows = list(workbook["Aid Given In Lieu Of"].values)
            self.assertIn("Incomplete", rows[0][0])
            self.assertEqual(rows[2][:2], (recipient, "Recipient"))
            self.assertEqual(rows[4][:3], ("incident-1", giver, "Giving Department"))
            self.assertTrue(any("out-of-state" in str(row[0]) for row in rows))
            self.assertEqual(list(workbook["Summary"].values)[1][:2], (recipient, "Recipient"))
            self.assertEqual(workbook["Summary"].max_row, 2)
            issue = list(workbook["Fetch Issues"].values)[1]
            self.assertEqual(issue[:3], (failed, "Failed Department", "unavailable"))

    def test_massachusetts_report_excludes_non_massachusetts_recipient_ids(self):
        giver = "FD25000001"
        other = "FD26000001"
        entities = {giver: "Giving Department", other: "Other Department"}
        incident = {
            "neris_id": "incident-1",
            "base": {"department_neris_id": giver},
            "aids": [{"aid_direction": "GIVEN", "aid_type": "IN_LIEU_AID",
                      "department_neris_id": other}],
        }
        config = ("MA", "statewide", None, None, [], "All Records", None, None)
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report, "INCLUDE_AID_IN_LIEU_DETAILS", True), \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=Mock()), \
                patch.object(report, "fetch_entities", return_value=entities), \
                patch.object(report, "fetch_incidents", side_effect=[[incident], []]), \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            report.main()
            workbook = load_workbook(next((Path(temp_dir) / "Documents").glob("*.xlsx")))
            self.addCleanup(workbook.close)
            rows = list(workbook["Aid Given In Lieu Of"].values)
            self.assertTrue(rows[0][0].startswith("No aid-given-in-lieu-of incidents"))
            self.assertFalse(any(row[0] == other for row in rows))

    @patch.object(report.time, "sleep")
    def test_failed_entities_are_recorded_and_later_entities_are_processed(self, sleep):
        config = ("MA", "statewide", None, None, [], "All Records", None, None)
        entities = {"FD1": "Timeout Department", "FD2": "Healthy Department", "FD3": "Forbidden Department"}
        client = Mock()
        partial = {
            "base": {"department_neris_id": "FD1"},
            "aids": [{"aid_direction": "GIVEN", "aid_type": "SUPPORT_AID"}],
        }
        complete = {
            "base": {"department_neris_id": "FD2"},
            "aids": [{"aid_direction": "GIVEN", "aid_type": "SUPPORT_AID"}],
        }
        client.list_incidents.side_effect = (
            [{"incidents": [partial], "next_cursor": "next"}]
            + [http_response(504, "") for _ in range(6)]
            + [{"incidents": [complete]}, http_response(403, "")]
        )
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "prompt_config", return_value=config), \
                patch.object(report, "authenticate", return_value=client), \
                patch.object(report, "fetch_entities", return_value=entities), \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            report.main()
            paths = list((Path(temp_dir) / "Documents").glob("*.xlsx"))
            self.assertEqual(len(paths), 1)
            workbook = load_workbook(paths[0])
            self.addCleanup(workbook.close)
            summary = workbook["Summary"]
            rows = {row[0]: row for row in summary.iter_rows(min_row=2, values_only=True)}
            self.assertTrue(all(value is None for value in rows["FD1"][2:10]))
            self.assertEqual(rows["FD1"][10], "Data unavailable - see Fetch Issues")
            self.assertEqual(rows["FD2"][2], 1)
            self.assertTrue(all(value is None for value in rows["FD3"][2:10]))
            for row_number in (2, 4):
                self.assertTrue(all(cell.fill.fgColor.rgb == "00FFC7CE" for cell in summary[row_number]))
            self.assertNotEqual(summary["A3"].fill.fgColor.rgb, "00FFC7CE")
            issues = list(workbook["Fetch Issues"].values)
            self.assertEqual([row[0] for row in issues[1:]], ["FD1", "FD3"])
            self.assertIn("HTTP 504", issues[1][2])
            self.assertIn("HTTP 403", issues[2][2])
            self.assertIn("even if some pages were fetched", issues[1][3])
            self.assertTrue(all(cell.fill.fgColor.rgb == "00FFC7CE"
                                for row in workbook["Fetch Issues"].iter_rows(min_row=2)
                                for cell in row))
            detail_rows = list(workbook["Entity Aid Reports"].values)
            self.assertTrue(any("Data unavailable" in str(row[0]) for row in detail_rows))
            detail_sheet = workbook["Entity Aid Reports"]
            for row in detail_sheet.iter_rows():
                if row[0].value in ("FD1", "FD3") or str(row[0].value).startswith("Data unavailable"):
                    self.assertEqual(row[0].fill.fgColor.rgb, "00FFC7CE")
                elif row[0].value == "FD2":
                    self.assertNotEqual(row[0].fill.fgColor.rgb, "00FFC7CE")
            self.assertIn("temporarily disabled", workbook["Aid Given In Lieu Of"]["A1"].value)
            self.assertTrue(any("Incomplete" in str(row[1]) for row in workbook["Report Sources"].values))
            self.assertEqual(client.list_incidents.call_count, 9)

    def test_complete_report_retains_zero_counts_without_fetch_issues(self):
        with tempfile.TemporaryDirectory() as temp_dir, \
                patch.object(report.os.path, "expanduser", return_value=temp_dir):
            path = report.export_excel({}, {}, {}, {"FD1": "Department"}, {}, "MA")
            workbook = load_workbook(path)
            self.addCleanup(workbook.close)
            self.assertNotIn("Fetch Issues", workbook.sheetnames)
            self.assertEqual(workbook["Summary"]["C2"].value, 0)
            self.assertEqual(workbook["Summary"].max_column, 10)


if __name__ == "__main__":
    unittest.main()
