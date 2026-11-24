import importlib.util
import importlib
import io
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from prompt_snippets import prompt_entity_scope, select_entities


ENTITIES = {
    "FD25017001": "First Department",
    "FD25017002": "Second Department",
    "FD25025001": "Third Department",
}


class EntityScopeTests(unittest.TestCase):
    def test_selects_state_county_individual_and_multiple(self):
        self.assertEqual(select_entities(ENTITIES, "statewide"), ENTITIES)
        self.assertEqual(
            list(select_entities(ENTITIES, "fd25_prefix", entity_prefix="fd25017*")),
            ["FD25017001", "FD25017002"],
        )
        self.assertEqual(
            list(select_entities(ENTITIES, "entity", entity_id="first department")),
            ["FD25017001"],
        )
        self.assertEqual(
            list(select_entities(ENTITIES, "multiple_entities",
                                 entity_names=["FD25025001", "Second Department"])),
            ["FD25025001", "FD25017002"],
        )

    def test_missing_ambiguous_and_empty_scopes_fail(self):
        with self.assertRaisesRegex(ValueError, "No entity matched"):
            select_entities(ENTITIES, "multiple_entities", entity_names=["FD25017001", "missing"])
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            select_entities({"FD1": "Same", "FD2": "same"}, "entity", entity_id="same")
        with self.assertRaisesRegex(ValueError, "No entities matched"):
            select_entities(ENTITIES, "fd25_prefix", entity_prefix="FD25003")

    def test_prompt_validates_county_and_state(self):
        with patch("builtins.input", side_effect=["2", "FD25002", "fd25017*"]):
            self.assertEqual(prompt_entity_scope("MA"), ("fd25_prefix", None, "FD25017", []))
        with patch("builtins.input", return_value="2"):
            with self.assertRaisesRegex(ValueError, "Massachusetts"):
                prompt_entity_scope("VA")

    def test_tagged_report_prompts_for_scope(self):
        path = Path(__file__).parent / "client-credentials" / "StateActivityReport_w_IncTypeCount.py"
        spec = importlib.util.spec_from_file_location("client_state_report", path)
        report = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(report)
        with patch("builtins.input", side_effect=["MA", "4", "FD25017001, Third Department", "1"]):
            self.assertEqual(
                report.prompt_config(),
                ("MA", "multiple_entities", None, None,
                 ["FD25017001", "Third Department"], "All Records", None, None),
            )

    def test_tagged_report_only_fetches_selected_departments(self):
        path = Path(__file__).parent / "client-credentials" / "StateActivityReport_w_IncTypeCount.py"
        spec = importlib.util.spec_from_file_location("client_state_report_main", path)
        report = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(report)
        selections = [
            ("statewide", None, None, [], set(ENTITIES)),
            ("fd25_prefix", None, "FD25017", [], {"FD25017001", "FD25017002"}),
            ("entity", "FD25025001", None, [], {"FD25025001"}),
            ("multiple_entities", None, None, ["First Department", "FD25025001"],
             {"FD25017001", "FD25025001"}),
        ]
        for scope, entity_id, prefix, names, expected in selections:
            with self.subTest(scope=scope), \
                    patch("sys.stdout", new_callable=io.StringIO), \
                    patch.object(report, "ensure_dependencies"), \
                    patch.object(report, "prompt_config",
                                 return_value=("MA", scope, entity_id, prefix, names, "All Records", None, None)), \
                    patch.object(report, "authenticate", return_value=Mock()), \
                    patch.object(report, "fetch_all_entities", return_value=ENTITIES), \
                    patch.object(report, "fetch_incidents_for_entity",
                                 return_value={"months": {}, "types": {}, "total": 0, "multiple_types": 0}) as incidents, \
                    patch.object(report, "fetch_nars_for_entity", return_value=set()) as nars, \
                    patch.object(report, "generate_month_columns", return_value=["Jan-2025"]), \
                    patch.object(report, "export_to_excel", return_value="report.xlsx"):
                report.main()
                self.assertEqual({call.args[2] for call in incidents.call_args_list}, expected)
                self.assertEqual({call.args[2] for call in nars.call_args_list}, expected)

    def test_legacy_incident_exports_filter_before_fetching(self):
        modules = [
            "IncidentPull_StructureFires", "IncidentPull_ResidentialFires",
            "IncidentPull_SingleFamily", "IncidentPull_Multifamily",
            "IncidentPull_LithiumIonBatteries", "IncidentPull_VegetationFires",
        ]
        batch = [{"neris_id": eid, "name": name} for eid, name in ENTITIES.items()]
        for module_name in modules:
            module = importlib.import_module(module_name)
            client = Mock()
            client.list_entities.side_effect = [{"entities": batch, "total_count": len(batch)},
                                                {"entities": []}]
            kwargs = {"entity_scope": "multiple_entities",
                      "entity_names": ["First Department", "FD25025001"]}
            with self.subTest(module=module_name), \
                    patch.object(module, "get_department_incidents", return_value=[]) as fetch:
                incidents, entity_ids = module.get_statewide_incidents(client, state="MA", **kwargs)
                self.assertEqual(incidents, [])
                self.assertEqual(set(entity_ids), {"FD25017001", "FD25025001"})
                self.assertEqual({call.args[1] for call in fetch.call_args_list}, set(entity_ids))

    def test_legacy_incident_exports_fetch_only_county(self):
        modules = [
            "IncidentPull_StructureFires", "IncidentPull_ResidentialFires",
            "IncidentPull_SingleFamily", "IncidentPull_Multifamily",
            "IncidentPull_LithiumIonBatteries", "IncidentPull_VegetationFires",
        ]
        batch = [{"neris_id": eid, "name": name} for eid, name in ENTITIES.items()]
        for module_name in modules:
            module = importlib.import_module(module_name)
            client = Mock()
            client.list_entities.side_effect = [{"entities": batch, "total_count": len(batch)},
                                                {"entities": []}]
            prefix_arg = "county_prefix" if module_name == "IncidentPull_LithiumIonBatteries" else "entity_prefix"
            with self.subTest(module=module_name), \
                    patch.object(module, "get_department_incidents", return_value=[]) as fetch:
                incidents, entity_ids = module.get_statewide_incidents(
                    client, state="MA", entity_scope="fd25_prefix", **{prefix_arg: "FD25017"},
                )
                self.assertEqual(incidents, [])
                self.assertEqual(set(entity_ids), {"FD25017001", "FD25017002"})
                self.assertEqual({call.args[1] for call in fetch.call_args_list}, set(entity_ids))

    def test_unit_response_reports_fetch_selected_entities(self):
        batch = [{"neris_id": eid, "name": name} for eid, name in ENTITIES.items()]
        for module_name in ("IncidentPullMAUnitResponses", "IncidentPullMAUnitResponsestoMVAs"):
            module = importlib.import_module(module_name)
            client = Mock()
            client.list_entities.return_value = {"entities": batch, "total_count": len(batch)}
            with self.subTest(module=module_name), \
                    patch.object(module, "get_entity_incidents", return_value=[]) as fetch:
                module.get_scoped_incidents(
                    client, "MA", entity_scope="multiple_entities",
                    entity_names=["First Department", "FD25025001"],
                )
                self.assertEqual(
                    {call.args[1] for call in fetch.call_args_list},
                    {"FD25017001", "FD25025001"},
                )

    def test_lithium_report_passes_incidents_and_entity_ids_to_export(self):
        report = importlib.import_module("IncidentPull_LithiumIonBatteries")
        args = SimpleNamespace(
            range_type=None, start_date=None, end_date=None, entity_id=None,
            statewide=False, max_incidents=None, page_size=None, state=None,
            national=False, require_structure_battery=False, require_outside_battery=False,
        )
        incident = {"neris_id": "incident-1"}
        with patch("sys.stdout", new_callable=io.StringIO), \
                patch.object(report, "ensure_dependencies"), \
                patch.object(report, "parse_cli_args", return_value=args), \
                patch.object(report, "prompt_config",
                             return_value=("fd25_prefix", None, "FD25017", [], "All Records", None, None)), \
                patch.object(report, "authenticate", return_value=Mock()), \
                patch.object(report, "get_statewide_incidents",
                             return_value=([incident], ["FD25017001"])) as fetch, \
                patch.object(report, "filter_incidents_by_date", side_effect=lambda items, *_: items), \
                patch.object(report, "incident_matches_lithium_filter", return_value=True), \
                patch.object(report, "export_to_excel") as export:
            report.main()
            self.assertEqual(fetch.call_args.kwargs["county_prefix"], "FD25017")
            self.assertEqual(export.call_args.args[0], [incident])
            self.assertEqual(export.call_args.kwargs["entity_ids"], ["FD25017001"])


if __name__ == "__main__":
    unittest.main()
