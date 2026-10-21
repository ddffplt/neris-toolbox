#!/usr/bin/env python3
# pyright: reportMissingImports=false
"""NERIS aid activity report — incidents with aid given or received, by entity."""

import os
import subprocess
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone

DATE_RANGE_OPTIONS = [
    "All Records", "Occurred Today", "Occurred Yesterday",
    "Occurred Last 7 Days", "Occurred Last 30 Days", "Occurred Last 90 Days",
    "Occurred in 2025", "Custom Date Range",
]

METRIC_COLUMNS = [
    "Aid Given In Support of", "Aid Given In Lieu of", "Aid Given Acting As", "Total Aid Given",
    "Aid Received In Support of", "Aid Received In Lieu of", "Aid Received Acting As", "Total Aid Received",
]

# NERIS API aid_type enum values are ACTING_AS_AID / IN_LIEU_AID / SUPPORT_AID (not *_OF/ACTING_AS).
AID_COLUMNS = {
    ("GIVEN", "SUPPORT_AID"): "Aid Given In Support of",
    ("GIVEN", "IN_LIEU_AID"): "Aid Given In Lieu of",
    ("GIVEN", "ACTING_AS_AID"): "Aid Given Acting As",
    ("RECEIVED", "SUPPORT_AID"): "Aid Received In Support of",
    ("RECEIVED", "IN_LIEU_AID"): "Aid Received In Lieu of",
    ("RECEIVED", "ACTING_AS_AID"): "Aid Received Acting As",
}

# Incident date/time fields pulled from dispatch, each paired with an Eastern Time Zone conversion column.
DATE_FIELDS = ["Call Date"]

OVERVIEW_COLUMNS = [
    "Incident NERIS ID", "Aid-Giving Department ID", "Aid-Giving Department Name",
    "Incident Number",
    "Call Date", "Call Date (Eastern Time Zone)",
    "First Incident Type", "Second Incident Type", "Third Incident Type",
    "People Present", "Animals Rescued", "Outcome Narrative",
    "Displacement Count",
]


def ensure_dependencies():
    try:
        import openpyxl  # type: ignore[import-not-found]  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"], check=True)
    try:
        from neris_api_client import Config, NerisApiClient  # type: ignore[import-not-found]  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "neris-api-client", "--quiet"], check=True)


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Aid Activity Report by Entity")
    print("=" * 60)
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()

    print("\nDepartment Scope Options:")
    print("  1. Entire state")
    print("  2. County (entity ID prefix)")
    print("  3. Individual department")
    while True:
        scope_choice = input("Select scope [1]: ").strip() or "1"
        if scope_choice in {"1", "2", "3"}:
            break
        print("  Invalid selection, choose 1, 2, or 3.")

    entity_id = None
    entity_prefix = None
    if scope_choice == "2":
        entity_prefix = input("County entity ID prefix (e.g. FD25025): ").strip().upper() or None
        if not entity_prefix:
            print("  A prefix is required for this scope.")
            return prompt_config()
    elif scope_choice == "3":
        entity_id = input("NERIS Entity ID (e.g. FD26163151): ").strip() or None
        if not entity_id:
            print("  An entity ID is required for this scope.")
            return prompt_config()

    print("\nDate Range Options:")
    for number, option in enumerate(DATE_RANGE_OPTIONS, 1):
        print(f"  {number}. {option}")
    while True:
        choice = input("Select option [1]: ").strip() or "1"
        if choice.isdigit() and 1 <= int(choice) <= len(DATE_RANGE_OPTIONS):
            range_type = DATE_RANGE_OPTIONS[int(choice) - 1]
            break
        print("  Invalid selection, try again.")

    start_date = end_date = None
    if range_type == "Custom Date Range":
        while True:
            try:
                start_date = datetime.strptime(input("Start date (YYYY-MM-DD): ").strip(), "%Y-%m-%d").date()
                break
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")
        while True:
            try:
                end_date = datetime.strptime(input("End date (YYYY-MM-DD): ").strip(), "%Y-%m-%d").date()
                if end_date >= start_date:
                    break
                print("  End date must be on or after the start date.")
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")

    if not username or not password or not state_code:
        sys.exit("Email, password, and state code are required.")
    return username, password, state_code, entity_id, entity_prefix, range_type, start_date, end_date


def calculate_date_range(range_type, start_date=None, end_date=None):
    now = datetime.now(timezone.utc)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    if range_type == "All Records":
        return None, None
    if range_type == "Occurred Today":
        return today, None
    if range_type == "Occurred Yesterday":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday.replace(hour=23, minute=59, second=59)
    if range_type.startswith("Occurred Last "):
        days = int(range_type.split()[2])
        return today - timedelta(days=days), None
    if range_type.startswith("Occurred in "):
        year = int(range_type.split()[-1])
        return datetime(year, 1, 1, tzinfo=timezone.utc), datetime(year, 12, 31, 23, 59, 59, tzinfo=timezone.utc)
    if range_type == "Custom Date Range":
        return (datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc),
                datetime.combine(end_date, datetime.max.time().replace(microsecond=0), tzinfo=timezone.utc))
    return None, None


def date_range_label(range_type, start_date=None, end_date=None):
    """Return a compact, filename-safe label for the selected date range."""
    if range_type == "All Records":
        return "all-records"
    if range_type == "Custom Date Range":
        return f"{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}"
    if range_type.startswith("Occurred in "):
        return range_type.split()[-1]
    if range_type == "Occurred Today":
        return datetime.now().strftime("%Y-%m-%d")
    if range_type == "Occurred Yesterday":
        return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    if range_type.startswith("Occurred Last "):
        return f"last-{range_type.split()[2].lower()}-days"
    return "selected-range"


def authenticate(username, password):
    os.environ.update({
        "NERIS_BASE_URL": "https://api.neris.fsri.org/v1",
        "NERIS_GRANT_TYPE": "password",
        "NERIS_USERNAME": username,
        "NERIS_PASSWORD": password,
    })
    from neris_api_client import Config, NerisApiClient  # type: ignore[import-not-found]
    print("\nConnecting to NERIS API...")
    client = NerisApiClient(Config(
        base_url="https://api.neris.fsri.org/v1",
        grant_type="password",
        username=username,
        password=password,
    ))
    print("\nCHECK YOUR EMAIL FOR THE MFA CODE")
    client.list_incidents(page_size=1)
    print("Authentication successful!")
    return client


def fetch_incidents(client, state_code, entity_id, start=None, end=None, page_size=100):
    # neris_id_entity already scopes the query; adding "state" alongside it can trigger
    # spurious 403s if the entity's registered state record doesn't exactly match state_code.
    incidents = []
    cursor = None
    cursor_only = False
    while True:
        if cursor_only:
            params = {"cursor": cursor, "page_size": page_size}
        else:
            params = {"neris_id_entity": entity_id, "page_size": page_size}
            if start:
                params["call_create_start"] = start
            if end:
                params["call_create_end"] = end
            if cursor:
                params["cursor"] = cursor
        try:
            response = client.list_incidents(**params)
            status_code = getattr(response, "status_code", None)
            if status_code in (401, 403):
                if cursor and not cursor_only:
                    print("  API rejected the continuation cursor with filters; retrying with cursor only.")
                    cursor_only = True
                    continue
                try:
                    detail = response.json()
                except Exception:
                    detail = getattr(response, "text", "")
                raise RuntimeError(
                    f"NERIS API returned HTTP {status_code} while retrieving incidents for {entity_id}: {detail}"
                )
            response = response if isinstance(response, dict) else response.json()
        except Exception as error:
            status_code = getattr(getattr(error, "response", None), "status_code", None)
            if isinstance(error, RuntimeError) and "NERIS API returned HTTP 403" in str(error):
                raise
            raise RuntimeError(
                f"Unable to retrieve incidents for {entity_id}"
                f" (HTTP {status_code})" if status_code else
                f"Unable to retrieve incidents for {entity_id}: {error}"
            ) from error
        batch = response.get("incidents", [])
        if not batch:
            break
        incidents.extend(batch)
        cursor = response.get("next_cursor")
        if not cursor:
            break
        cursor_only = False
    return incidents


def fetch_entities(client, state_code, entity_id):
    if entity_id:
        try:
            entity = client.get_entity(entity_id)
            return {entity_id: entity.get("name", "") if isinstance(entity, dict) else ""}
        except Exception:
            return {entity_id: ""}
    entities = {}
    page = 1
    while True:
        response = client.list_entities(state=state_code, page_size=100, page_number=page)
        response = response if isinstance(response, dict) else response.json()
        batch = response.get("entities", [])
        if not batch:
            break
        for entity in batch:
            if entity.get("neris_id"):
                entities[entity["neris_id"]] = entity.get("name", "")
        if len(batch) < 100 or len(entities) >= response.get("total_count", 0):
            break
        page += 1
    return entities


def resolve_entity_name(client, entity_id, name_cache):
    if entity_id in name_cache:
        return name_cache[entity_id]
    try:
        entity = client.get_entity(entity_id)
        name = entity.get("name", "") if isinstance(entity, dict) else ""
    except Exception:
        name = ""
    name_cache[entity_id] = name
    return name


def _all_type_strings(incident):
    return {
        (item.get("type") or "").strip().upper()
        for item in incident.get("incident_types", []) or []
        if isinstance(item, dict) and item.get("type")
    } or {"UNSPECIFIED"}


def _normalized_aids(incident):
    aids = []
    for aid in incident.get("aids") or []:
        if not isinstance(aid, dict):
            continue
        aids.append({
            "direction": str(aid.get("aid_direction") or "").upper(),
            "type": str(aid.get("aid_type") or "").upper().replace("-", "_").replace(" ", "_"),
            "entity_id": aid.get("department_neris_id", ""),
        })
    return aids


def _has_aid(incident):
    return bool(incident.get("aids"))


def _metric_values(incident):
    values = {column: 0 for column in METRIC_COLUMNS}
    for aid in _normalized_aids(incident):
        column = AID_COLUMNS.get((aid["direction"], aid["type"]))
        if column:
            values[column] += 1
    values["Total Aid Given"] = (
        values["Aid Given In Support of"] + values["Aid Given In Lieu of"] + values["Aid Given Acting As"]
    )
    values["Total Aid Received"] = (
        values["Aid Received In Support of"] + values["Aid Received In Lieu of"] + values["Aid Received Acting As"]
    )
    return values


def _overview_fields(incident, department_name):
    base = incident.get("base") or {}
    dispatch = incident.get("dispatch") or {}
    inc_types = incident.get("incident_types", []) or []
    type_list = [
        item.get("type", "")
        for item in inc_types
        if isinstance(item, dict) and item.get("type")
    ]
    return {
        "Incident NERIS ID": incident.get("neris_id", ""),
        "Aid-Giving Department ID": base.get("department_neris_id", ""),
        "Aid-Giving Department Name": department_name,
        "Incident Number": base.get("incident_number", ""),
        "Call Date": dispatch.get("call_create", ""),
        "First Incident Type": type_list[0] if len(type_list) > 0 else "",
        "Second Incident Type": type_list[1] if len(type_list) > 1 else "",
        "Third Incident Type": type_list[2] if len(type_list) > 2 else "",
        "People Present": base.get("people_present", ""),
        "Animals Rescued": base.get("animals_rescued", ""),
        "Outcome Narrative": base.get("outcome_narrative", ""),
        "Displacement Count": base.get("displacement_count", ""),
    }


def aggregate(incidents, entity_names):
    detail = defaultdict(lambda: {column: 0 for column in METRIC_COLUMNS})
    summary = defaultdict(lambda: {column: 0 for column in METRIC_COLUMNS})
    aid_incidents = [incident for incident in incidents if _has_aid(incident)]

    for incident in aid_incidents:
        entity_id = (incident.get("base") or {}).get("department_neris_id", "")
        if not entity_id:
            continue
        metrics = _metric_values(incident)
        incident_types = _all_type_strings(incident)
        for incident_type in incident_types:
            key = (entity_id, incident_type)
            for column in METRIC_COLUMNS:
                detail[key][column] += metrics[column]
        # Increment the entity total once per incident regardless of type count.
        for column in METRIC_COLUMNS:
            summary[entity_id][column] += metrics[column]

    return detail, summary, aid_incidents, entity_names


def collect_aid_in_lieu(client, incidents, entity_names, name_cache):
    """Group incidents by the recipient entity that other departments reported as aid given in lieu of."""
    aid_in_lieu = defaultdict(list)
    for incident in incidents:
        giver_id = (incident.get("base") or {}).get("department_neris_id", "")
        for aid in _normalized_aids(incident):
            if aid["direction"] != "GIVEN" or aid["type"] != "IN_LIEU_AID":
                continue
            recipient_id = aid["entity_id"]
            if not recipient_id:
                continue
            giver_name = entity_names.get(giver_id) or resolve_entity_name(client, giver_id, name_cache)
            aid_in_lieu[recipient_id].append(_overview_fields(incident, giver_name))
    return aid_in_lieu


def _eastern_formula(utc_cell):
    """Build an Excel formula converting a UTC ISO-8601 string cell to US Eastern time (DST-aware)."""
    return (
        f'=IF({utc_cell}="","",'
        f'DATE(VALUE(LEFT({utc_cell},4)),VALUE(MID({utc_cell},6,2)),'
        f'VALUE(MID({utc_cell},9,2)))+'
        f'TIME(VALUE(MID({utc_cell},12,2)),VALUE(MID({utc_cell},15,2)),'
        f'VALUE(MID({utc_cell},18,2)))-IF('
        f'AND(DATE(VALUE(LEFT({utc_cell},4)),VALUE(MID({utc_cell},6,2)),'
        f'VALUE(MID({utc_cell},9,2)))+TIME(VALUE(MID({utc_cell},12,2)),'
        f'VALUE(MID({utc_cell},15,2)),VALUE(MID({utc_cell},18,2)))>='
        f'DATE(VALUE(LEFT({utc_cell},4)),3,1)+'
        f'MOD(8-WEEKDAY(DATE(VALUE(LEFT({utc_cell},4)),3,1)),7)+7+TIME(7,0,0),'
        f'DATE(VALUE(LEFT({utc_cell},4)),3,1)+'
        f'MOD(8-WEEKDAY(DATE(VALUE(LEFT({utc_cell},4)),3,1)),7)+7+TIME(7,0,0)<'
        f'DATE(VALUE(LEFT({utc_cell},4)),11,1)+'
        f'MOD(8-WEEKDAY(DATE(VALUE(LEFT({utc_cell},4)),11,1)),7)+TIME(6,0,0)), '
        f'TIME(4,0,0),TIME(5,0,0)))'
    )


def export_excel(detail, summary, aid_in_lieu, entity_names, recipient_names, state_code,
                 range_label="selected-range"):
    from openpyxl import Workbook  # type: ignore[import-not-found]
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore[import-not-found]
    from openpyxl.utils import get_column_letter  # type: ignore[import-not-found]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(
        os.path.expanduser("~"), "Documents",
        f"neris_aid_activity_{state_code}_{range_label}_{timestamp}.xlsx",
    )
    os.makedirs(os.path.dirname(path), exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(*(Side(style="thin", color="B7C9D6") for _ in range(4)))
    heading_border = Border(*(Side(style="thin", color="000000") for _ in range(4)))

    def style_header_row(target_sheet, row_number):
        for cell in target_sheet[row_number]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    def apply_value_format(cell, column):
        cell.border = border
        if column in METRIC_COLUMNS:
            cell.number_format = '#,##0;(#,##0);-'

    def autosize(target_sheet, first_data_row, skip_columns=()):
        for column_index in range(1, target_sheet.max_column + 1):
            if column_index in skip_columns:
                continue
            column_letter = get_column_letter(column_index)
            longest_value = 0
            for row_number in range(1, target_sheet.max_row + 1):
                cell = target_sheet.cell(row_number, column_index)
                if cell.value is None:
                    continue
                if cell.number_format == '#,##0;(#,##0);-':
                    value_length = len(f"{round(float(cell.value or 0)):,}")
                else:
                    value_length = len(str(cell.value))
                longest_value = max(longest_value, value_length)
            target_sheet.column_dimensions[column_letter].width = max(10, min(longest_value + 2, 40))

    # ── Summary sheet: one row per department, totals across all incident types ──
    summary_sheet = workbook.create_sheet("Summary")
    summary_headers = ["Entity ID", "Department Name"] + METRIC_COLUMNS
    summary_sheet.append(summary_headers)
    style_header_row(summary_sheet, summary_sheet.max_row)
    for entity_id in sorted(set(entity_names) | set(summary)):
        values = summary.get(entity_id, {column: 0 for column in METRIC_COLUMNS})
        row = [entity_id, entity_names.get(entity_id, "")]
        for column in METRIC_COLUMNS:
            row.append(values[column])
        summary_sheet.append(row)
        for column_index, column in enumerate(summary_headers, 1):
            apply_value_format(summary_sheet.cell(summary_sheet.max_row, column_index), column)
    summary_sheet.row_dimensions[1].height = 48
    autosize(summary_sheet, 2)

    # ── Entity Aid Reports sheet: incident-type breakdown per department ──
    detail_sheet = workbook.create_sheet("Entity Aid Reports")
    detail_headers = ["Incident Type"] + METRIC_COLUMNS

    def append_detail_table(entity_id):
        detail_sheet.append(detail_headers)
        style_header_row(detail_sheet, detail_sheet.max_row)
        rows_for_entity = sorted(key for key in detail if key[0] == entity_id)
        totals = {column: 0 for column in METRIC_COLUMNS}

        def append_row(label, values):
            row = [label]
            for column in METRIC_COLUMNS:
                row.append(values[column])
            detail_sheet.append(row)
            for column_index, column in enumerate(detail_headers, 1):
                apply_value_format(detail_sheet.cell(detail_sheet.max_row, column_index), column)

        for key in rows_for_entity:
            append_row(key[1], detail[key])
            for column in METRIC_COLUMNS:
                totals[column] += detail[key][column]
        append_row("Total", totals)
        total_row = detail_sheet.max_row
        for cell in detail_sheet[total_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="E2F0D9")

    entity_ids = sorted(set(entity_names) | {key[0] for key in detail})
    for entity_index, entity_id in enumerate(entity_ids):
        if entity_index:
            detail_sheet.append([])
            detail_sheet.append([])
        detail_sheet.append([entity_id, entity_names.get(entity_id, "")])
        heading_row = detail_sheet.max_row
        detail_sheet.merge_cells(start_row=heading_row, start_column=2, end_row=heading_row, end_column=5)
        detail_sheet.row_dimensions[heading_row].height = 48
        for column_index in range(1, 6):
            cell = detail_sheet.cell(heading_row, column_index)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.border = heading_border
            if column_index == 1:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            elif column_index == 2:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        append_detail_table(entity_id)

    for row_number in range(1, detail_sheet.max_row + 1):
        if detail_sheet.cell(row_number, 1).value == "Incident Type":
            detail_sheet.row_dimensions[row_number].height = 48
    autosize(detail_sheet, 2, skip_columns={2})

    # ── Aid Given In Lieu Of sheet: incidents other entities reported giving in lieu of, by recipient ──
    lieu_sheet = workbook.create_sheet("Aid Given In Lieu Of")
    recipient_ids = sorted(aid_in_lieu)
    if not recipient_ids:
        lieu_sheet.append(["No aid-given-in-lieu-of incidents found for the selected scope and date range."])
    for entity_index, recipient_id in enumerate(recipient_ids):
        if entity_index:
            lieu_sheet.append([])
            lieu_sheet.append([])
        recipient_name = recipient_names.get(recipient_id, "")
        lieu_sheet.append([recipient_id, recipient_name])
        heading_row = lieu_sheet.max_row
        lieu_sheet.merge_cells(
            start_row=heading_row, start_column=2, end_row=heading_row, end_column=len(OVERVIEW_COLUMNS)
        )
        lieu_sheet.row_dimensions[heading_row].height = 24
        for column_index in range(1, len(OVERVIEW_COLUMNS) + 1):
            cell = lieu_sheet.cell(heading_row, column_index)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.border = heading_border
            if column_index == 1:
                cell.font = Font(bold=True, size=12)
            elif column_index == 2:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        lieu_sheet.append(OVERVIEW_COLUMNS)
        style_header_row(lieu_sheet, lieu_sheet.max_row)
        for record in aid_in_lieu[recipient_id]:
            row = [record.get(column, "") for column in OVERVIEW_COLUMNS]
            lieu_sheet.append(row)
            row_number = lieu_sheet.max_row
            for column_index in range(1, len(OVERVIEW_COLUMNS) + 1):
                lieu_sheet.cell(row_number, column_index).border = border
            for field in DATE_FIELDS:
                utc_col_idx = OVERVIEW_COLUMNS.index(field) + 1
                eastern_col_idx = OVERVIEW_COLUMNS.index(f"{field} (Eastern Time Zone)") + 1
                utc_cell = f"{get_column_letter(utc_col_idx)}{row_number}"
                eastern_cell = lieu_sheet.cell(row_number, eastern_col_idx)
                eastern_cell.value = _eastern_formula(utc_cell)
                eastern_cell.number_format = "mm/dd/yy hh:mm"
                eastern_cell.border = border

    for row_number in range(1, lieu_sheet.max_row + 1):
        if lieu_sheet.cell(row_number, 1).value == "Incident NERIS ID":
            lieu_sheet.row_dimensions[row_number].height = 32
    autosize(lieu_sheet, 2, skip_columns={2})

    from report_metadata import add_report_sources_sheet

    add_report_sources_sheet(
        workbook,
        "AidActivityReport.py",
        range_label,
        [
            ("State code", state_code),
            ("Included sheets", "Summary, Entity Aid Reports, Aid Given In Lieu Of"),
            ("Incident scope", "Incidents with at least one aid given or received record"),
            ("Aid Given In Lieu Of scope", "Incidents reported by other entities as aid given in lieu of, grouped by recipient"),
        ],
    )

    workbook.save(path)
    return path


def main():
    ensure_dependencies()
    username, password, state_code, entity_id, entity_prefix, range_type, start_date, end_date = prompt_config()
    start, end = calculate_date_range(range_type, start_date, end_date)
    client = authenticate(username, password)
    entity_names = fetch_entities(client, state_code, entity_id)
    if entity_prefix:
        entity_names = {
            eid: name for eid, name in entity_names.items()
            if str(eid).upper().startswith(entity_prefix)
        }
    incidents = []
    for current_entity in entity_names:
        print(f"Fetching {current_entity}...")
        try:
            incidents.extend(fetch_incidents(client, state_code, current_entity, start, end))
        except RuntimeError as error:
            if "HTTP 403" in str(error):
                print(f"  Skipping {current_entity}: account is not authorized to retrieve this entity.")
                continue
            raise
    if not entity_names:
        print("No departments found for the selected state/entity.")
        return

    detail, summary, aid_incidents, entity_names = aggregate(incidents, entity_names)

    name_cache = dict(entity_names)
    aid_in_lieu = collect_aid_in_lieu(client, aid_incidents, entity_names, name_cache)
    recipient_names = {
        recipient_id: name_cache.get(recipient_id) or resolve_entity_name(client, recipient_id, name_cache)
        for recipient_id in aid_in_lieu
    }

    range_label = date_range_label(range_type, start_date, end_date)
    path = export_excel(
        detail, summary, aid_in_lieu, entity_names, recipient_names, state_code, range_label
    )
    print(f"\nExported: {path}")
    print(f"Incidents retrieved: {len(incidents)}")
    print(f"Aid incidents: {len(aid_incidents)}; detail rows: {len(detail)}; departments: {len(entity_names)}")
    print(f"Aid given in lieu of — recipient entities: {len(aid_in_lieu)}")


if __name__ == "__main__":
    main()
