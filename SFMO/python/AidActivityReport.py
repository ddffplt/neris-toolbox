#!/usr/bin/env python3
# pyright: reportMissingImports=false
"""NERIS aid activity report — incidents with aid given or received, by entity."""

import os
import subprocess
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from prompt_snippets import prompt_entity_scope, select_entities

# Temporarily disable the separate statewide collection for recipient-level details.
INCLUDE_AID_IN_LIEU_DETAILS = False

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
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()

    try:
        scope, entity_id, entity_prefix, selected_names = prompt_entity_scope(state_code)
    except ValueError as exc:
        sys.exit(str(exc))

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

    if not state_code:
        sys.exit("State code is required.")
    return state_code, scope, entity_id, entity_prefix, selected_names, range_type, start_date, end_date


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


def authenticate():
    from neris_auth import authenticate_with_fallback
    return authenticate_with_fallback()


class IncidentFetchError(RuntimeError):
    def __init__(self, message, status_code=None):
        super().__init__(message)
        self.status_code = status_code


class EntityFetchError(RuntimeError):
    pass


def _list_incidents_with_retry(client, entity_id, params):
    max_retries = 5
    for attempt in range(max_retries + 1):
        error = None
        try:
            response = client.list_incidents(**params)
            status_code = getattr(response, "status_code", None)
        except Exception as exc:
            error = exc
            status_code = getattr(getattr(exc, "response", None), "status_code", None)
            if status_code != 504:
                raise

        if status_code != 504:
            return response
        if attempt == max_retries:
            raise IncidentFetchError(
                f"NERIS API returned HTTP 504 while retrieving incidents for {entity_id} "
                f"after {max_retries} retries. Try running the report again later.",
                504,
            ) from error
        wait = 5 * 2 ** attempt
        print(f"  HTTP 504 for {entity_id}; retry {attempt + 1}/{max_retries} in {wait}s...")
        time.sleep(wait)


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
            response = _list_incidents_with_retry(client, entity_id, params)
        except IncidentFetchError:
            raise
        except Exception as error:
            status_code = getattr(getattr(error, "response", None), "status_code", None)
            if status_code in (401, 403) and cursor and not cursor_only:
                print("  API rejected the continuation cursor with filters; retrying with cursor only.")
                cursor_only = True
                continue
            if status_code is not None:
                raise IncidentFetchError(
                    f"NERIS API returned HTTP {status_code} while retrieving incidents for {entity_id}.",
                    status_code,
                ) from error
            raise IncidentFetchError(f"Unable to retrieve incidents for {entity_id}: {error}") from error

        status_code = getattr(response, "status_code", None)
        if status_code is not None and status_code >= 400:
            if status_code in (401, 403) and cursor and not cursor_only:
                print("  API rejected the continuation cursor with filters; retrying with cursor only.")
                cursor_only = True
                continue
            raise IncidentFetchError(
                f"NERIS API returned HTTP {status_code} while retrieving incidents for {entity_id}.",
                status_code,
            )
        if not isinstance(response, dict):
            try:
                response = response.json()
            except (AttributeError, ValueError) as error:
                status = f"HTTP {status_code}" if status_code is not None else "HTTP status unavailable"
                raise IncidentFetchError(
                    f"NERIS API returned a response that was not valid JSON while retrieving "
                    f"incidents for {entity_id} ({status})."
                ) from error
        if not isinstance(response, dict) or not isinstance(response.get("incidents"), list):
            raise IncidentFetchError(
                f"NERIS API returned an unexpected incident response for {entity_id}: "
                "expected an object containing an incidents list."
            )
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
    from requests.exceptions import RequestException

    while True:
        max_retries = 5
        for attempt in range(max_retries + 1):
            error = None
            try:
                response = client.list_entities(state=state_code, page_size=100, page_number=page)
                status_code = getattr(response, "status_code", None)
            except RequestException as exc:
                error = exc
                status_code = getattr(getattr(exc, "response", None), "status_code", None)
                if status_code is None:
                    raise EntityFetchError(
                        f"Unable to retrieve entities for {state_code} (page {page}): {exc}"
                    ) from exc

            if status_code not in (503, 504):
                if status_code is not None and status_code >= 400:
                    raise EntityFetchError(
                        f"NERIS API returned HTTP {status_code} while retrieving entities "
                        f"for {state_code} (page {page})."
                    ) from error
                if error is not None:
                    raise EntityFetchError(
                        f"Unable to retrieve entities for {state_code} (page {page}): {error}"
                    ) from error
                break
            if attempt == max_retries:
                raise EntityFetchError(
                    f"NERIS API returned HTTP {status_code} while retrieving entities for "
                    f"{state_code} (page {page}) after {max_retries} retries. "
                    "Try running the report again later."
                ) from error
            wait = 5 * 2 ** attempt
            print(f"  HTTP {status_code} for {state_code} entities, page {page}; "
                  f"retry {attempt + 1}/{max_retries} in {wait}s...")
            time.sleep(wait)

        if not isinstance(response, dict):
            try:
                response = response.json()
            except (AttributeError, ValueError) as error:
                status = f"HTTP {status_code}" if status_code is not None else "HTTP status unavailable"
                raise EntityFetchError(
                    f"NERIS API returned a response that was not valid JSON while retrieving "
                    f"entities for {state_code} (page {page}; {status})."
                ) from error
        if not isinstance(response, dict) or not isinstance(response.get("entities"), list):
            raise EntityFetchError(
                f"NERIS API returned an unexpected entity response for {state_code} "
                f"(page {page}): expected an object containing an entities list."
            )
        batch = response["entities"]
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
    """Group other departments' reports by recipient within the selected entities."""
    aid_in_lieu = defaultdict(list)
    for incident in incidents:
        giver_id = (incident.get("base") or {}).get("department_neris_id", "")
        recipients_seen = set()
        for aid in _normalized_aids(incident):
            if aid["direction"] != "GIVEN" or aid["type"] != "IN_LIEU_AID":
                continue
            recipient_id = aid["entity_id"]
            if (recipient_id not in entity_names or not giver_id or recipient_id == giver_id
                    or recipient_id in recipients_seen):
                continue
            recipients_seen.add(recipient_id)
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
                 range_label="selected-range", fetch_errors=None, fetch_entity_names=None):
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
    fetch_errors = fetch_errors or {}
    if fetch_entity_names is None:
        fetch_entity_names = entity_names
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    error_fill = PatternFill("solid", fgColor="FFC7CE")
    error_font = Font(color="9C0006", bold=True)
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
            row.append(None if entity_id in fetch_errors else values[column])
        summary_sheet.append(row)
        for column_index, column in enumerate(summary_headers, 1):
            apply_value_format(summary_sheet.cell(summary_sheet.max_row, column_index), column)
        if entity_id in fetch_errors:
            summary_sheet.cell(summary_sheet.max_row, len(summary_headers) + 1, "Data unavailable - see Fetch Issues")
            for cell in summary_sheet[summary_sheet.max_row]:
                cell.fill = error_fill
                cell.font = error_font
    if fetch_errors:
        summary_sheet.cell(1, len(summary_headers) + 1, "Fetch Status")
        style_header_row(summary_sheet, 1)
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
        if entity_id in fetch_errors:
            for cell in detail_sheet[heading_row][:5]:
                cell.fill = error_fill
                cell.font = error_font
            detail_sheet.append(["Data unavailable - see Fetch Issues. No incidents from this entity are included."])
            detail_sheet.cell(detail_sheet.max_row, 1).fill = error_fill
            detail_sheet.cell(detail_sheet.max_row, 1).font = error_font
        else:
            append_detail_table(entity_id)

    for row_number in range(1, detail_sheet.max_row + 1):
        if detail_sheet.cell(row_number, 1).value == "Incident Type":
            detail_sheet.row_dimensions[row_number].height = 48
    autosize(detail_sheet, 2, skip_columns={2})

    # ── Aid Given In Lieu Of sheet: incidents other entities reported giving in lieu of, by recipient ──
    lieu_sheet = workbook.create_sheet("Aid Given In Lieu Of")
    if not INCLUDE_AID_IN_LIEU_DETAILS:
        lieu_sheet.append(["Aid Given In Lieu Of incident details are temporarily disabled."])
    elif fetch_errors:
        lieu_sheet.append([
            "Incomplete: incidents from entities listed in Fetch Issues are excluded; "
            "aid given in lieu of may be missing."
        ])
        lieu_sheet["A1"].font = Font(bold=True, color="9C0006")
        lieu_sheet.append([])
    recipient_ids = sorted(aid_in_lieu) if INCLUDE_AID_IN_LIEU_DETAILS else []
    if INCLUDE_AID_IN_LIEU_DETAILS and not recipient_ids:
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
    if INCLUDE_AID_IN_LIEU_DETAILS:
        lieu_sheet.append([])
        lieu_sheet.append([
            f"Reporting scope: incidents reported by entities in {state_code} only; "
            "aid reported by out-of-state entities is not included."
        ])
    autosize(lieu_sheet, 2, skip_columns={2})

    if fetch_errors:
        issues_sheet = workbook.create_sheet("Fetch Issues")
        issues_sheet.append(["Entity ID", "Department Name", "Error", "Impact"])
        style_header_row(issues_sheet, 1)
        for entity_id, error in sorted(fetch_errors.items()):
            issues_sheet.append([
                entity_id, fetch_entity_names.get(entity_id, ""), str(error),
                "No incidents from this entity are included, even if some pages were fetched.",
            ])
        for row in issues_sheet.iter_rows(min_row=2):
            issues_sheet.row_dimensions[row[0].row].height = 60
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
                cell.fill = error_fill
                cell.font = error_font
        issues_sheet.column_dimensions["A"].width = 18
        issues_sheet.column_dimensions["B"].width = 35
        issues_sheet.column_dimensions["C"].width = 70
        issues_sheet.column_dimensions["D"].width = 70

    from report_metadata import add_report_sources_sheet

    included_sheets = "Summary, Entity Aid Reports, Aid Given In Lieu Of"
    if fetch_errors:
        included_sheets += ", Fetch Issues"
    add_report_sources_sheet(
        workbook,
        "AidActivityReport.py",
        range_label,
        [
            ("State code", state_code),
            ("Included sheets", included_sheets),
            ("Incident scope", "Incidents with at least one aid given or received record"),
            ("Aid Given In Lieu Of scope",
             f"Incidents reported by other entities in {state_code}, grouped by selected recipient; "
             "out-of-state reporting entities are not included" if INCLUDE_AID_IN_LIEU_DETAILS
             else "Incident details temporarily disabled; aggregate aid counts are still included."),
            ("Fetch status", f"Incomplete: {len(fetch_errors)} entities could not be retrieved; see Fetch Issues."
             if fetch_errors else "Complete"),
        ],
    )

    workbook.save(path)
    return path


def main():
    ensure_dependencies()
    state_code, scope, entity_id, entity_prefix, selected_names, range_type, start_date, end_date = prompt_config()
    start, end = calculate_date_range(range_type, start_date, end_date)
    client = authenticate()
    try:
        all_entities = fetch_entities(client, state_code, None)
    except EntityFetchError as exc:
        sys.exit(str(exc))
    try:
        entity_names = select_entities(all_entities, scope, entity_id, entity_prefix, selected_names)
    except ValueError as exc:
        sys.exit(str(exc))
    reporting_entity_names = dict(all_entities) if INCLUDE_AID_IN_LIEU_DETAILS else {}
    incidents = []
    reporting_incidents = []
    fetch_errors = {}
    fetch_entity_names = dict(entity_names)
    fetch_entity_names.update(reporting_entity_names)
    for current_entity in fetch_entity_names:
        print(f"Fetching {current_entity}...")
        try:
            fetched = fetch_incidents(client, state_code, current_entity, start, end)
            if current_entity in reporting_entity_names:
                reporting_incidents.extend(fetched)
            if current_entity in entity_names:
                incidents.extend(fetched)
        except IncidentFetchError as error:
            fetch_errors[current_entity] = str(error)
            print(f"  Could not retrieve {current_entity}: {error} Continuing; see Fetch Issues in the report.")
    detail, summary, aid_incidents, entity_names = aggregate(incidents, entity_names)

    aid_in_lieu = {}
    recipient_names = {}
    if INCLUDE_AID_IN_LIEU_DETAILS:
        name_cache = dict(reporting_entity_names)
        recipients = {
            eid: name for eid, name in entity_names.items()
            if state_code != "MA" or str(eid).upper().startswith("FD25")
        }
        aid_in_lieu = collect_aid_in_lieu(client, reporting_incidents, recipients, name_cache)
        recipient_names = {
            recipient_id: name_cache.get(recipient_id) or resolve_entity_name(client, recipient_id, name_cache)
            for recipient_id in aid_in_lieu
        }

    range_label = date_range_label(range_type, start_date, end_date)
    path = export_excel(
        detail, summary, aid_in_lieu, entity_names, recipient_names, state_code, range_label, fetch_errors,
        fetch_entity_names,
    )
    print(f"\nExported: {path}")
    print(f"Incidents retrieved: {len(incidents)}")
    print(f"Aid incidents: {len(aid_incidents)}; detail rows: {len(detail)}; departments: {len(entity_names)}")
    if INCLUDE_AID_IN_LIEU_DETAILS:
        print(f"Aid given in lieu of — recipient entities: {len(aid_in_lieu)}")
    else:
        print("Aid Given In Lieu Of incident details are temporarily disabled.")
    if fetch_errors:
        print(f"Warning: {len(fetch_errors)} entities could not be retrieved. See Fetch Issues in the report.")


if __name__ == "__main__":
    started_at = time.perf_counter()
    try:
        main()
    finally:
        elapsed = time.perf_counter() - started_at
        hours, remainder = divmod(elapsed, 3600)
        minutes, seconds = divmod(remainder, 60)
        print(f"\nTotal runtime: {int(hours):02d}:{int(minutes):02d}:{seconds:05.2f} (hh:mm:ss)")
