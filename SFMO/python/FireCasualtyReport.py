#!/usr/bin/env python3
# pyright: reportMissingImports=false
"""NERIS firefighter and civilian casualty report for fire incidents, by month."""

import os
import sys
from collections import defaultdict
from datetime import datetime, timezone

from EntityFireActivityReport import (
    DATE_RANGE_OPTIONS as _BASE_DATE_RANGE_OPTIONS,
    _metric_values,
    _type_strings,
    calculate_date_range as _base_calculate_date_range,
    date_range_label as _base_date_range_label,
    ensure_dependencies,
    fetch_entities,
    fetch_incidents,
)

from neris_auth import authenticate_with_fallback as _authenticate
from prompt_snippets import prompt_entity_scope, select_entities

CURRENT_YTD = "Occurred in current calendar year to date"
DATE_RANGE_OPTIONS = _BASE_DATE_RANGE_OPTIONS[:-1] + [CURRENT_YTD, _BASE_DATE_RANGE_OPTIONS[-1]]


def calculate_date_range(range_type, start_date=None, end_date=None):
    if range_type == CURRENT_YTD:
        now = datetime.now(timezone.utc)
        return datetime(now.year, 1, 1, tzinfo=timezone.utc), now
    return _base_calculate_date_range(range_type, start_date, end_date)


def date_range_label(range_type, start_date=None, end_date=None):
    if range_type == CURRENT_YTD:
        return f"{datetime.now().year}-ytd"
    return _base_date_range_label(range_type, start_date, end_date)


CASUALTY_COLUMNS = [
    "Number of Fires",
    "Number of FF Deaths", "Number of FF Injuries",
    "Number of nonFF Deaths", "Number of nonFF Injuries",
    "Number of FF Rescues", "Number of nonFF Rescues",
]


def authenticate_client_credentials():
    """Authenticate with an OAuth client; no email, password, or MFA needed."""
    return _authenticate(verify_ssl=False)


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Fire Casualty Report")
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


def incident_month(incident):
    value = (incident.get("dispatch") or {}).get("call_create") or ""
    return str(value)[:7] if len(str(value)) >= 7 else None


def aggregate_by_month(incidents):
    months = defaultdict(lambda: {column: 0 for column in CASUALTY_COLUMNS})
    for incident in incidents:
        if not _type_strings(incident):
            continue
        month = incident_month(incident)
        if not month:
            continue
        metrics = _metric_values(incident)
        months[month]["Number of Fires"] += 1
        for column in CASUALTY_COLUMNS[1:]:
            months[month][column] += metrics[column]
    return months


def aggregate_by_entity(incidents):
    entities = defaultdict(lambda: {column: 0 for column in CASUALTY_COLUMNS})
    for incident in incidents:
        if not _type_strings(incident):
            continue
        entity_id = (incident.get("base") or {}).get("department_neris_id", "")
        if not entity_id:
            continue
        metrics = _metric_values(incident)
        entities[entity_id]["Number of Fires"] += 1
        for column in CASUALTY_COLUMNS[1:]:
            entities[entity_id][column] += metrics[column]
    return entities


def month_sequence(months):
    if not months:
        return []
    first, last = min(months), max(months)
    year, month = int(first[:4]), int(first[5:7])
    result = []
    while f"{year:04d}-{month:02d}" <= last:
        result.append(f"{year:04d}-{month:02d}")
        month += 1
        if month > 12:
            year, month = year + 1, 1
    return result


def month_label(month):
    year, month_number = int(month[:4]), int(month[5:7])
    return f"{year} - {datetime(year, month_number, 1).strftime('%b')}"


def export_excel(months, entities, entity_names, scope_label, state_code, range_label, range_type,
                 fetch_errors=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from report_metadata import add_fetch_issues_sheet, highlight_fetch_issue

    fetch_errors = fetch_errors or {}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(
        os.path.expanduser("~"), "Documents",
        f"neris_fire_casualties_{state_code}_{range_label}_{timestamp}.xlsx",
    )
    os.makedirs(os.path.dirname(path), exist_ok=True)
    workbook = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    side = Side(style="thin", color="B7C9D6")
    border = Border(left=side, right=side, top=side, bottom=side)

    def style_header(sheet, row_number, count):
        for column_index in range(1, count + 1):
            cell = sheet.cell(row_number, column_index)
            cell.fill, cell.font, cell.border = header_fill, header_font, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Entity Summary sheet: one row per entity with casualty totals for the date range.
    entity_sheet = workbook.active
    entity_sheet.title = "Entity Summary"
    entity_sheet.append([f"Fire casualty summary by entity: {scope_label}"])
    entity_sheet.cell(entity_sheet.max_row, 1).font = Font(bold=True, size=12)
    entity_headers = ["Entity ID", "Entity Name"] + CASUALTY_COLUMNS
    entity_sheet.append(entity_headers)
    header_row = entity_sheet.max_row
    style_header(entity_sheet, header_row, len(entity_headers))
    entity_sheet.row_dimensions[header_row].height = 32

    entity_totals = {column: 0 for column in CASUALTY_COLUMNS}
    for entity_id in sorted(entity_names):
        values = entities.get(entity_id, {column: 0 for column in CASUALTY_COLUMNS})
        entity_sheet.append(
            [entity_id, entity_names[entity_id]]
            + [None if entity_id in fetch_errors else values[column] for column in CASUALTY_COLUMNS]
        )
        if entity_id not in fetch_errors:
            for column in CASUALTY_COLUMNS:
                entity_totals[column] += values[column]
        for column_index in range(1, len(entity_headers) + 1):
            entity_sheet.cell(entity_sheet.max_row, column_index).border = border
        if entity_id in fetch_errors:
            entity_sheet.cell(entity_sheet.max_row, len(entity_headers) + 1, "Data unavailable - see Fetch Issues")
            highlight_fetch_issue(entity_sheet[entity_sheet.max_row])
    if fetch_errors:
        entity_sheet.cell(header_row, len(entity_headers) + 1, "Fetch Status")
        style_header(entity_sheet, header_row, len(entity_headers) + 1)
        entity_sheet.column_dimensions[get_column_letter(len(entity_headers) + 1)].width = 38
    entity_sheet.append(["Total", ""] + [entity_totals[column] for column in CASUALTY_COLUMNS])
    for column_index in range(1, len(entity_headers) + 1):
        cell = entity_sheet.cell(entity_sheet.max_row, column_index)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")
        cell.border = border

    entity_sheet.column_dimensions["A"].width = 16
    entity_sheet.column_dimensions["B"].width = 40
    for column_index in range(3, len(entity_headers) + 1):
        entity_sheet.column_dimensions[get_column_letter(column_index)].width = 14

    # Monthly Breakdown sheet: casualty totals aggregated across all selected entities, by month.
    month_sheet = workbook.create_sheet("Monthly Breakdown")
    month_sheet.append([f"Monthly breakdown: {scope_label}"])
    month_sheet.cell(month_sheet.max_row, 1).font = Font(bold=True, size=12)
    headers = ["Month"] + CASUALTY_COLUMNS
    month_sheet.append(headers)
    header_row = month_sheet.max_row
    style_header(month_sheet, header_row, len(headers))
    month_sheet.row_dimensions[header_row].height = 32

    totals = {column: 0 for column in CASUALTY_COLUMNS}
    for month in month_sequence(months):
        values = months.get(month, {column: 0 for column in CASUALTY_COLUMNS})
        month_sheet.append([month_label(month)] + [values[column] for column in CASUALTY_COLUMNS])
        for column in CASUALTY_COLUMNS:
            totals[column] += values[column]
        for column_index in range(1, len(headers) + 1):
            month_sheet.cell(month_sheet.max_row, column_index).border = border
    month_sheet.append(["Total"] + [totals[column] for column in CASUALTY_COLUMNS])
    for column_index in range(1, len(headers) + 1):
        cell = month_sheet.cell(month_sheet.max_row, column_index)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")
        cell.border = border

    month_sheet.column_dimensions["A"].width = 16
    for column_index in range(2, len(headers) + 1):
        month_sheet.column_dimensions[get_column_letter(column_index)].width = 14
    if fetch_errors:
        entity_sheet.cell(1, 1, f"INCOMPLETE - {len(fetch_errors)} entities unavailable; see Fetch Issues")
        highlight_fetch_issue(entity_sheet[1][:1])
        month_sheet.cell(1, 1, f"INCOMPLETE - {len(fetch_errors)} entities unavailable; see Fetch Issues")
        highlight_fetch_issue(month_sheet[1][:1])

    from report_metadata import add_report_sources_sheet

    add_fetch_issues_sheet(workbook, fetch_errors, entity_names)
    add_report_sources_sheet(
        workbook, "FireCasualtyReport.py", range_label,
        [("State code", state_code), ("Scope", scope_label), ("Date range", range_type),
         ("Incident filter", "Incidents with a FIRE incident type"),
         ("Included sheets", "Entity Summary, Monthly Breakdown"
          + (", Fetch Issues" if fetch_errors else "")),
         ("Fetch status", f"Incomplete: {len(fetch_errors)} entities unavailable; see Fetch Issues."
          if fetch_errors else "Complete")],
    )
    workbook.save(path)
    return path


def main():
    ensure_dependencies()
    (state_code, scope, entity_id, entity_prefix, selected_names,
     range_type, start_date, end_date) = prompt_config()
    start, end = calculate_date_range(range_type, start_date, end_date)
    client = authenticate_client_credentials()

    all_entities = fetch_entities(client, state_code, None)
    try:
        entity_names = select_entities(all_entities, scope, entity_id, entity_prefix, selected_names)
    except ValueError as exc:
        sys.exit(str(exc))

    incidents = []
    failed = {}
    for current_entity in entity_names:
        print(f"Fetching {current_entity}...")
        try:
            incidents.extend(fetch_incidents(client, state_code, current_entity, start, end))
        except RuntimeError as error:
            print(f"  FAILED {current_entity}: {error}")
            failed[current_entity] = str(error)
    if failed:
        print(f"\nWARNING: {len(failed)} entities could not be retrieved and are missing from the report: "
              f"{', '.join(failed)}")

    scope_label = {"statewide": f"Entire state ({state_code})",
                   "fd25_prefix": f"County {entity_prefix}",
                   "entity": f"Entity {entity_id}",
                   "multiple_entities": f"Entities {', '.join(selected_names)}"}[scope]
    months = aggregate_by_month(incidents)
    entities = aggregate_by_entity(incidents)
    path = export_excel(months, entities, entity_names, scope_label, state_code,
                        date_range_label(range_type, start_date, end_date), range_type, failed)
    print(f"\nExported: {path}")
    print(f"Incidents retrieved: {len(incidents)}")


if __name__ == "__main__":
    main()
