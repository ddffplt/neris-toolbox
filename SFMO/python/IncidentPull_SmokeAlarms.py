"""
NERIS Smoke Alarm Report
------------------------
Fetches residential Structure Fire incidents from the NERIS API and reports on
every smoke alarm field:

  smoke_alarm.presence.type              PRESENT / NOT_PRESENT / NOT_APPLICABLE
  smoke_alarm.presence.working           working (True/False)
  smoke_alarm.presence.alarm_types       alarm type(s) present
  smoke_alarm.presence.operation         operated / failed / other, with
                                         occupant action and failure reason

Workbook sheets: Summary, Individual Counts, By Location Group, Raw Data,
Report Sources. Summary, Individual Counts and By Location Group are
hierarchical: indented, color-coded rows are parts of the row above them.

Uses the same built-in client-credentials login as FireCasualtyReport.py
(no email, password, or MFA prompt).
Entity scope and date range options match IncidentPullFalseAlarm.py, whose
query helpers are reused here.
"""

import sys
import pathlib
from collections import Counter
from datetime import datetime

from FireCasualtyReport import authenticate_client_credentials
from IncidentPullFalseAlarm import (
    DATE_RANGE_OPTIONS,
    calculate_date_range,
    ensure_dependencies,
    get_scoped_incidents,
    normalize_fd25_prefix,
)

# Location-use category (first segment of base.location_use.use_type) to keep.
LOCATION_USE_CATEGORY = "RESIDENTIAL"
NO_DATA = "(No data reported)"

# Location-use subtype (second segment of use_type) to report group.
LOCATION_GROUPS = [
    ("Single Family Homes", {"ATTACHED_SINGLE_FAMILY_DWELLING", "DETATCHED_SINGLE_FAMILY_DWELLING",
                             "DETACHED_SINGLE_FAMILY_DWELLING"}),
    ("Multifamily Dwellings", {"MULTI_FAMILY_LOWRISE_DWELLING", "MULTI_FAMILY_MIDRISE_DWELLING",
                               "MULTI_FAMILY_HIGHRISE_DWELLING"}),
    ("Temporary Lodging (Hotel/Motel)", {"TEMPORARY_LODGING_HOTEL_MOTEL"}),
    ("Unhoused Temporary Housing", {"UNHOUSED_TEMPORARY_SHELTER"}),
    ("Congregate Housing", {"CONGREGATE_HOUSING"}),
    ("Manufactured - Mobile Home", {"MANUFACTURED_MOBILE_HOME"}),
    ("Detached Residential Garage", {"DETATCHED_GARAGE", "DETACHED_GARAGE"}),
]
OTHER_GROUP = "Other Residential"

ENTITY_SCOPE_LABELS = {
    "entity": "Individual entity",
    "fd25_prefix": "FD25 county prefix",
    "statewide": "Entire state",
    "multiple_entities": "Multiple entities",
}

OPERATION_LABELS = {
    "OPERATED_ALERTED_OCCUPANT": "Operated - Alerted Occupant",
    "OPERATED_FAILED_TO_ALERT_OCCUPANT": "Operated - Failed to Alert Occupant",
    "NO_OCCUPANT_TO_NOTIFY": "No Occupant to Notify",
    "INSUFFICIENT_SOURCE": "Insufficient Source",
    "FAILED_TO_OPERATE": "Failed to Operate",
}

# Row styling by hierarchy level: (fill color, bold)
LEVEL_STYLES = {
    0: ("DDEBF7", True),
    1: ("FFF2CC", False),
    2: ("E2EFDA", False),
}


def pretty(value):
    if not value or value == NO_DATA:
        return value
    return str(value).replace("DETATCHED", "DETACHED").replace("_", " ").title()


def location_parts(use_type):
    parts = str(use_type or "").strip().upper().split("||")
    return parts[1] if len(parts) > 1 else ""


def location_group(use_type):
    subtype = location_parts(use_type)
    for name, subtypes in LOCATION_GROUPS:
        if subtype in subtypes:
            return name
    return OTHER_GROUP


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Smoke Alarm Report")
    print("=" * 60)

    print("\n-- Query Parameters -----------------------------")
    state_code = input("State Code (e.g. MA, MI, VA): ").strip().upper()

    print("\nEntity Scope Options:")
    print("  1. Entire state")
    print("  2. County (FD25 prefix)")
    print("  3. Single entity")
    print("  4. Multiple entities")
    while True:
        scope_choice = input("Select entity scope [1]: ").strip() or "1"
        if scope_choice in {"1", "2", "3", "4"}:
            break
        print("  Invalid selection, choose 1, 2, 3, or 4.")

    entity_id = None
    entity_prefix = None
    entity_names = []
    if scope_choice == "2":
        while True:
            entity_prefix = normalize_fd25_prefix(
                input("FD25 county prefix (FD25001*, FD25003*, ... FD25027*): ")
            )
            if entity_prefix:
                break
            print("  Invalid prefix. Use FD25 followed by an odd county code from 001 through 027.")
    elif scope_choice == "3":
        while not entity_id:
            entity_id = input("NERIS Entity ID or exact entity name: ").strip()
    elif scope_choice == "4":
        while not entity_names:
            entity_names = [
                value.strip()
                for value in input("Entity IDs or exact names, separated by commas: ").split(",")
                if value.strip()
            ]

    entity_scope = {
        "1": "statewide",
        "2": "fd25_prefix",
        "3": "entity",
        "4": "multiple_entities",
    }[scope_choice]

    print("\nDate Range Options:")
    for i, opt in enumerate(DATE_RANGE_OPTIONS, 1):
        print(f"  {i}. {opt}")
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
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")
                continue
            if end_date < start_date:
                print("  End date must be on or after the start date.")
                continue
            break

    return (
        state_code, entity_id, entity_prefix, entity_names,
        entity_scope, range_type, start_date, end_date,
    )


def is_residential_structure_fire(incident):
    types = []
    for item in incident.get("incident_types") or []:
        value = item.get("type") if isinstance(item, dict) else item
        if value:
            types.append(str(value).upper())
    if not any("STRUCTURE_FIRE" in value for value in types):
        return False

    base = incident.get("base") or {}
    location_use = base.get("location_use") or {}
    use_type = location_use.get("use_type") if isinstance(location_use, dict) else None
    return str(use_type or "").strip().upper().split("||", 1)[0] == LOCATION_USE_CATEGORY


def extract_smoke_alarm(incident):
    """Flatten the smoke_alarm object into one dict of report fields."""
    smoke = incident.get("smoke_alarm") or {}
    presence = smoke.get("presence") or {}
    operation = presence.get("operation") or {}
    outcome = operation.get("alerted_failed_other") or {}

    alarm_types = presence.get("alarm_types") or []
    working = presence.get("working")
    return {
        "presence": presence.get("type") or NO_DATA,
        "working": NO_DATA if working is None else ("Yes" if working else "No"),
        "alarm_types": [str(t) for t in alarm_types if t],
        "operation": outcome.get("type") or NO_DATA,
        "occupant_action": outcome.get("occupant_action") or NO_DATA,
        "failure_reason": outcome.get("failure_reason") or NO_DATA,
    }


def pct(count, denominator):
    return round(count / denominator, 4) if denominator else 0


def build_tables(incidents):
    """Return (summary_rows, count_rows, raw_rows, group_rows).

    summary_rows : (label, incidents, percent, percent basis, level)
    count_rows   : (field, value, incidents, percent, level)
    group_rows   : (label, incidents, smoke present, %, working, %, fire present, %, level)
    """
    total = len(incidents)
    rows = []
    for incident in incidents:
        base = incident.get("base") or {}
        dispatch = incident.get("dispatch") or {}
        location = base.get("location") or {}
        use_type = (base.get("location_use") or {}).get("use_type", "")
        fields = extract_smoke_alarm(incident)
        rows.append({
            "Incident NERIS ID": incident.get("neris_id", ""),
            "Department/Entity ID": base.get("department_neris_id", ""),
            "Incident Number": base.get("incident_number", ""),
            "Call Date": str(dispatch.get("call_create", "") or ""),
            "City": location.get("incorporated_municipality", "") if isinstance(location, dict) else "",
            "County": location.get("county", "") if isinstance(location, dict) else "",
            "Location Use": use_type,
            "Location Group": location_group(use_type),
            "Location Subtype": pretty(location_parts(use_type)) or NO_DATA,
            "Fire Alarm Presence": (((incident.get("fire_alarm") or {}).get("presence") or {}).get("type")
                                    or NO_DATA),
            "Smoke Alarm Presence": fields["presence"],
            "Smoke Alarm Working": fields["working"],
            "Smoke Alarm Types": "||".join(fields["alarm_types"]),
            "Smoke Alarm Operation": fields["operation"],
            "Occupant Action": fields["occupant_action"],
            "Failure Reason": fields["failure_reason"],
            "_fields": fields,
        })

    group_names = [name for name, _ in LOCATION_GROUPS] + [OTHER_GROUP]
    subtype_counts = {
        name: Counter(r["Location Subtype"] for r in rows if r["Location Group"] == name)
        for name in group_names
    }
    shown_groups = [n for n in group_names if n != OTHER_GROUP or subtype_counts[n]]

    # ---- Individual Counts -------------------------------------------------
    counters = {
        "Smoke Alarm Presence": Counter(r["_fields"]["presence"] for r in rows),
        "Smoke Alarm Working": Counter(r["_fields"]["working"] for r in rows),
        "Smoke Alarm Type": Counter(
            t for r in rows for t in (r["_fields"]["alarm_types"] or [NO_DATA])
        ),
        "Smoke Alarm Operation": Counter(r["_fields"]["operation"] for r in rows),
        "Occupant Action": Counter(r["_fields"]["occupant_action"] for r in rows),
        "Failure Reason": Counter(r["_fields"]["failure_reason"] for r in rows),
        "Fire Alarm Presence": Counter(r["Fire Alarm Presence"] for r in rows),
    }
    count_rows = []
    for field, counter in counters.items():
        for value, count in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0])):
            count_rows.append((field, value, count, pct(count, total), 0))
    for name in shown_groups:
        group_total = sum(subtype_counts[name].values())
        count_rows.append(("Location Group", name, group_total, pct(group_total, total), 0))
        for subtype, count in sorted(subtype_counts[name].items(), key=lambda kv: (-kv[1], kv[0])):
            count_rows.append(("Location Group", subtype, count, pct(count, total), 1))

    # ---- Summary -----------------------------------------------------------
    present = [r for r in rows if r["_fields"]["presence"] == "PRESENT"]
    n_present = len(present)
    presence = counters["Smoke Alarm Presence"]
    summary = []

    def add(label, count, denominator, basis, level):
        summary.append((label, count, pct(count, denominator), basis, level))

    summary.append(("Total residential structure fire incidents", total, "", "", 0))
    summary.append(("", "", "", "", 0))

    summary.append(("SMOKE ALARM PRESENCE", "", "", "", 0))
    add("Smoke Alarm Present", n_present, total, "All incidents", 0)

    def add_breakdown(key, order, labels, parent_rows, level, basis):
        denominator = len(parent_rows)
        counts = Counter(r["_fields"][key] for r in parent_rows)
        for value in order:
            if counts.get(value):
                add(labels.get(value, pretty(value)), counts[value], denominator, basis, level)
        for value in sorted(set(counts) - set(order) - {NO_DATA}):
            add(labels.get(value, pretty(value)), counts[value], denominator, basis, level)
        if counts.get(NO_DATA):
            add("Not reported", counts[NO_DATA], denominator, basis, level)

    # Working status under Present
    summary.append(("Smoke Alarm Working", "", "", "", 1))
    for value, label in (("Yes", "Working"), ("No", "Not working"), (NO_DATA, "Not reported")):
        add(label, sum(1 for r in present if r["_fields"]["working"] == value),
            n_present, "Present incidents", 2)

    # Operation under Present; occupant action / failure reason nested under their operation
    summary.append(("Smoke Alarm Operation", "", "", "", 1))
    operation_order = [
        "OPERATED_ALERTED_OCCUPANT", "OPERATED_FAILED_TO_ALERT_OCCUPANT",
        "NO_OCCUPANT_TO_NOTIFY", "INSUFFICIENT_SOURCE", "FAILED_TO_OPERATE",
    ]
    operation_counts = Counter(r["_fields"]["operation"] for r in present)
    for op in operation_order + sorted(set(operation_counts) - set(operation_order) - {NO_DATA}):
        count = operation_counts.get(op, 0)
        if not count and op not in ("OPERATED_ALERTED_OCCUPANT", "FAILED_TO_OPERATE"):
            continue
        add(OPERATION_LABELS.get(op, pretty(op)), count, n_present, "Present incidents", 2)
        members = [r for r in present if r["_fields"]["operation"] == op]
        if op == "OPERATED_ALERTED_OCCUPANT":
            summary.append(("Occupant Action", "", "", "", 2))
            add_breakdown("occupant_action", [], {}, members, 3, "Operated - Alerted Occupant incidents")
        elif op == "FAILED_TO_OPERATE":
            summary.append(("Smoke Alarm Failure Reason", "", "", "", 2))
            add_breakdown("failure_reason", [], {}, members, 3, "Failed to Operate incidents")
    if operation_counts.get(NO_DATA):
        add("Not reported", operation_counts[NO_DATA], n_present, "Present incidents", 2)

    add("Smoke Alarm Not Present", presence.get("NOT_PRESENT", 0), total, "All incidents", 0)
    add("Smoke Alarm Not Applicable", presence.get("NOT_APPLICABLE", 0), total, "All incidents", 0)
    add("Smoke Alarm Presence Not Reported", presence.get(NO_DATA, 0), total, "All incidents", 0)
    summary.append(("", "", "", "", 0))

    summary.append(("FIRE ALARM PRESENCE", "", "", "", 0))
    fire_presence = counters["Fire Alarm Presence"]
    for label, key in (("Fire Alarm Present", "PRESENT"), ("Fire Alarm Not Present", "NOT_PRESENT"),
                       ("Fire Alarm Not Applicable", "NOT_APPLICABLE"),
                       ("Fire Alarm Presence Not Reported", NO_DATA)):
        add(label, fire_presence.get(key, 0), total, "All incidents", 0)
    summary.append(("", "", "", "", 0))

    summary.append(("LOCATION GROUPS", "", "", "", 0))
    for name in shown_groups:
        group_total = sum(subtype_counts[name].values())
        add(name, group_total, total, "All incidents", 0)
        for subtype, count in sorted(subtype_counts[name].items(), key=lambda kv: (-kv[1], kv[0])):
            add(subtype, count, total, "All incidents", 1)

    summary.append(("", "", "", "", 0))
    summary.append(("DATA COMPLETENESS (fields left blank)", "", "", "", 0))
    operated = [r for r in present if r["_fields"]["operation"] == "OPERATED_ALERTED_OCCUPANT"]
    failed_rows = [r for r in present if r["_fields"]["operation"] == "FAILED_TO_OPERATE"]
    completeness = [
        ("Smoke Alarm Presence", rows, lambda r: r["_fields"]["presence"] != NO_DATA, "All incidents"),
        ("Smoke Alarm Working", present, lambda r: r["_fields"]["working"] != NO_DATA, "Present incidents"),
        ("Smoke Alarm Types", present, lambda r: bool(r["_fields"]["alarm_types"]), "Present incidents"),
        ("Smoke Alarm Operation", present, lambda r: r["_fields"]["operation"] != NO_DATA, "Present incidents"),
        ("Occupant Action", operated, lambda r: r["_fields"]["occupant_action"] != NO_DATA,
         "Operated - Alerted Occupant incidents"),
        ("Smoke Alarm Failure Reason", failed_rows, lambda r: r["_fields"]["failure_reason"] != NO_DATA,
         "Failed to Operate incidents"),
        ("Fire Alarm Presence", rows, lambda r: r["Fire Alarm Presence"] != NO_DATA, "All incidents"),
        ("Location Subtype", rows, lambda r: r["Location Subtype"] != NO_DATA, "All incidents"),
    ]
    for label, eligible, is_complete, basis in completeness:
        done = sum(1 for r in eligible if is_complete(r))
        summary.append((label, len(eligible), "", f"{basis} (eligible)", 0))
        add("Completed", done, len(eligible), basis, 1)
        add("Not completed", len(eligible) - done, len(eligible), basis, 1)
    # ---- By Location Group -------------------------------------------------
    def metrics(label, members, level):
        n = len(members)
        smoke_present = sum(1 for r in members if r["_fields"]["presence"] == "PRESENT")
        smoke_working = sum(1 for r in members
                            if r["_fields"]["presence"] == "PRESENT" and r["_fields"]["working"] == "Yes")
        fire_present = sum(1 for r in members if r["Fire Alarm Presence"] == "PRESENT")
        return (label, n, smoke_present, pct(smoke_present, n), smoke_working,
                pct(smoke_working, smoke_present), fire_present, pct(fire_present, n), level)

    group_rows = []
    for name in shown_groups:
        members = [r for r in rows if r["Location Group"] == name]
        group_rows.append(metrics(name, members, 0))
        for subtype, _ in sorted(subtype_counts[name].items(), key=lambda kv: (-kv[1], kv[0])):
            group_rows.append(metrics(
                subtype, [r for r in members if r["Location Subtype"] == subtype], 1))

    return summary, count_rows, rows, group_rows


ENTITY_METRIC_COLUMNS = [
    "Residential Structure Fires",
    "Smoke Alarm Present",
    "Present and Working",
    "Present and Not Working",
    "Present - Working Not Reported",
    "Smoke Alarm Not Present",
    "Smoke Alarm Not Applicable",
    "Smoke Alarm Presence Not Reported",
    "Operated - Alerted Occupant",
    "Operated - Failed to Alert Occupant",
    "Failed to Operate",
    "Fire Alarm Present",
]


def entity_metric_values(members):
    """Count the ENTITY_METRIC_COLUMNS measures for one set of raw incident rows."""
    def count(test):
        return sum(1 for r in members if test(r["_fields"], r))

    present = lambda f, r: f["presence"] == "PRESENT"
    return dict(zip(ENTITY_METRIC_COLUMNS, [
        len(members),
        count(present),
        count(lambda f, r: present(f, r) and f["working"] == "Yes"),
        count(lambda f, r: present(f, r) and f["working"] == "No"),
        count(lambda f, r: present(f, r) and f["working"] == NO_DATA),
        count(lambda f, r: f["presence"] == "NOT_PRESENT"),
        count(lambda f, r: f["presence"] == "NOT_APPLICABLE"),
        count(lambda f, r: f["presence"] == NO_DATA),
        count(lambda f, r: present(f, r) and f["operation"] == "OPERATED_ALERTED_OCCUPANT"),
        count(lambda f, r: present(f, r) and f["operation"] == "OPERATED_FAILED_TO_ALERT_OCCUPANT"),
        count(lambda f, r: present(f, r) and f["operation"] == "FAILED_TO_OPERATE"),
        count(lambda f, r: r["Fire Alarm Presence"] == "PRESENT"),
    ]))


def lookup_entity_names(client, entity_ids):
    names = {}
    for entity_id in sorted(entity_ids):
        try:
            names[entity_id] = (client.get_entity(entity_id) or {}).get("name", "")
        except Exception:
            names[entity_id] = ""
    return names


def write_entity_sheet(wb, raw_rows, entity_names):
    """Entity Smoke Alarm Reports: one heading + location-group table per entity,
    laid out like 'Entity Aid Reports' (AidActivityReport.py)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Entity Smoke Alarm Reports")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(*(Side(style="thin", color="B7C9D6") for _ in range(4)))
    heading_border = Border(*(Side(style="thin", color="000000") for _ in range(4)))
    heading_fill = PatternFill("solid", fgColor="D9EAF7")
    headers = ["Location Group"] + ENTITY_METRIC_COLUMNS
    count_format = '#,##0;(#,##0);-'

    def append_row(label, values):
        ws.append([label] + [values[c] for c in ENTITY_METRIC_COLUMNS])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, col)
            cell.border = border
            if col > 1:
                cell.number_format = count_format

    by_entity = {}
    for row in raw_rows:
        by_entity.setdefault(row["Department/Entity ID"], []).append(row)

    group_names = [name for name, _ in LOCATION_GROUPS] + [OTHER_GROUP]
    for index, entity_id in enumerate(sorted(by_entity)):
        if index:
            ws.append([])
            ws.append([])
        ws.append([entity_id, entity_names.get(entity_id, "")])
        heading_row = ws.max_row
        ws.merge_cells(start_row=heading_row, start_column=2, end_row=heading_row, end_column=5)
        ws.row_dimensions[heading_row].height = 48
        for col in range(1, 6):
            cell = ws.cell(heading_row, col)
            cell.fill = heading_fill
            cell.border = heading_border
            if col <= 2:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(
                    vertical="center", horizontal="left" if col == 2 else None, wrap_text=True)

        ws.append(headers)
        for cell in ws[ws.max_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[ws.max_row].height = 48

        members = by_entity[entity_id]
        for name in group_names:
            group_members = [r for r in members if r["Location Group"] == name]
            if group_members:
                append_row(name, entity_metric_values(group_members))
        append_row("Total", entity_metric_values(members))
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="E2F0D9")

    if not by_entity:
        ws.append(["No residential structure fire incidents found."])
    for col in range(1, len(headers) + 1):
        if col == 2:
            continue
        longest = 0
        for (value,) in ws.iter_rows(min_col=col, max_col=col, values_only=True):
            if isinstance(value, (int, float)):
                longest = max(longest, len(f"{value:,}"))
            elif value is not None and value not in headers:
                longest = max(longest, len(str(value)))
        ws.column_dimensions[get_column_letter(col)].width = max(10, min(longest + 2, 40))
    ws.column_dimensions["B"].width = 14


def export_to_excel(incidents, state_code, entity_scope, entity_id, entity_prefix,
                    entity_names, range_type, client=None):
    from report_metadata import add_report_sources_sheet, incident_date_range
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    summary_rows, count_rows, raw_rows, group_rows = build_tables(incidents)
    department_ids = {r["Department/Entity ID"] for r in raw_rows} - {"", None}
    department_names = lookup_entity_names(client, department_ids) if client else {}
    entity_sheet_rows = [dict(r) for r in raw_rows]
    for row in raw_rows:
        row.pop("_fields")

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")

    def write_sheet(wb, name, headers, data, pct_columns=(), levels=None, indent_column=1):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for record in data:
            ws.append(list(record))
        for column in pct_columns:
            for cell in ws[get_column_letter(column)][1:]:
                cell.number_format = "0.0%"
        if levels:
            for row_number, level in enumerate(levels, start=2):
                fill_color, bold = LEVEL_STYLES.get(level, LEVEL_STYLES[2])
                for cell in ws[row_number]:
                    cell.fill = PatternFill("solid", fgColor=fill_color)
                    cell.font = Font(bold=bold)
                label_cell = ws.cell(row=row_number, column=indent_column)
                label_cell.alignment = Alignment(indent=level * 2)
        for idx, header in enumerate(headers, 1):
            longest = max(
                [len(str(header))] + [len(str(ws.cell(row=r, column=idx).value or ""))
                                      + (4 * 2 if idx == indent_column else 0)
                                      for r in range(2, ws.max_row + 1)]
            )
            ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 12), 70)
        ws.freeze_panes = "A2"

    wb = Workbook()
    wb.remove(wb.active)
    write_entity_sheet(wb, entity_sheet_rows, department_names)
    write_sheet(
        wb, "Summary", ["Measure", "Incidents", "Percent", "Percent Basis"],
        [r[:4] for r in summary_rows], pct_columns=(3,),
        levels=[r[4] for r in summary_rows],
    )
    write_sheet(
        wb, "Individual Counts", ["Field", "Value", "Incidents", "% of Incidents"],
        [r[:4] for r in count_rows], pct_columns=(4,),
        levels=[r[4] for r in count_rows], indent_column=2,
    )
    write_sheet(
        wb, "By Location Group",
        ["Location Group", "Incidents", "Smoke Alarm Present", "% Present",
         "Present and Working", "% of Present Working", "Fire Alarm Present", "% Fire Alarm Present"],
        [r[:8] for r in group_rows], pct_columns=(4, 6, 8),
        levels=[r[8] for r in group_rows],
    )
    raw_headers = list(raw_rows[0].keys()) if raw_rows else ["Incident NERIS ID"]
    write_sheet(wb, "Raw Data", raw_headers, [r.values() for r in raw_rows])

    source_details = [
        ("State", state_code),
        ("Entity scope", ENTITY_SCOPE_LABELS.get(entity_scope, entity_scope)),
        ("Entity ID", entity_id or ""),
        ("Entities requested", " || ".join(entity_names)),
        ("County prefix", entity_prefix or ""),
        ("Date range", range_type),
        ("Incident filter", f"Structure Fire at {LOCATION_USE_CATEGORY.title()} locations"),
        ("Hierarchy", "Indented, color-coded rows are parts of the row above; the Percent Basis column states each denominator"),
        ("Included sheets", "Entity Smoke Alarm Reports, Summary, Individual Counts, By Location Group, Raw Data"),
    ]
    add_report_sources_sheet(wb, "IncidentPull_SmokeAlarms.py", incident_date_range(incidents), source_details)

    docs_dir = pathlib.Path.home() / "Documents"
    docs_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    filename = docs_dir / f"NERIS Smoke Alarm Incidents {state_code} {stamp}.xlsx"
    wb.save(filename)
    print(f"\n✓ Exported: {filename}")
    print(f"  Incidents : {len(incidents)}")
    return str(filename)


def main():
    ensure_dependencies()

    (
        state_code, entity_id, entity_prefix, entity_names,
        entity_scope, range_type, start_date, end_date,
    ) = prompt_config()
    if not state_code:
        sys.exit("✗ State code is required.")

    client = authenticate_client_credentials()
    start_filter, end_filter = calculate_date_range(range_type, start_date, end_date)

    incidents = get_scoped_incidents(
        client,
        state_code,
        entity_id=entity_id,
        entity_prefix=entity_prefix,
        entity_names=entity_names,
        entity_scope=entity_scope,
        call_create_start=start_filter,
        call_create_end=end_filter,
    )

    total_before = len(incidents)
    incidents = [inc for inc in incidents if is_residential_structure_fire(inc)]
    print(f"\nFiltered incidents: {len(incidents)} of {total_before} are residential Structure Fire incidents.")
    if not incidents:
        print("⚠ No incidents found for the given parameters.")
        return

    export_to_excel(
        incidents, state_code, entity_scope, entity_id, entity_prefix,
        entity_names, range_type, client=client,
    )

    print("\n" + "=" * 60)
    print("  PROCESS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
