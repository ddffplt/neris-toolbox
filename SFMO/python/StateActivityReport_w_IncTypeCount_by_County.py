#!/usr/bin/env python3
"""
NERIS Monthly Incident Count Report
------------------------------------
Fetches all incidents and no-activity reports for every county in a
state and exports a pivot table (departments x months) to Excel. 
Also produces a summary of incident counts by fire-part incident type.

Usage:
    python neris_monthly_counts.py

Requirements (auto-installed if missing):
    pip install pandas openpyxl neris-api-client
"""

import sys
import subprocess
import os
from datetime import datetime, timedelta, timezone
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed


# ─────────────────────────────────────────────
# 1. DEPENDENCY INSTALLATION
# ─────────────────────────────────────────────

def ensure_dependencies():
    def pip_install(*packages):
        subprocess.run(
            [sys.executable, "-m", "pip", "install", *packages, "--quiet"],
            check=True
        )

    try:
        import pandas   # noqa
        import openpyxl # noqa
    except ImportError:
        print("Installing pandas and openpyxl...")
        pip_install("pandas", "openpyxl")
        print("✓ pandas and openpyxl installed")

    try:
        from neris_api_client import NerisApiClient  # noqa
        print("✓ Dependencies ready")
    except ImportError:
        print("Installing neris-api-client...")
        pip_install("neris-api-client")
        print("✓ neris-api-client installed")


# ─────────────────────────────────────────────
# 2. USER INPUT
# ─────────────────────────────────────────────

DATE_RANGE_OPTIONS = [
    "All Records",
    "Occurred Today",
    "Occurred Yesterday",
    "Occurred Last 7 Days",
    "Occurred Last 30 Days",
    "Occurred Last 90 Days",
    "Occurred in 2025",
    "Custom Date Range",
]

COUNTY_NAMES = {
    "FD25001": "Barnstable",
    "FD25003": "Berkshire",
    "FD25005": "Bristol",
    "FD25007": "Dukes",
    "FD25009": "Essex",
    "FD25011": "Franklin",
    "FD25013": "Hampden",
    "FD25015": "Hampshire",
    "FD25017": "Middlesex",
    "FD25019": "Nantucket",
    "FD25021": "Norfolk",
    "FD25023": "Plymouth",
    "FD25025": "Suffolk",
    "FD25027": "Worcester",
}

def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Monthly Incident Count Report")
    print("=" * 60)
    print("\nNote: This report covers all data from 1 Jan 2025 onward.")
    print("      NAR = No-Activity Report (confirmed zero incidents).")
    print("      Blank = no data of either kind submitted.\n")

    print("-- Credentials --")
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()

    print("\n-- Query Parameters --")
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()
    entity_id  = input("NERIS Entity ID (optional — leave blank for all departments): ").strip() or None

    print("\nDate Range Options:")
    for i, option in enumerate(DATE_RANGE_OPTIONS, 1):
        print(f"  {i}. {option}")
    while True:
        choice = input("Select option [1]: ").strip() or "1"
        if choice.isdigit() and 1 <= int(choice) <= len(DATE_RANGE_OPTIONS):
            range_type = DATE_RANGE_OPTIONS[int(choice) - 1]
            break
        print("  Invalid selection, try again.")

    start_date = end_date = None
    if range_type == "Custom Date Range":
        while True:
            value = input("Start date (YYYY-MM-DD): ").strip()
            try:
                start_date = datetime.strptime(value, "%Y-%m-%d").date()
                break
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")
        while True:
            value = input("End date (YYYY-MM-DD): ").strip()
            try:
                end_date = datetime.strptime(value, "%Y-%m-%d").date()
                if end_date < start_date:
                    print("  End date must be on or after the start date.")
                    continue
                break
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")

    if not username or not password or not state_code:
        sys.exit("✗ Email, password, and state code are required.")

    return username, password, state_code, entity_id, range_type, start_date, end_date


def calculate_date_range(range_type, start_date=None, end_date=None):
    """Return UTC-aware call_create bounds for the selected date range."""
    now = datetime.now(tz=timezone.utc)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if range_type == "All Records":
        return None, None
    if range_type == "Occurred Today":
        return today, None
    if range_type == "Occurred Yesterday":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday.replace(hour=23, minute=59, second=59)
    if range_type == "Occurred Last 7 Days":
        return today - timedelta(days=7), None
    if range_type == "Occurred Last 30 Days":
        return today - timedelta(days=30), None
    if range_type == "Occurred Last 90 Days":
        return today - timedelta(days=90), None
    if range_type.startswith("Occurred in "):
        year = int(range_type.split()[-1])
        return (
            datetime(year, 1, 1, tzinfo=timezone.utc),
            datetime(year, 12, 31, 23, 59, 59, tzinfo=timezone.utc),
        )
    if range_type == "Custom Date Range" and start_date and end_date:
        return (
            datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc),
            datetime.combine(end_date, datetime.max.time().replace(microsecond=0), tzinfo=timezone.utc),
        )
    return None, None


# ─────────────────────────────────────────────
# 3. AUTHENTICATION
# ─────────────────────────────────────────────

def authenticate(username, password):
    os.environ["NERIS_BASE_URL"]   = "https://api.neris.fsri.org/v1"
    os.environ["NERIS_GRANT_TYPE"] = "password"
    os.environ["NERIS_USERNAME"]   = username
    os.environ["NERIS_PASSWORD"]   = password

    from neris_api_client import NerisApiClient

    print("\nConnecting to NERIS API...")
    client = NerisApiClient()

    print("\n" + "=" * 60)
    print("  CHECK YOUR EMAIL FOR THE MFA CODE")
    print("=" * 60)

    client.list_incidents(page_size=1)

    print("\n✓ Authentication successful!")
    return client


# ─────────────────────────────────────────────
# 4. DATA FETCHING
# ─────────────────────────────────────────────

def fetch_all_entities(client, state_code, neris_id_entity=None, page_size=100):
    """
    Fetch every registered department/entity for the state.
    Returns a dict: { neris_id: name }
    If neris_id_entity is supplied, returns only that one department.
    """
    if neris_id_entity:
        try:
            entity = client.get_entity(neris_id_entity)
            name   = entity.get("name", "") if isinstance(entity, dict) else ""
            print(f"✓ Single entity lookup: {neris_id_entity} → {name}")
            return {neris_id_entity: name}
        except Exception as e:
            print(f"⚠ Could not fetch entity {neris_id_entity}: {e}")
            return {neris_id_entity: ""}

    all_entities = {}
    page_number  = 1

    print(f"\nFetching all registered entities for state: {state_code}")
    while True:
        print(f"  Page {page_number}... ", end="", flush=True)
        try:
            res = client.list_entities(state=state_code, page_size=page_size,
                                       page_number=page_number)
            if not isinstance(res, dict):
                res = res.json()
        except Exception as e:
            print(f"\n⚠ list_entities failed (page {page_number}): {e}")
            break

        batch = res.get("entities", [])
        if not batch:
            print("empty — done.")
            break

        for ent in batch:
            eid  = ent.get("neris_id", "")
            name = ent.get("name", "")
            if eid:
                all_entities[eid] = name

        print(f"retrieved {len(batch)} (total so far: {len(all_entities)})")

        total_count = res.get("total_count", 0)
        if len(all_entities) >= total_count or len(batch) < page_size:
            print("  ✓ All pages retrieved")
            break

        page_number += 1

    print(f"\n{'='*50}")
    print(f"Total registered departments: {len(all_entities)}")
    print(f"{'='*50}")
    return all_entities


def _month_label(dt_or_str):
    """Convert a datetime or 'MM/YYYY' / ISO string to 'Mon-YYYY'. Returns None on failure."""
    if dt_or_str is None:
        return None
    if isinstance(dt_or_str, str):
        if "/" in dt_or_str and len(dt_or_str) <= 7:
            try:
                m, y = dt_or_str.split("/")
                return datetime(int(y), int(m), 1).strftime("%b-%Y")
            except Exception:
                pass
        try:
            dt_or_str = datetime.fromisoformat(dt_or_str.replace("Z", "+00:00"))
        except Exception:
            return None
    try:
        return dt_or_str.strftime("%b-%Y")
    except Exception:
        return None


def fetch_incidents_for_entity(client, state_code, neris_id_entity,
                               call_create_start=None, call_create_end=None,
                               page_size=100):
    """
    Fetch all incidents for a single department.
    Returns month counts and incident type counts for the department.
    """
    month_counts = defaultdict(int)
    type_counts = defaultdict(int)
    total_incidents = 0
    multiple_type_incidents = 0
    next_cursor  = None

    while True:
        kwargs = dict(state=state_code, neris_id_entity=neris_id_entity,
                      page_size=page_size)
        if call_create_start:
            kwargs["call_create_start"] = call_create_start
        if call_create_end:
            kwargs["call_create_end"] = call_create_end
        if next_cursor:
            kwargs["cursor"] = next_cursor

        res = client.list_incidents(**kwargs)
        if not isinstance(res, dict):
            res = res.json()

        batch = res.get("incidents", [])
        if not batch:
            break

        for inc in batch:
            total_incidents += 1
            disp = inc.get("dispatch") or {}
            ts   = disp.get("call_create") or disp.get("call_create_start")
            lbl  = _month_label(ts)
            if lbl:
                month_counts[lbl] += 1

            fire_parts = {
                incident_type.get("type", "").split("||", 1)[0].strip()
                for incident_type in (inc.get("incident_types") or [])
                if isinstance(incident_type, dict)
                and incident_type.get("type")
                and incident_type.get("type").split("||", 1)[0].strip()
            }
            for fire_part in fire_parts:
                type_counts[fire_part] += 1
            if len(fire_parts) > 1:
                multiple_type_incidents += 1

        next_cursor = res.get("next_cursor")
        if not next_cursor:
            break

    return {
        "months": dict(month_counts),
        "types": dict(type_counts),
        "total": total_incidents,
        "multiple_types": multiple_type_incidents,
    }


def fetch_nars_for_entity(client, state_code, neris_id_entity, page_size=100):
    """
    Fetch all no-activity reports for a single department.
    Returns a set of 'Mon-YYYY' labels.
    """
    nar_months  = set()
    next_cursor = None
    base_url    = os.environ.get("NERIS_BASE_URL", "https://api.neris.fsri.org/v1")

    while True:
        params = dict(state=state_code, neris_id_entity=neris_id_entity,
                      page_size=page_size)
        if next_cursor:
            params["cursor"] = next_cursor

        r   = client._session.get(f"{base_url}/no_activity_report", params=params)
        res = r.json()

        batch = res.get("reports", [])
        if not batch:
            break

        for report in batch:
            lbl = _month_label(report.get("month_year", ""))
            if lbl:
                nar_months.add(lbl)

        next_cursor = res.get("next_cursor")
        if not next_cursor:
            break

    return nar_months


def generate_month_columns(call_create_start=None, call_create_end=None):
    """Return month labels intersecting the selected date range."""
    start = call_create_start or datetime(2025, 1, 1, tzinfo=timezone.utc)
    end = call_create_end or datetime.now(tz=timezone.utc)
    d = datetime(start.year, start.month, 1, tzinfo=timezone.utc)
    end_month = datetime(end.year, end.month, 1, tzinfo=timezone.utc)
    cols = []
    while d <= end_month:
        cols.append(d.strftime("%b-%Y"))
        m = d.month + 1
        d = datetime(d.year + (m // 13), ((m - 1) % 12) + 1, 1,
                     tzinfo=timezone.utc)
    return cols


def aggregate_county_results(county_names, all_entities, inc_counts,
                            type_counts, nar_flags):
    """Combine department results into one record per county."""
    county_inc_counts = {county_id: {} for county_id in county_names}
    county_type_counts = {
        county_id: {
            "name": county_names[county_id],
            "types": defaultdict(int),
            "total": 0,
            "multiple_types": 0,
        }
        for county_id in county_names
    }
    county_nar_flags = {county_id: set() for county_id in county_names}

    for department_id in all_entities:
        county_id = next(
            (
                county_id
                for county_id in county_names
                if department_id.upper().startswith(county_id)
            ),
            None,
        )
        if not county_id:
            continue

        for month, count in inc_counts.get(department_id, {}).items():
            county_inc_counts[county_id][month] = (
                county_inc_counts[county_id].get(month, 0) + count
            )

        department_data = type_counts.get(department_id, {})
        county_data = county_type_counts[county_id]
        for incident_type, count in department_data.get("types", {}).items():
            county_data["types"][incident_type] += count
        county_data["total"] += department_data.get("total", 0)
        county_data["multiple_types"] += department_data.get("multiple_types", 0)
        county_nar_flags[county_id].update(nar_flags.get(department_id, set()))

    return county_inc_counts, county_type_counts, county_nar_flags


# ─────────────────────────────────────────────
# 5. EXCEL EXPORT
# ─────────────────────────────────────────────

def export_to_excel(df_pivot, month_cols, state_code, inc_counts, nar_flags,
                    type_counts):
    from report_metadata import add_report_sources_sheet
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename  = f"neris_incident_type_counts_by_county_{state_code}_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Incident Type Summary"

    type_names = sorted({
        type_name
        for department_data in type_counts.values()
        for type_name in department_data["types"]
    })
    summary_headers = (
        ["County NERIS ID", "County Name"]
        + type_names
        + ["Total Incidents", "Multiple Incident Types"]
    )
    ws.append(summary_headers)

    header_cols = ["County NERIS ID", "County Name"] + month_cols

    hdr_fill     = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    nar_fill     = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_fill   = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    no_data_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    hdr_font     = Font(color="FFFFFF", bold=True, size=11)
    total_font   = Font(bold=True, size=11)
    hdr_aln      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ctr_aln      = Alignment(horizontal="center", vertical="center")
    left_aln     = Alignment(horizontal="left",   vertical="center")
    thin_side    = Side(style="thin", color="BFBFBF")
    thin_border  = Border(left=thin_side, right=thin_side,
                          top=thin_side,  bottom=thin_side)

    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = hdr_aln

    summary_rows = []
    for eid in sorted(type_counts):
        department_data = type_counts[eid]
        summary_rows.append(
            [eid, department_data["name"]]
            + [department_data["types"].get(type_name, 0) for type_name in type_names]
            + [department_data["total"], department_data["multiple_types"]]
        )
    summary_rows.append(
        ["TOTAL", ""]
        + [sum(row[2 + index] for row in summary_rows) for index in range(len(type_names))]
        + [sum(row[-2] for row in summary_rows), sum(row[-1] for row in summary_rows)]
    )

    total_row_idx = len(summary_rows) + 1
    rows = summary_rows
    for r_idx, row_data in enumerate(rows, start=2):
        is_total = (r_idx == total_row_idx)
        for c_idx, val in enumerate(row_data, start=1):
            cell        = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border

            if is_total:
                cell.fill      = total_fill
                cell.font      = total_font
                cell.alignment = ctr_aln
            elif c_idx <= 2:
                cell.alignment = left_aln
            elif val == "NAR":
                cell.fill      = nar_fill
                cell.font      = Font(italic=True, color="7F6000")
                cell.alignment = ctr_aln
            elif val == "":
                cell.fill      = no_data_fill
                cell.alignment = ctr_aln
            else:
                cell.alignment = ctr_aln

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 35
    for i in range(3, len(summary_headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11

    ws.freeze_panes = "C2"

    monthly_ws = wb.create_sheet("Monthly Incident Counts")
    monthly_ws.append(header_cols)
    for cell in monthly_ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = hdr_aln

    monthly_rows = df_pivot.values.tolist()
    for r_idx, row_data in enumerate(monthly_rows, start=2):
        is_total = (r_idx == len(monthly_rows) + 1)
        for c_idx, val in enumerate(row_data, start=1):
            cell = monthly_ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if is_total:
                cell.fill = total_fill
                cell.font = total_font
                cell.alignment = ctr_aln
            elif c_idx <= 2:
                cell.alignment = left_aln
            elif val == "NAR":
                cell.fill = nar_fill
                cell.font = Font(italic=True, color="7F6000")
                cell.alignment = ctr_aln
            elif val == "":
                cell.fill = no_data_fill
                cell.alignment = ctr_aln
            else:
                cell.alignment = ctr_aln

    monthly_ws.column_dimensions["A"].width = 20
    monthly_ws.column_dimensions["B"].width = 35
    for i in range(3, len(header_cols) + 1):
        monthly_ws.column_dimensions[get_column_letter(i)].width = 11
    monthly_ws.freeze_panes = "C2"

    # Legend sheet
    ws_legend  = wb.create_sheet("Legend")
    legend_data = [
        ("Symbol",       "Meaning"),
        ("(number)",     "Count of incidents for that department and fire-part incident type"),
        ("Total Incidents", "Total incidents reported for that department"),
        ("Multiple Incident Types", "Incidents with more than one fire-part incident type"),
        ("NAR",          "No-Activity Report filed — department confirmed zero incidents"),
        ("(grey blank)", "No incidents and no no-activity report submitted for that month"),
        ("TOTAL row",    "Sum of numeric incident counts across all departments per month"),
    ]
    for lr in legend_data:
        ws_legend.append(lr)
    for cell in ws_legend[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
    ws_legend.column_dimensions["A"].width = 14
    ws_legend.column_dimensions["B"].width = 75

    add_report_sources_sheet(
        wb,
        "StateActivityReport_w_IncTypeCount_by_County.py",
        f"{month_cols[0]} through {month_cols[-1]}",
        [("Monthly columns", f"{month_cols[0]} through {month_cols[-1]}"),
         ("Included sheets", "County incident type summary, monthly incident counts, Legend")],
    )

    wb.save(filename)
    return filename


# ─────────────────────────────────────────────
# 6. MAIN
# ─────────────────────────────────────────────

def main():
    ensure_dependencies()

    import pandas as pd

    (
        username, password, state_code, entity_id,
        range_type, start_date, end_date,
    ) = prompt_config()
    call_create_start, call_create_end = calculate_date_range(
        range_type, start_date, end_date
    )

    county_names = COUNTY_NAMES

    client = authenticate(username, password)

    # Master department list
    all_entities = fetch_all_entities(client, state_code, entity_id)
    all_entities = {
        department_id: name
        for department_id, name in all_entities.items()
        if any(
            department_id.upper().startswith(county_id)
            for county_id in county_names
        )
    }
    all_eids     = list(all_entities.keys())

    if not all_eids:
        sys.exit("✗ No fetched departments matched the embedded county prefixes.")

    print(f"\nDate range: {range_type}")
    if call_create_start:
        print(f"  call_create >= {call_create_start:%Y-%m-%d %H:%M:%S} UTC")
    if call_create_end:
        print(f"  call_create <= {call_create_end:%Y-%m-%d %H:%M:%S} UTC")
    print(f"\nFetching incidents and NARs for {len(all_eids)} departments in parallel...")
    print("(progress updates every 50 departments)\n")

    inc_counts = {}   # { eid: { 'Mon-YYYY': count } }
    type_counts = {}  # { eid: { name, types, total, multiple_types } }
    nar_flags  = {}   # { eid: { 'Mon-YYYY', ... } }
    completed  = 0
    errors     = []

    _first_dept_done = {"done": False}

    def fetch_dept_data(eid):
        incident_data = fetch_incidents_for_entity(
            client,
            state_code,
            eid,
            call_create_start=call_create_start,
            call_create_end=call_create_end,
        )
        nars   = fetch_nars_for_entity(client, state_code, eid)
        if not _first_dept_done["done"]:
            _first_dept_done["done"] = True
            print(f"\n  [Diagnostic] First dept {eid}: "
                  f"{incident_data['total']} incidents, {len(nars)} NAR months")
        return eid, incident_data, nars

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(fetch_dept_data, eid): eid for eid in all_eids}

        first_errors_shown = 0
        for future in as_completed(futures):
            eid = futures[future]
            try:
                eid, incident_data, nars = future.result()
                inc_counts[eid] = incident_data["months"]
                type_counts[eid] = {
                    "name": all_entities.get(eid, ""),
                    "types": incident_data["types"],
                    "total": incident_data["total"],
                    "multiple_types": incident_data["multiple_types"],
                }
                nar_flags[eid] = nars
            except Exception as e:
                errors.append((eid, str(e)))
                inc_counts[eid] = {}
                type_counts[eid] = {
                    "name": all_entities.get(eid, ""),
                    "types": {},
                    "total": 0,
                    "multiple_types": 0,
                }
                nar_flags[eid]  = set()
                if first_errors_shown < 3:
                    print(f"\n  ⚠ ERROR for {eid}: {e}")
                    first_errors_shown += 1

            completed += 1
            if completed % 50 == 0 or completed == len(all_eids):
                total_incidents = sum(sum(v.values()) for v in inc_counts.values())
                total_nars      = sum(len(v) for v in nar_flags.values())
                print(f"  {completed}/{len(all_eids)} departments complete "
                      f"| {total_incidents} incidents | {total_nars} NAR months")

    if errors:
        print(f"\n⚠ {len(errors)} department(s) had fetch errors:")
        for eid, err in errors[:10]:
            print(f"    {eid}: {err}")
        if len(errors) > 10:
            print(f"    ... and {len(errors) - 10} more")

    inc_counts, type_counts, nar_flags = aggregate_county_results(
        county_names,
        all_entities,
        inc_counts,
        type_counts,
        nar_flags,
    )
    all_entities = county_names
    all_eids = list(county_names.keys())

    total_incidents = sum(sum(v.values()) for v in inc_counts.values())
    total_nars      = sum(len(v) for v in nar_flags.values())
    print(f"\n{'='*50}")
    print(f"Total incidents retrieved : {total_incidents}")
    print(f"Total NAR months on record: {total_nars}")
    print(f"{'='*50}")

    # Build pivot DataFrame
    month_cols = generate_month_columns(call_create_start, call_create_end)
    rows = []

    for eid in sorted(all_eids):
        row = {
            "County NERIS ID": eid,
            "County Name": all_entities.get(eid, ""),
        }
        dept_counts = inc_counts.get(eid, {})
        dept_nars   = nar_flags.get(eid, set())

        for mc in month_cols:
            if mc in dept_counts:
                row[mc] = dept_counts[mc]
            elif mc in dept_nars:
                row[mc] = "NAR"
            else:
                row[mc] = ""
        rows.append(row)

    df_pivot = pd.DataFrame(rows)

    # Total summary row
    summary = {"County NERIS ID": "TOTAL", "County Name": ""}
    for mc in month_cols:
        nums = pd.to_numeric(df_pivot[mc], errors="coerce").dropna()
        summary[mc] = int(nums.sum()) if not nums.empty else ""
    df_pivot = pd.concat([df_pivot, pd.DataFrame([summary])], ignore_index=True)

    print(f"\nPivot table: {len(df_pivot)-1} departments x {len(month_cols)} months")

    filename = export_to_excel(
        df_pivot, month_cols, state_code, inc_counts, nar_flags, type_counts
    )

    print(f"\n✓ Exported: {filename}")
    print(f"  Counties    : {len(df_pivot)-1}")
    print(f"  Months      : {len(month_cols)}  ({month_cols[0]} - {month_cols[-1]})")
    print(f"  Incidents   : {total_incidents}")
    print(f"  NAR months  : {total_nars}")
    print("\nKey:")
    print("  (number)     = incident count")
    print("  NAR          = no-activity report filed (confirmed 0 incidents)")
    print("  (grey blank) = no data submitted for that month")

    print("\n" + "=" * 60)
    print("  PROCESS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
