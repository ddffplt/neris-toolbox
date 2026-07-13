"""
NERIS Lithium-ion Battery Fires Report Generator
------------------------------------------------
Fetches lithium-ion battery-related fire incidents from the NERIS API and exports
them to a formatted Excel file with:
1. Detailed incident list with battery-specific fields
2. Aggregate summary statistics
3. Report sources and metadata

Queries for incidents with:
- Outside Fire Cause: Battery_Power_storage
- Structure Fire Cause: Battery_Power_storage
- Emerging Hazards: elec_vehicle_status, elec_reignition, elec_suppress, elec_target
- Analysis Module: product_type, original_battery, installation_status, etc.
"""

import sys
import subprocess
import os
import copy
import re
import traceback
import builtins
import argparse
from datetime import datetime, timedelta, timezone
from collections import defaultdict

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

DATE_RANGE_OPTIONS = [
    "All Records",
    "Occurred Today",
    "Occurred Yesterday",
    "Occurred Last 7 Days",
    "Occurred Last 30 Days",
    "Occurred Last 90 Days",
    "Occurred in 2025",
    "Occurred in current calendar year to date",
    "Custom Date Range",
]


def parse_cli_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--username', '-u')
    parser.add_argument('--password', '-p')
    parser.add_argument('--entity-id', '-e')
    parser.add_argument('--statewide', action='store_true', help='Run statewide (no entity filter), overrides --entity-id if provided')
    parser.add_argument('--national', action='store_true', help='Allow nationwide run (no state filter)')
    parser.add_argument('--mfa-code', help='One-time MFA code to supply non-interactively')
    parser.add_argument('--range-type', choices=DATE_RANGE_OPTIONS)
    parser.add_argument('--start-date', help='YYYY-MM-DD')
    parser.add_argument('--end-date', help='YYYY-MM-DD')
    parser.add_argument('--state', help='State code to restrict entity listing (e.g. CA)')
    parser.add_argument('--max-incidents', type=int)
    parser.add_argument('--page-size', type=int)
    parser.add_argument('--require-structure-battery', action='store_true', help='Require Structure Fire Cause: Battery_Power_storage')
    parser.add_argument('--require-outside-battery', action='store_true', help='Require Outside Fire Cause: Battery_Power_storage')
    return parser.parse_args()


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Lithium-ion Battery Fires Report")
    print("=" * 60)

    print("\n── Credentials ──────────────────────────────────")
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()

    print("\n── Query Parameters ─────────────────────────────")
    print("Enter a NERIS Entity ID (e.g. FD26163151) to fetch incidents for a single department.")
    print("Leave blank to fetch incidents for the entire state (no department filter).")
    entity_id_raw = input("NERIS Entity ID (or leave blank for statewide): ").strip()
    entity_id = entity_id_raw if entity_id_raw else None

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
            s = input("Start date (YYYY-MM-DD): ").strip()
            try:
                start_date = datetime.strptime(s, "%Y-%m-%d").date()
                break
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")
        while True:
            e = input("End date (YYYY-MM-DD): ").strip()
            try:
                end_date = datetime.strptime(e, "%Y-%m-%d").date()
                break
            except ValueError:
                print("  Invalid format, use YYYY-MM-DD.")

    return username, password, entity_id, range_type, start_date, end_date


def calculate_date_range(range_type, start_date=None, end_date=None):
    """Return (call_create_start, call_create_end) as UTC-aware datetimes."""
    now   = datetime.now(tz=timezone.utc)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if range_type == "All Records":
        return None, None
    elif range_type == "Occurred Today":
        return today, None
    elif range_type == "Occurred Yesterday":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday.replace(hour=23, minute=59, second=59)
    elif range_type == "Occurred Last 7 Days":
        return today - timedelta(days=7), None
    elif range_type == "Occurred Last 30 Days":
        return today - timedelta(days=30), None
    elif range_type == "Occurred Last 90 Days":
        return today - timedelta(days=90), None
    elif range_type == "Occurred in current calendar year to date":
        return datetime(now.year, 1, 1, 0, 0, 0, tzinfo=timezone.utc), None
    elif range_type.startswith("Occurred in "):
        year = int(range_type.split()[-1])
        return (
            datetime(year, 1,  1,  0,  0,  0, tzinfo=timezone.utc),
            datetime(year, 12, 31, 23, 59, 59, tzinfo=timezone.utc),
        )
    elif range_type == "Custom Date Range" and start_date and end_date:
        start_dt = datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc)
        end_dt   = datetime.combine(
            end_date, datetime.max.time().replace(microsecond=0), tzinfo=timezone.utc
        )
        return start_dt, end_dt

    return None, None

def authenticate(username, password, mfa_code=None):
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

    orig_input = None
    if mfa_code is not None:
        orig_input = builtins.input
        builtins.input = lambda prompt=None: mfa_code
    else:
        print("\nAbout to trigger the MFA challenge. When you are ready to receive the code, press Enter to continue.")
        try:
            builtins.input("")
        except Exception:
            input("")

    try:
        client.list_incidents(page_size=1)
    finally:
        if orig_input is not None:
            builtins.input = orig_input

    print("\n✓ Authentication successful!")
    return client


def get_department_incidents(client, neris_id_entity,
                              call_create_start=None, call_create_end=None,
                              max_incidents=None, page_size=100, state=None):
    all_incidents = []
    next_cursor   = None
    page_count    = 0

    if neris_id_entity:
        print(f"\nFetching incidents for entity: {neris_id_entity}")
    else:
        print("\nFetching incidents for the entire state (no entity filter)")

    if call_create_start:
        print(f"  call_create >= {call_create_start:%Y-%m-%d %H:%M:%S}")
    if call_create_end:
        print(f"  call_create <= {call_create_end:%Y-%m-%d %H:%M:%S}")
    if not call_create_start and not call_create_end:
        print("  No date filter — retrieving all records")

    while True:
        page_count += 1
        print(f"  Page {page_count}... ", end="", flush=True)

        kwargs = {"page_size": page_size}
        if neris_id_entity:
            kwargs["neris_id_entity"] = neris_id_entity
        if state:
            kwargs["state"] = state
        if call_create_start: kwargs["call_create_start"] = call_create_start
        if call_create_end:   kwargs["call_create_end"]   = call_create_end
        if next_cursor:       kwargs["cursor"]             = next_cursor

        try:
            res = client.list_incidents(**kwargs)
            if not isinstance(res, dict):
                res = res.json()
        except Exception as e:
            print(f"\n✗ Error on page {page_count}: {e}")
            traceback.print_exc()
            break

        incidents = res.get("incidents", [])
        if not incidents:
            print("empty page — done.")
            break

        all_incidents.extend(incidents)
        print(f"{len(incidents)} fetched  (running total: {len(all_incidents)})")

        if max_incidents and len(all_incidents) >= max_incidents:
            all_incidents = all_incidents[:max_incidents]
            print(f"✓ Reached max limit of {max_incidents}")
            break

        next_cursor = res.get("next_cursor")
        if not next_cursor:
            print("  No more pages.")
            break

    print(f"\n{'='*50}")
    print(f"Total incidents retrieved: {len(all_incidents)}")
    print(f"{'='*50}")
    return all_incidents


def get_statewide_incidents(client, call_create_start=None, call_create_end=None, max_incidents=None, page_size=100, state=None):
    """Fetch statewide incidents by querying each registered entity."""
    print("\nAttempting statewide fetch by enumerating entities...")

    entities = []
    entity_ids = set()
    entity_prefix = "FD25" if state and state.upper() == "MA" else None
    page_number = 1
    try:
        while True:
            kwargs = {'page_size': page_size, 'page_number': page_number}
            if state:
                kwargs['state'] = state
            res = client.list_entities(**kwargs)
            if not isinstance(res, dict):
                res = res.json()
            batch = res.get('entities', [])
            if not batch:
                break
            new_entities = [
                entity for entity in batch
                if isinstance(entity, dict)
                and entity.get('neris_id')
                and (not entity_prefix or entity.get('neris_id').startswith(entity_prefix))
                and entity.get('neris_id') not in entity_ids
            ]
            if not new_entities:
                print("  No new entities returned — stopping pagination.")
                break
            entities.extend(new_entities)
            entity_ids.update(entity.get('neris_id') for entity in new_entities)
            print(f"  Retrieved {len(new_entities)} entities (total so far: {len(entities)})")
            page_number += 1
    except Exception as e:
        print(f"\n✗ Could not enumerate entities: {e}")
        traceback.print_exc()
        return []

    print(f"\nFound {len(entities)} entities to query")

    aggregated = []
    for ent in entities:
        eid = ent.get('neris_id') if isinstance(ent, dict) else None
        if not eid:
            continue
        print(f"\nQuerying entity {eid} ({ent.get('name','')}) ...")
        incs = get_department_incidents(client, eid, call_create_start=call_create_start, call_create_end=call_create_end, max_incidents=None, page_size=page_size)
        if incs:
            aggregated.extend(incs)
        if max_incidents and len(aggregated) >= max_incidents:
            aggregated = aggregated[:max_incidents]
            print(f"✓ Reached global max_incidents limit of {max_incidents}")
            break

    print(f"\nTotal aggregated incidents: {len(aggregated)}")
    return aggregated

def export_to_excel(incidents, client, department_id=None, range_type=None, start_date=None, end_date=None):
    from report_metadata import add_report_sources_sheet, incident_date_range
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    if not incidents:
        print("⚠ No incidents to export.")
        return None

    # Build department names lookup
    print("\nLooking up department names...")
    dept_ids = set()
    for inc in incidents:
        base = inc.get("base", {}) or {}
        dept_id = base.get("department_neris_id", "")
        if dept_id:
            dept_ids.add(dept_id)
    
    dept_names = {}
    for did in dept_ids:
        try:
            entity = client.get_entity(did)
            dept_names[did] = entity.get("name", "")
        except Exception:
            dept_names[did] = ""
    print(f"✓ Department lookup complete ({len(dept_names)} departments)")

    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    dept_label = department_id.replace("FD", "") if department_id else "statewide"
    filename   = f"lithium_ion_batteries_{dept_label}_{timestamp}.xlsx"

    def extract_list_field(lst, key):
        if not lst or not isinstance(lst, list):
            return ""
        return "||".join(
            str(item.get(key, ""))
            for item in lst
            if isinstance(item, dict) and item.get(key)
        )

    def join_list(lst):
        if not lst or not isinstance(lst, list):
            return ""
        return "||".join(str(x) for x in lst if x)

    def get_coordinates_from_point(point_obj):
        """Extract (lat, lon) from a GeoJSON point geometry object."""
        if not point_obj or not isinstance(point_obj, dict):
            return None, None
        geometry = point_obj.get("geometry", {})
        if geometry and isinstance(geometry, dict):
            coords = geometry.get("coordinates", [])
            if coords and len(coords) >= 2:
                return coords[1], coords[0]   # GeoJSON is [lon, lat]
        return None, None

    def get_location_data(loc_dict, point_obj=None):
        """Return a flat dict of address fields from a location object."""
        if not loc_dict or not isinstance(loc_dict, dict):
            return {}
        lat, lon = get_coordinates_from_point(point_obj)
        
        # Fallback to latitude/longitude directly in location dict if point didn't have coordinates
        if lat is None:
            lat = loc_dict.get("latitude", "")
        if lon is None:
            lon = loc_dict.get("longitude", "")
        
        return {
            "street_number": loc_dict.get("number", ""),
            "street_prefix": loc_dict.get("street_prefix_direction", ""),
            "street_name":   loc_dict.get("street", ""),
            "street_type":   loc_dict.get("street_postfix", ""),
            "street_suffix": loc_dict.get("street_postfix_direction", ""),
            "apartment":     loc_dict.get("unit_value", ""),
            "city":          loc_dict.get("incorporated_municipality", ""),
            "state":         loc_dict.get("state", ""),
            "zip_code":      loc_dict.get("postal_code", ""),
            "county":        loc_dict.get("county", ""),
            "latitude":      lat if lat != "" else "",
            "longitude":     lon if lon != "" else "",
        }

    def build_address_field(location_data):
        """Combine location fields into a single address, excluding bland/empty fields."""
        parts = []
        if location_data.get("street_number"):
            parts.append(str(location_data["street_number"]))
        if location_data.get("street_prefix"):
            parts.append(str(location_data["street_prefix"]))
        if location_data.get("street_name"):
            parts.append(str(location_data["street_name"]))
        if location_data.get("street_type"):
            parts.append(str(location_data["street_type"]))
        if location_data.get("street_suffix"):
            parts.append(str(location_data["street_suffix"]))
        if location_data.get("apartment"):
            parts.append(f"Apt {location_data['apartment']}")
        return " ".join(parts)

    ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

    def sanitize(val):
        return ILLEGAL_CHARS_RE.sub("", val) if isinstance(val, str) else val

    def extract_list_field(lst, key):
        """Extract a specific field from all items in a list and join with ||."""
        if not lst or not isinstance(lst, list):
            return ""
        return "||".join(
            str(item.get(key, ""))
            for item in lst
            if isinstance(item, dict) and item.get(key)
        )

    def is_fatal(injury_type):
        """Check if injury type is fatal."""
        value = (injury_type or "").upper()
        return "FATAL" in value and "NONFATAL" not in value

    def extract_injury_death_data(inc):
        """Extract civilian and firefighter injury/death counts from casualty_rescues."""
        injuries = {
            "Civilian Deaths": 0,
            "Civilian Injuries": 0,
            "Firefighter Deaths": 0,
            "Firefighter Injuries": 0,
        }
        
        for casualty_rescue in inc.get("casualty_rescues") or []:
            if not isinstance(casualty_rescue, dict):
                continue
            casualty_type = str(casualty_rescue.get("type") or "").upper()
            injury = (casualty_rescue.get("casualty") or {}).get("injury_or_noninjury") or {}
            injury_type = injury.get("type") or ""
            
            if casualty_type == "FF":
                prefix = "Firefighter "
            else:
                prefix = "Civilian "
            
            if is_fatal(injury_type):
                injuries[prefix + "Deaths"] += 1
            elif "NONFATAL" in str(injury_type).upper():
                injuries[prefix + "Injuries"] += 1
        
        return injuries

    detail_rows = []
    summary_stats = defaultdict(int)

    for inc in incidents:
        base       = inc.get("base", {}) or {}
        dispatch   = inc.get("dispatch", {}) or {}
        status_obj = inc.get("incident_status", {}) or {}
        fire_detail = inc.get("fire_detail", {}) or {}
        emerging_hazards = inc.get("emerging_hazards", {}) or {}
        electric_hazards = inc.get("electric_hazards", []) or []
        incident_analysis = inc.get("incident_analysis", {}) or {}

        inc_types = inc.get("incident_types", []) or []
        type_list = [
            it.get("type", "")
            for it in inc_types
            if isinstance(it, dict) and it.get("type")
        ]

        # Get location data
        base_loc       = base.get("location", {}) or {}
        base_geo       = base_loc.get("geocoded_location", {}) or {}
        base_geo_point = base_geo.get("point", {}) if base_geo else {}

        disp_loc       = dispatch.get("location", {}) or {}
        disp_geo       = disp_loc.get("geocoded_location", {}) or {}
        disp_geo_point = disp_geo.get("point", {}) if disp_geo else {}

        final_location  = {}
        if base_geo and any(base_geo.values()):
            final_location  = get_location_data(base_geo, base_geo_point)
        elif base_loc and any(base_loc.values()):
            final_location  = get_location_data(base_loc)
        elif disp_geo and any(disp_geo.values()):
            final_location  = get_location_data(disp_geo, disp_geo_point)
        elif disp_loc and any(disp_loc.values()):
            final_location  = get_location_data(disp_loc)

        # Location use - split into Primary and Secondary
        location_use_str = base.get("location_use", "")
        if isinstance(location_use_str, dict):
            # If it's a dict, extract the use_type and use_class
            location_use_str = location_use_str.get("use_type", "")
        location_use_str = str(location_use_str) if location_use_str else ""
        
        # Split by || delimiter
        if "||" in location_use_str:
            use_parts = location_use_str.split("||")
            primary_location_use = use_parts[0].strip() if len(use_parts) > 0 else ""
            secondary_location_use = use_parts[1].strip() if len(use_parts) > 1 else ""
        else:
            primary_location_use = location_use_str
            secondary_location_use = ""

        # Split incident types into primary, secondary, tertiary
        primary_incident_type = type_list[0] if len(type_list) > 0 else ""
        secondary_incident_type = type_list[1] if len(type_list) > 1 else ""
        tertiary_incident_type = type_list[2] if len(type_list) > 2 else ""

        # Get call_arrival for date/time
        call_arrival_utc = dispatch.get("call_arrival", "")
        
        # Get battery and lithium-ion fields from appropriate modules
        structure_fire_cause = fire_detail.get("structure_fire_cause", {}) or {}
        outside_fire_cause = fire_detail.get("outside_fire_cause", {}) or {}
        
        # Extract injury/death data
        injury_death_data = extract_injury_death_data(inc)
        
        row = {
            "Entity ID": base.get("department_neris_id", ""),
            "Department Name": dept_names.get(base.get("department_neris_id", ""), ""),
            "Incident Date/Time (UTC)": call_arrival_utc,
            "Incident Date": "",  # Will be calculated in Excel
            "Incident Time": "",  # Will be calculated in Excel
            "Incident Month": "",  # Will be calculated in Excel
            "Incident Year": "",  # Will be calculated in Excel
            "Incident Hour": "",  # Will be calculated in Excel
            "Incident Number": base.get("incident_number", ""),
            "Primary Incident Type": primary_incident_type,
            "Secondary Incident Type": secondary_incident_type,
            "Tertiary Incident Type": tertiary_incident_type,
            "Address": build_address_field(final_location),
            "City": str(final_location.get("city", "")) if final_location.get("city") else "",
            "State": str(final_location.get("state", "")) if final_location.get("state") else "",
            "Zip Code": str(final_location.get("zip_code", "")) if final_location.get("zip_code") else "",
            "County": str(final_location.get("county", "")) if final_location.get("county") else "",
            "Primary Location Use": primary_location_use,
            "Secondary Location Use": secondary_location_use,
            "Latitude": str(final_location.get("latitude", "")) if final_location.get("latitude") != "" else "",
            "Longitude": str(final_location.get("longitude", "")) if final_location.get("longitude") != "" else "",
            "NERIS ID": inc.get("neris_id", ""),
            # Fire Cause - Battery Power Storage
            "Structure Fire Cause - Battery_Power_storage": structure_fire_cause.get("Battery_Power_storage", ""),
            "Outside Fire Cause - Battery_Power_storage": outside_fire_cause.get("Battery_Power_storage", ""),
            # Emerging Hazards - Electric Vehicle
            "Electric Vehicle Status": emerging_hazards.get("elec_vehicle_status", ""),
            "Electric Reignition": emerging_hazards.get("elec_reignition", ""),
            "Electric Suppression": emerging_hazards.get("elec_suppress", ""),
            "Electric Target": emerging_hazards.get("elec_target", ""),
            # Electric Hazards - All available fields
            "Electric Hazard Types": extract_list_field(electric_hazards, "type"),
            "Electric Hazard Source/Target": extract_list_field(electric_hazards, "source_or_target"),
            "Electric Hazard Electrification Hazard": extract_list_field(electric_hazards, "electrification_hazard"),
            "Electric Hazard Battery Powered/Stored Energy Category": extract_list_field(electric_hazards, "battery_powered_stored_energy_category"),
            "Electric Hazard Specific Type": extract_list_field(electric_hazards, "specific_type"),
            # Incident Analysis - Battery/Lithium-ion fields
            "Product Type": incident_analysis.get("product_type", ""),
            "Original Battery": incident_analysis.get("original_battery", ""),
            "Installation Status": incident_analysis.get("installation_status", ""),
            "Safety Listed": incident_analysis.get("safety_listed", ""),
            "Vehicle Battery": incident_analysis.get("vehicle_battery", ""),
            "Charging at Ignition": incident_analysis.get("charging_at_ignition", ""),
            "Incident Indoor/Outdoor": incident_analysis.get("Incident_indoor_outdoor", ""),
            "Impacted from Source": incident_analysis.get("impacted_from_source", ""),
            "Thermal Runaway": incident_analysis.get("thermal_runaway", ""),
            "FD Suppression": incident_analysis.get("fd_suppression", ""),
            "Battery Cell": incident_analysis.get("battery_cell", ""),
            "Battery Chemistry": incident_analysis.get("battery_chemistry", ""),
            # Injury and Death Data
            "Civilian Deaths": injury_death_data.get("Civilian Deaths", 0),
            "Civilian Injuries": injury_death_data.get("Civilian Injuries", 0),
            "Firefighter Deaths": injury_death_data.get("Firefighter Deaths", 0),
            "Firefighter Injuries": injury_death_data.get("Firefighter Injuries", 0),
        }
        
        detail_rows.append(row)
        
        # Build summary statistics
        if fire_detail.get("structure_fire_cause", {}).get("Battery_Power_storage") or \
           fire_detail.get("outside_fire_cause", {}).get("Battery_Power_storage"):
            summary_stats["Total Lithium-ion Related Fires"] += 1
        
        if emerging_hazards.get("elec_vehicle_status"):
            summary_stats["Electric Vehicle Related"] += 1
        
        if electric_hazards:
            summary_stats["Incidents with Electric Hazards"] += 1
            
        if incident_analysis.get("thermal_runaway"):
            summary_stats["Thermal Runaway Events"] += 1

    df_detail = pd.DataFrame(detail_rows).fillna("")

    # Create Summary statistics sheet data
    summary_rows = [
        {"Metric": "Total Incidents Reported", "Count": len(incidents)},
        {"Metric": "Battery-Related Structure Fires", "Count": sum(1 for r in detail_rows if r.get("Structure Fire Cause - Battery_Power_storage"))},
        {"Metric": "Battery-Related Outside Fires", "Count": sum(1 for r in detail_rows if r.get("Outside Fire Cause - Battery_Power_storage"))},
        {"Metric": "Electric Vehicle Related", "Count": summary_stats["Electric Vehicle Related"]},
        {"Metric": "Incidents with Electric Hazards", "Count": summary_stats["Incidents with Electric Hazards"]},
        {"Metric": "Thermal Runaway Events", "Count": summary_stats["Thermal Runaway Events"]},
        {"Metric": "Incidents with Original Battery Info", "Count": sum(1 for r in detail_rows if r.get("Original Battery"))},
        {"Metric": "Incidents with Battery Chemistry Info", "Count": sum(1 for r in detail_rows if r.get("Battery Chemistry"))},
    ]

    df_summary = pd.DataFrame(summary_rows)

    print(f"  Incident Details: {len(df_detail.columns)} columns, {len(df_detail)} rows")
    print(f"  Summary Stats:    {len(df_summary.columns)} columns, {len(df_summary)} rows")

    def write_sheet(wb, df, name, is_summary=False):
        ws      = wb.create_sheet(name)
        headers = list(df.columns)
        
        utc_col_idx = None
        eastern_col_idx = None
        if "Incident Date/Time (UTC)" in headers and not is_summary:
            utc_col_idx = headers.index("Incident Date/Time (UTC)") + 1
            if "Incident Date" in headers:
                eastern_col_idx = headers.index("Incident Date") + 1

        ws.append(headers)
        
        for idx, (_, row) in enumerate(df.iterrows()):
            row_data = [sanitize(row[h]) if h in row else "" for h in headers]
            ws.append(row_data)
            
            if utc_col_idx is not None and eastern_col_idx is not None:
                row_num = ws.max_row
                utc_cell = f"{get_column_letter(utc_col_idx)}{row_num}"
                eastern_cell = f"{get_column_letter(eastern_col_idx)}{row_num}"
                
                # First, create formula for eastern datetime conversion
                ws.cell(row_num, eastern_col_idx).value = (
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
                ws.cell(row_num, eastern_col_idx).number_format = "mm/dd/yy hh:mm"
                
                # Now derive other fields from the eastern datetime
                if "Incident Time" in headers:
                    time_col_idx = headers.index("Incident Time") + 1
                    ws.cell(row_num, time_col_idx).value = f'=IF({eastern_cell}="","",MOD({eastern_cell},1))'
                    ws.cell(row_num, time_col_idx).number_format = "hh:mm"
                
                if "Incident Month" in headers:
                    month_col_idx = headers.index("Incident Month") + 1
                    ws.cell(row_num, month_col_idx).value = f'=IF({eastern_cell}="","",TEXT({eastern_cell},"mmm"))'
                
                if "Incident Year" in headers:
                    year_col_idx = headers.index("Incident Year") + 1
                    ws.cell(row_num, year_col_idx).value = f'=IF({eastern_cell}="","",YEAR({eastern_cell}))'
                
                if "Incident Hour" in headers:
                    hour_col_idx = headers.index("Incident Hour") + 1
                    ws.cell(row_num, hour_col_idx).value = f'=IF({eastern_cell}="","",HOUR({eastern_cell}))'

        fill  = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font  = Font(color="FFFFFF", bold=True, size=11)
        align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.fill      = fill
            cell.font      = font
            cell.alignment = align

        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            max_len    = len(str(header))
            for row_num in range(2, min(ws.max_row + 1, 52)):
                val = ws.cell(row_num, col_idx).value
                if val:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 60)

        ws.freeze_panes = "A2"

    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.remove(wb.active)
    
    write_sheet(wb, df_detail, "Incident Details")
    write_sheet(wb, df_summary, "Summary", is_summary=True)

    # Build Report Sources details
    date_range_label = f"{range_type}"
    if range_type == "Custom Date Range" and start_date and end_date:
        date_range_label += f" ({start_date} to {end_date})"
    
    add_report_sources_sheet(
        wb,
        "IncidentPull_LithiumIonBatteries.py",
        incident_date_range(incidents),
        [
            ("Report Type", "Lithium-ion Battery Fires"),
            ("Date Range", date_range_label),
            ("Entity Filter", department_id or "Statewide"),
            ("Included sheets", "Incident Details, Summary, Report Sources"),
        ],
    )
    
    wb.save(filename)

    print(f"\n✓ Exported: {filename}")
    print(f"  Incidents : {len(incidents)}")
    print(f"  Sheets    : Incident Details, Summary, Report Sources")
    return filename

def incident_matches_lithium_filter(inc, require_structure_battery=False, require_outside_battery=False):
    """Check if incident matches lithium-ion battery fire criteria.
    
    Args:
        inc: incident object
        require_structure_battery: if True, must have Structure Fire Cause: Battery_Power_storage
        require_outside_battery: if True, must have Outside Fire Cause: Battery_Power_storage
    """
    # Check fire causes
    fire_detail = inc.get("fire_detail", {}) or {}
    structure_cause = fire_detail.get("structure_fire_cause", {}) or {}
    outside_cause = fire_detail.get("outside_fire_cause", {}) or {}
    
    has_structure_battery = bool(structure_cause.get("Battery_Power_storage"))
    has_outside_battery = bool(outside_cause.get("Battery_Power_storage"))
    
    # If specific causes are required, check those constraints
    if require_structure_battery and not has_structure_battery:
        return False
    if require_outside_battery and not has_outside_battery:
        return False
    
    # If no specific requirements, check for any battery cause
    if not require_structure_battery and not require_outside_battery:
        has_battery_cause = has_structure_battery or has_outside_battery
    else:
        has_battery_cause = True  # Already passed the required checks above
    
    # Check emerging hazards or incident analysis for battery-related info
    emerging_hazards = inc.get("emerging_hazards", {}) or {}
    electric_hazards = inc.get("electric_hazards", []) or []
    incident_analysis = inc.get("incident_analysis", {}) or {}
    
    has_emerging = (
        emerging_hazards.get("elec_vehicle_status") or 
        emerging_hazards.get("elec_reignition") or
        emerging_hazards.get("elec_suppress") or
        emerging_hazards.get("elec_target")
    )
    
    has_electric_hazards = bool(electric_hazards)
    
    has_analysis = (
        incident_analysis.get("product_type") or
        incident_analysis.get("original_battery") or
        incident_analysis.get("battery_chemistry") or
        incident_analysis.get("thermal_runaway")
    )
    
    return has_battery_cause or has_emerging or has_electric_hazards or has_analysis

def main():
    ensure_dependencies()

    args = parse_cli_args()
    cli_used = any([
        args.username, args.password, args.range_type, args.start_date,
        args.end_date, args.entity_id, args.statewide, args.max_incidents, args.page_size
    ])

    if cli_used:
        username = args.username or os.environ.get("NERIS_USERNAME")
        password = args.password or os.environ.get("NERIS_PASSWORD")
        entity_id = None if args.statewide else (args.entity_id or None)
        range_type = args.range_type or "All Records"
        start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date() if args.start_date else None
        end_date = datetime.strptime(args.end_date, "%Y-%m-%d").date() if args.end_date else None
        max_incidents = args.max_incidents
        page_size = args.page_size or 100
    else:
        username, password, entity_id, range_type, start_date, end_date = prompt_config()
        max_incidents = None
        page_size = 100

    if not username or not password:
        sys.exit("✗ Email and password are required.")
    if not entity_id:
        print("No NERIS Entity ID provided — running statewide (no entity filter).")

    mfa_code = args.mfa_code if 'args' in locals() else None
    client = authenticate(username, password, mfa_code=mfa_code)

    start_filter, end_filter = calculate_date_range(range_type, start_date, end_date)

    if entity_id:
        incidents = get_department_incidents(
            client,
            entity_id,
            call_create_start=start_filter,
            call_create_end=end_filter,
            max_incidents=max_incidents,
            page_size=page_size,
        )
    else:
        state_param = None
        if 'args' in locals():
            if getattr(args, 'national', False):
                state_param = None
                print("--national flag provided — running nationwide (no state filter).")
            else:
                state_param = args.state

        if state_param is None and not ('args' in locals() and getattr(args, 'national', False)):
            state_param = 'MA'
            print("No state specified — defaulting to MA (Massachusetts) for statewide run.")

        incidents = get_statewide_incidents(
            client,
            call_create_start=start_filter,
            call_create_end=end_filter,
            max_incidents=max_incidents,
            page_size=page_size,
            state=state_param,
        )

    # Get fire cause filter requirements from CLI args
    require_structure = getattr(args, 'require_structure_battery', False) if 'args' in locals() else False
    require_outside = getattr(args, 'require_outside_battery', False) if 'args' in locals() else False
    
    # Filter to Lithium-ion battery related incidents
    total_before = len(incidents)
    filtered_incidents = [
        inc for inc in incidents 
        if incident_matches_lithium_filter(inc, require_structure_battery=require_structure, require_outside_battery=require_outside)
    ]
    incidents = filtered_incidents
    
    filter_desc = "lithium-ion battery fire criteria"
    if require_structure or require_outside:
        causes = []
        if require_structure:
            causes.append("Structure Fire Cause: Battery_Power_storage")
        if require_outside:
            causes.append("Outside Fire Cause: Battery_Power_storage")
        filter_desc = f"required fire causes: {', '.join(causes)}"
    
    print(f"\nFiltered incidents: {len(incidents)} of {total_before} match {filter_desc}.")

    if not incidents:
        print("⚠ No incidents found matching lithium-ion battery fire criteria.")
        return

    export_to_excel(incidents, client, department_id=entity_id, range_type=range_type, start_date=start_date, end_date=end_date)

    print("\n" + "=" * 60)
    print("  PROCESS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
