"""
NERIS Apparatus Export
-----------------------
This script fetches all departments for a state, then retrieves each department's
stations and units, outputting a flat CSV with one row per apparatus unit. Stations 
without units are surfaced once with blank unit fields. 

You must have state-level entity set permissions (via an FM Node) to run this script successfully.  
"""
 
import csv
import os
import sys
import time
from datetime import datetime
 
STATE = "XX"
 
EXPORT_PATH = "neris_apparatus_{state}_{timestamp}.csv"

USERNAME = ""
PASSWORD = ""
 
 
OUTPUT_FILE = EXPORT_PATH.format(
    state=STATE.upper(),
    timestamp=datetime.now().strftime("%Y%m%d_%H%M%S"),
)
 
CSV_FIELDS = [
    "department_neris_id",
    "department_name",
    "department_city",
    "department_state",
    "department_type",
    "station_neris_id",
    "station_id",
    "station_city",
    "station_state",
    "unit_neris_id",
    "unit_type",
    "cad_designation_1",
    "cad_designation_2",
    "unit_staffing",
    "unit_dedicated_staffing",
]

#Authenticate and MFA 
def authenticate(username, password):
    from neris_api_client import NerisApiClient, Config
    print("\nConnecting to NERIS API...")
    client = NerisApiClient(Config(
        base_url="https://api.neris.fsri.org/v1",
        grant_type="password",
        username=username,
        password=password,
    ))
    print("\n✓ Authentication successful!")
    return client
 
 

def get_all_entities(client, state) -> list:
    entities = []
    page = 1
 
    while True:
        print(f"  Fetching entity list page {page}...", end=" ", flush=True)
        try:
            response = client.list_entities(state=state, page_number=page)
        except Exception as exc:
            print(f"\nERROR fetching entity list page {page}: {exc}")
            break
 
        batch = response.get("entities", [])
        entities.extend(batch)
 
        total    = response.get("total_count", 0)
        pg_size  = response.get("page_size", len(batch)) or 1
        pg_count = response.get("page_count") or (
            (total + pg_size - 1) // pg_size if total else 1
        )
 
        print(f"got {len(batch)} (total so far: {len(entities)} / {total})")
 
        if page >= pg_count or not batch:
            break
        page += 1
 
    return entities
 
 
def get_units_for_entity(client, entity_summary) -> list:
    dept_neris_id = entity_summary.get("neris_id", "")
    dept_name     = entity_summary.get("name", "")
    dept_city     = entity_summary.get("city", "")
    dept_state    = entity_summary.get("state", "")
    dept_type     = entity_summary.get("department_type", "")
 
    try:
        detail = client.get_entity(dept_neris_id)
    except Exception as exc:
        print(f"    WARNING: could not fetch detail for {dept_neris_id}: {exc}")
        return []
 
    rows = []
    for station in detail.get("stations", []):
        station_neris_id = station.get("neris_id", "")
        station_id       = station.get("station_id", "")
        station_city     = station.get("city", "")
        station_state    = station.get("state", "")
 
        units = station.get("units", [])
        #If a department doesn't have units, we'll still pull back their station information
        if not units:
            rows.append({
                "department_neris_id":    dept_neris_id,
                "department_name":        dept_name,
                "department_city":        dept_city,
                "department_state":       dept_state,
                "department_type":        dept_type,
                "station_neris_id":       station_neris_id,
                "station_id":             station_id,
                "station_city":           station_city,
                "station_state":          station_state,
                "unit_neris_id":          "",
                "unit_type":              "",
                "cad_designation_1":      "",
                "cad_designation_2":      "",
                "unit_staffing":          "",
                "unit_dedicated_staffing": "",
            })
        else:
            for unit in units:
                rows.append({
                    "department_neris_id":    dept_neris_id,
                    "department_name":        dept_name,
                    "department_city":        dept_city,
                    "department_state":       dept_state,
                    "department_type":        dept_type,
                    "station_neris_id":       station_neris_id,
                    "station_id":             station_id,
                    "station_city":           station_city,
                    "station_state":          station_state,
                    "unit_neris_id":          unit.get("neris_id", ""),
                    "unit_type":              unit.get("type", ""),
                    "cad_designation_1":      unit.get("cad_designation_1", ""),
                    "cad_designation_2":      unit.get("cad_designation_2", ""),
                    "unit_staffing":          unit.get("staffing", ""),
                    "unit_dedicated_staffing": unit.get("dedicated_staffing", ""),
                })
    return rows
 
def main():
    if not STATE or len(STATE.strip()) != 2:
        print("ERROR: STATE must be a two-letter abbreviation (e.g. 'IL'). Check the SETTINGS section.")
        sys.exit(1)
 
    state = STATE.strip().upper()
 
    username = USERNAME.strip() if USERNAME.strip() else input("NERIS username (email): ")
    password = PASSWORD.strip() if PASSWORD.strip() else input("NERIS password: ")
 
    export_dir = os.path.dirname(OUTPUT_FILE)
    if export_dir:
        os.makedirs(export_dir, exist_ok=True)
 
    client = authenticate(username, password)
 
    #Get all entities for the state
    print(f"\n-- Step 1: Fetching all {state} departments --")
    entities = get_all_entities(client, state)
    print(f"\nFound {len(entities)} {state} department(s).\n")
 
    if not entities:
        print("No entities returned -- check credentials.")
        sys.exit(1)
 
    print("-- Step 2: Fetching unit details for each department --")
    all_rows = []
    for i, entity in enumerate(entities, 1):
        neris_id = entity.get("neris_id", "?")
        name     = entity.get("name", "?")
        print(f"  [{i}/{len(entities)}] {neris_id}  {name}")
        rows = get_units_for_entity(client, entity)
        all_rows.extend(rows)
        time.sleep(0.05)
 
    print(f"\nTotal unit rows: {len(all_rows)}")
 
    print(f"\n-- Step 3: Writing {OUTPUT_FILE} --")
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(all_rows)
 
    print(f"Done. Output saved to: {OUTPUT_FILE}")
 
if __name__ == "__main__":
    main()
