"""Copy-and-paste building blocks for new NERIS scripts: entity scope + date range prompts.

Everything here is standard library only. Either `from prompt_snippets import ...`
or copy the pieces into a new script. Usage inside a script's prompt_config():

    state_code = input("State Code (e.g. MA, MI, VA): ").strip().upper()
    entity_scope, entity_id, entity_prefix, entity_names = prompt_entity_scope()
    range_type, start_date, end_date = prompt_date_range()

and later, to query the API:

    call_create_start, call_create_end = calculate_date_range(range_type, start_date, end_date)

These return values match the existing scripts, so they plug straight into
get_scoped_incidents() and export_to_excel() (entity_scope is one of
"statewide", "fd25_prefix", "entity", "multiple_entities").
"""

from datetime import datetime, timedelta, timezone

# ── Time options ─────────────────────────────────────────────────────────────

DATE_RANGE_OPTIONS = [
    "All Records",
    "Occurred Today",
    "Occurred Yesterday",
    "Occurred Last 7 Days",
    "Occurred Last 30 Days",
    "Occurred Last 90 Days",
    "Occurred in 2025",
    "Occurred in 2026",
    "Occurred in current calendar year to date",
    "Custom Date Range",
]


def prompt_date_range():
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
    return range_type, start_date, end_date


def calculate_date_range(range_type, start_date=None, end_date=None):
    """Return (call_create_start, call_create_end) as UTC-aware datetimes (None = unbounded)."""
    now = datetime.now(tz=timezone.utc)
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
        return datetime(now.year, 1, 1, tzinfo=timezone.utc), now
    elif range_type.startswith("Occurred in "):
        year = int(range_type.split()[-1])
        return (
            datetime(year, 1, 1, 0, 0, 0, tzinfo=timezone.utc),
            datetime(year, 12, 31, 23, 59, 59, tzinfo=timezone.utc),
        )
    elif range_type == "Custom Date Range" and start_date and end_date:
        start_dt = datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc)
        end_dt = datetime.combine(end_date, datetime.max.time().replace(microsecond=0), tzinfo=timezone.utc)
        return start_dt, end_dt
    return None, None


# ── Entity choice options ────────────────────────────────────────────────────

ENTITY_SCOPE_LABELS = {
    "entity": "Individual entity",
    "fd25_prefix": "FD25 county prefix",
    "statewide": "Entire state",
    "multiple_entities": "Multiple entities",
}

FD25_COUNTY_CODES = {f"{number:03d}" for number in range(1, 28, 2)}


def normalize_fd25_prefix(value):
    prefix = (value or "").strip().upper().rstrip("*")
    if prefix.startswith("FD25") and prefix[4:] in FD25_COUNTY_CODES:
        return prefix
    return None


def prompt_entity_scope():
    """Return (entity_scope, entity_id, entity_prefix, entity_names)."""
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
    return entity_scope, entity_id, entity_prefix, entity_names
