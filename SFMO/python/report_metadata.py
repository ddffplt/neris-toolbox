# pyright: reportMissingImports=false
from datetime import datetime


def add_report_sources_sheet(workbook, report_name, date_range, details):
    from openpyxl.styles import Font  # type: ignore[import-not-found]

    worksheet = workbook.create_sheet("Report Sources")
    worksheet.append(["Report source", report_name])
    worksheet.append(["Report date/range", date_range])
    worksheet.append(["Generated at", datetime.now().astimezone().isoformat(timespec="seconds")])
    worksheet.append([])
    for label, value in details:
        worksheet.append([label, value])
    worksheet.column_dimensions["A"].width = 36
    worksheet.column_dimensions["B"].width = 85
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    for cell in worksheet[5] if len(details) else []:
        cell.font = Font(bold=True)


def incident_date_range(incidents):
    timestamps = []
    for incident in incidents:
        dispatch = incident.get("dispatch") or {}
        value = dispatch.get("call_create") or dispatch.get("call_create_start")
        if value:
            timestamps.append(str(value))
    if not timestamps:
        return "No incident timestamps found"
    return f"{min(timestamps)} through {max(timestamps)} (observed incident timestamps)"


def highlight_fetch_issue(cells):
    from openpyxl.styles import Font, PatternFill

    for cell in cells:
        cell.fill = PatternFill("solid", fgColor="FFC7CE")
        cell.font = Font(color="9C0006", bold=True)


def add_fetch_issues_sheet(workbook, errors, entity_names):
    if not errors:
        return
    from openpyxl.styles import Alignment

    sheet = workbook.create_sheet("Fetch Issues")
    sheet.append(["Entity ID", "Department Name", "Error", "Impact"])
    for entity_id, error in sorted(errors.items()):
        sheet.append([entity_id, entity_names.get(entity_id, ""), str(error),
                      "Data for this entity could not be fully retrieved; report totals may be incomplete."])
        row = sheet[sheet.max_row]
        highlight_fetch_issue(row)
        sheet.row_dimensions[sheet.max_row].height = 60
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for column, width in (("A", 18), ("B", 35), ("C", 70), ("D", 70)):
        sheet.column_dimensions[column].width = width


def flag_incomplete_entity_tables(workbook, errors, entity_names, county=False):
    if not errors:
        return
    for sheet_name in ("Incident Type Summary", "Monthly Incident Counts"):
        sheet = workbook[sheet_name]
        status_column = sheet.max_column + 1
        sheet.cell(1, status_column, "Fetch Status")
        for row_number in range(2, sheet.max_row + 1):
            entity_id = sheet.cell(row_number, 1).value
            if entity_id == "TOTAL":
                sheet.cell(row_number, status_column, "Incomplete total - see Fetch Issues")
                highlight_fetch_issue(sheet[row_number])
                continue
            failed = (any(str(eid).upper().startswith(str(entity_id).upper()) for eid in errors)
                      if county and entity_id else entity_id in errors)
            if not failed:
                continue
            if not county:
                for cell in sheet[row_number][2:status_column - 1]:
                    cell.value = None
            status = "Partial county data" if county else "Data unavailable"
            sheet.cell(row_number, status_column, f"{status} - see Fetch Issues")
            highlight_fetch_issue(sheet[row_number])
        from openpyxl.utils import get_column_letter
        sheet.column_dimensions[get_column_letter(status_column)].width = 38
    add_fetch_issues_sheet(workbook, errors, entity_names)