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