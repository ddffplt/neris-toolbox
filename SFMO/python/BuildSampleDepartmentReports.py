#!/usr/bin/env python3
"""Build per-entity combined workbooks from three existing NERIS outputs."""

import argparse
import re
from copy import copy
from pathlib import Path

import openpyxl


SAMPLE_DIR = Path(r"C:\Users\ddion\Python Scripts\Sample Output files")
FIRE_FILE = SAMPLE_DIR / "neris_fire_activity_MA_2026-04-01_to_2026-08-31_20260910_120038.xlsx"
TYPE_FILE = SAMPLE_DIR / "neris_incident_type_counts_MA_20260909_130127.xlsx"
MONTH_FILE = None
ENTITY_ID = re.compile(r"^FD[A-Za-z0-9]+$")
NUMBER_FORMAT = '#,##0;(#,##0);-'


def values(row):
    return [cell.value for cell in row]


def read_flat_rows(path, sheet_name):
    worksheet = openpyxl.load_workbook(path, read_only=True, data_only=True)[sheet_name]
    rows = list(worksheet.iter_rows(values_only=True))
    headers = list(rows[0])
    return headers, [list(row) for row in rows[1:] if row and row[0] and ENTITY_ID.match(str(row[0]))]


def read_fire_blocks(path):
    worksheet = openpyxl.load_workbook(path, read_only=False, data_only=False).active
    blocks = {}
    current_id = None
    current_name = ""
    current_rows = []
    current_start = None
    for row_number, row in enumerate(worksheet.iter_rows(), start=1):
        row_values = values(row)
        if row_values[0] and ENTITY_ID.match(str(row_values[0])) and row_values[1]:
            if current_id:
                blocks[current_id] = (current_name, current_rows, current_start, row_number - 1)
            current_id = str(row_values[0])
            current_name = str(row_values[1])
            current_rows = []
            current_start = row_number
        elif current_id and any(value is not None for value in row_values):
            current_rows.append(row_values)
    if current_id:
        blocks[current_id] = (current_name, current_rows, current_start, worksheet.max_row)
    return blocks


def safe_name(value):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._") or "department"


def add_sheet(workbook, title, headers, rows, entity_name, entity_id):
    worksheet = workbook.create_sheet(title[:31])
    worksheet.append([entity_name, entity_id])
    worksheet.append([])
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    for row in worksheet.iter_rows(min_row=4):
        for cell in row:
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = NUMBER_FORMAT
    worksheet.freeze_panes = "A4"
    for cell in worksheet[3]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
    for column in worksheet.columns:
        worksheet.column_dimensions[column[0].column_letter].width = min(
            max(len(str(cell.value or "")) for cell in column) + 2, 36
        )


def add_fire_sheet(workbook, source_worksheet, block, entity_name, entity_id):
    worksheet = workbook.create_sheet("Fire Activity")
    _, _, start_row, end_row = block
    row_offset = 1 - start_row
    for source_row in range(start_row, end_row + 1):
        target_row = source_row + row_offset
        if source_row in source_worksheet.row_dimensions:
            worksheet.row_dimensions[target_row].height = source_worksheet.row_dimensions[source_row].height
        for source_column in range(1, source_worksheet.max_column + 1):
            source_cell = source_worksheet.cell(source_row, source_column)
            target_cell = worksheet.cell(target_row, source_column, source_cell.value)
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)
            if source_cell.number_format:
                target_cell.number_format = source_cell.number_format
            if source_cell.alignment:
                target_cell.alignment = copy(source_cell.alignment)
            if source_cell.font:
                target_cell.font = copy(source_cell.font)
            if source_cell.fill:
                target_cell.fill = copy(source_cell.fill)
            if source_cell.border:
                target_cell.border = copy(source_cell.border)
    for column_letter, dimension in source_worksheet.column_dimensions.items():
        worksheet.column_dimensions[column_letter].width = dimension.width
        worksheet.column_dimensions[column_letter].hidden = dimension.hidden
    for merged_range in source_worksheet.merged_cells.ranges:
        if merged_range.min_row >= start_row and merged_range.max_row <= end_row:
            worksheet.merge_cells(
                start_row=merged_range.min_row + row_offset,
                start_column=merged_range.min_col,
                end_row=merged_range.max_row + row_offset,
                end_column=merged_range.max_col,
            )
    worksheet.freeze_panes = source_worksheet.freeze_panes
    worksheet.sheet_view.showGridLines = source_worksheet.sheet_view.showGridLines
    worksheet.sheet_properties.pageSetUpPr = copy(source_worksheet.sheet_properties.pageSetUpPr)
    worksheet.oddHeader.center.text = source_worksheet.oddHeader.center.text
    worksheet.oddFooter.center.text = source_worksheet.oddFooter.center.text
    return worksheet


def add_sources_sheet(workbook, entity_name, entity_id, date_range):
    worksheet = workbook.create_sheet("Report Sources")
    worksheet.append([entity_name, entity_id])
    worksheet.append(["Source file", "Date/range information"])
    worksheet.append([FIRE_FILE.name, "Fire activity source workbook"])
    worksheet.append([TYPE_FILE.name, "Incident type count source workbook"])
    if MONTH_FILE:
        worksheet.append([MONTH_FILE.name, "Monthly incident count source workbook"])
    worksheet.append(["Report date/range", date_range])
    worksheet.append(["Sample note", "This workbook combines only rows belonging to the entity shown above."])
    worksheet.column_dimensions["A"].width = 70
    worksheet.column_dimensions["B"].width = 85
    for cell in worksheet[2]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")


def build_reports(output_dir, limit, requested_ids=None, date_range="Source workbook range"):
    fire_workbook = openpyxl.load_workbook(FIRE_FILE, read_only=False, data_only=False)
    fire_worksheet = fire_workbook.active
    fire_blocks = read_fire_blocks(FIRE_FILE)
    type_headers, type_rows = read_flat_rows(TYPE_FILE, "Incident Type Summary")
    type_month_headers, type_month_rows = read_flat_rows(TYPE_FILE, "Monthly Incident Counts")
    rows_by_entity = {
        "type": {str(row[0]): row for row in type_rows},
        "type_month": {str(row[0]): row for row in type_month_rows},
    }
    entity_ids = sorted(set(fire_blocks) | set(rows_by_entity["type"]))
    if requested_ids:
        missing_ids = sorted(set(requested_ids) - set(entity_ids))
        if missing_ids:
            raise ValueError(f"Requested entity IDs were not found: {', '.join(missing_ids)}")
        entity_ids = [entity_id for entity_id in entity_ids if entity_id in requested_ids]
    if limit:
        entity_ids = entity_ids[:limit]
    output_dir.mkdir(parents=True, exist_ok=True)

    for entity_id in entity_ids:
        entity_name = (
            fire_blocks.get(entity_id, ("", [], 0, 0))[0]
            or rows_by_entity["type"].get(entity_id, ["", ""])[1]
            or entity_id
        )
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
        fire_block = fire_blocks.get(entity_id)
        if fire_block:
            add_fire_sheet(workbook, fire_worksheet, fire_block, entity_name, entity_id)
        else:
            add_sheet(workbook, "Fire Activity", [], [], entity_name, entity_id)
        add_sheet(workbook, "Incident Types", type_headers, [rows_by_entity["type"][entity_id]], entity_name, entity_id) if entity_id in rows_by_entity["type"] else None
        add_sheet(workbook, "Type Monthly Activity", type_month_headers, [rows_by_entity["type_month"][entity_id]], entity_name, entity_id) if entity_id in rows_by_entity["type_month"] else None
        add_sources_sheet(workbook, entity_name, entity_id, date_range)
        path = output_dir / f"{safe_name(entity_name)}_{safe_name(entity_id)}_sample.xlsx"
        workbook.save(path)
        print(path)

    print(f"Created {len(entity_ids)} combined department report(s) in {output_dir}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--output-dir", default=r"C:\Users\ddion\Python Scripts\Email\test-email-examples\first-10-reports")
    parser.add_argument("--entity-id", action="append", dest="entity_ids", help="Include only this NERIS entity ID; repeat for multiple departments.")
    parser.add_argument("--date-range", default="Source workbook range", help="Date range recorded in each generated workbook.")
    args = parser.parse_args()
    if args.limit is not None and args.limit < 1:
        raise ValueError("--limit must be at least 1")
    build_reports(Path(args.output_dir), args.limit, args.entity_ids, args.date_range)


if __name__ == "__main__":
    main()