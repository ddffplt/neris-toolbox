import tempfile
import unittest
from pathlib import Path

import openpyxl

import DepartmentREports_by_Prompt as reports


class DepartmentReportsTests(unittest.TestCase):
    def test_simultaneous_summary_rows_are_not_treated_as_entities(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            fire = folder / "fire.xlsx"
            simultaneous = folder / "simultaneous.xlsx"

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Entity Fire Reports"
            sheet.append(["FD25001000", "Example Fire Department"])
            sheet.append(["Fire Incident Type", "Number of Fires"])
            sheet.append(["FIRE", 1])
            workbook.save(fire)

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Overview"
            headers = [
                "Incident NERIS ID", "Department/Entity ID", "Department Name",
                "Call Date", "Simultaneous Call Count",
            ]
            sheet.append(headers)
            sheet.append(["incident-1", "FD25001000", "Example Fire Department",
                          "2026-01-01T12:00:00+00:00", 2])
            sheet.append(headers)
            sheet.append(["FD25001000", "Example Fire Department"])
            sheet.append(["Month", "Year", "Simultaneous Incidents"])
            sheet.append(["January", 2026, 5])
            workbook.save(simultaneous)

            _, rows = reports.read_simultaneous_rows(simultaneous)
            self.assertEqual(list(rows), ["FD25001000"])
            self.assertEqual(len(rows["FD25001000"]), 1)

            output = folder / "output"
            reports.build_reports([fire, simultaneous], output)
            output_files = list(output.glob("*.xlsx"))
            self.assertEqual([path.name for path in output_files],
                             ["Example_Fire_Department_FD25001000_combined.xlsx"])
            result = openpyxl.load_workbook(output_files[0], read_only=True, data_only=True)
            self.assertIn("Simultaneous Calls", result.sheetnames)
            result.close()


if __name__ == "__main__":
    unittest.main()
