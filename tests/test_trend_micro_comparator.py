from __future__ import annotations

import csv
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from TrendMicroComparator import (
    CustomerRecord,
    TrendMicroFormatError,
    compare_records,
    compare_workbooks,
    read_trend_micro_workbook,
    write_comparison_csv,
)


class TrendMicroComparatorTests(unittest.TestCase):
    def _save_workbook(
        self,
        path: Path,
        rows: list[list[object]],
        *,
        complete_export: bool = False,
    ) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Small and medium business"
        if complete_export:
            sheet.append(["Note: peak values are used for billing"])
            sheet.append(
                [
                    "Customer ID",
                    "Customer",
                    "City",
                    "State",
                    "Service/Product",
                    "Service Plan",
                    "Provisioned",
                    "Used",
                    "License Start Date",
                    "License Expiration Date",
                ]
            )
            for index, row in enumerate(rows, start=1):
                customer, city, state, plan, provisioned, used = row
                sheet.append(
                    [
                        f"id-{index}",
                        customer,
                        city,
                        state,
                        "Worry-Free Services",
                        plan,
                        provisioned,
                        used,
                        "2026-01-01",
                        "2027-01-01",
                    ]
                )
            sheet.append([None, None, None, None, None, None, None, 999])
        else:
            sheet.append(
                ["Customer", "City", "State", "Service Plan", "Provisioned", "Unit", "Used"]
            )
            for customer, city, state, plan, provisioned, used in rows:
                sheet.append([customer, city, state, plan, provisioned, "Seats", used])
        workbook.save(path)
        workbook.close()

    def test_reads_complete_export_and_ignores_totals(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "complete.xlsx"
            self._save_workbook(
                path,
                [["Acme", "Mebane", "NC", "Annual", 5, 4]],
                complete_export=True,
            )
            data = read_trend_micro_workbook(path)

        self.assertEqual(data.sheet_name, "Small and medium business")
        self.assertEqual(len(data.records), 1)
        self.assertEqual(data.records[0].customer_id, "id-1")
        self.assertEqual(data.records[0].provisioned, 5)
        self.assertEqual(data.records[0].used, 4)

    def test_reports_each_requested_difference(self) -> None:
        old = [
            CustomerRecord("Alpha", "Annual", 5, 4),
            CustomerRecord("Beta", "Annual", 2, 2),
            CustomerRecord("Dropped Co", "Biennial", 1, 1),
        ]
        new = [
            CustomerRecord("Alpha", "EDR Annual", 7, 3),
            CustomerRecord("Beta", "Annual", 2, 2),
            CustomerRecord("New Co", "Annual", 4, 4),
        ]

        rows = {row["Customer"]: row for row in compare_records(old, new)}

        self.assertEqual(
            rows["Alpha"]["Status"],
            "SERVICE PLAN CHANGE; PROVISIONED CHANGE; USED SEATS CHANGE",
        )
        self.assertEqual(rows["Alpha"]["Provisioned Change"], 2)
        self.assertEqual(rows["Alpha"]["Used Seats Change"], -1)
        self.assertEqual(rows["Beta"]["Status"], "-")
        self.assertEqual(rows["New Co"]["Status"], "NEW CUSTOMER")
        self.assertEqual(rows["New Co"]["Provisioned Change"], 4)
        self.assertEqual(rows["Dropped Co"]["Status"], "DROPPED CUSTOMER")
        self.assertEqual(rows["Dropped Co"]["New Provisioned"], "")
        self.assertEqual(rows["Dropped Co"]["Provisioned Change"], "")

    def test_matches_customer_by_id_when_name_changes(self) -> None:
        old = [CustomerRecord("Old Name", "Annual", 3, 3, customer_id="ABC")]
        new = [CustomerRecord("New Name", "Annual", 3, 3, customer_id="abc")]
        rows = compare_records(old, new)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Customer"], "New Name")
        self.assertEqual(rows[0]["Status"], "CUSTOMER NAME CHANGE")

    def test_duplicate_customer_rows_prefer_matching_plan(self) -> None:
        old = [
            CustomerRecord("Acme", "Annual", 5, 5),
            CustomerRecord("Acme", "EDR Annual", 2, 2),
        ]
        new = [
            CustomerRecord("Acme", "EDR Annual", 3, 3),
            CustomerRecord("Acme", "Annual", 5, 5),
        ]
        rows = compare_records(old, new)

        self.assertEqual(len(rows), 2)
        annual = next(row for row in rows if row["New Service Plan"] == "Annual")
        edr = next(row for row in rows if row["New Service Plan"] == "EDR Annual")
        self.assertEqual(annual["Status"], "-")
        self.assertEqual(edr["Status"], "PROVISIONED CHANGE; USED SEATS CHANGE")

    def test_end_to_end_comparison_writes_excel_friendly_csv(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            old_path = directory / "old.xlsx"
            new_path = directory / "new.xlsx"
            output_path = directory / "result.csv"
            self._save_workbook(
                old_path,
                [["Acme", "Mebane", "NC", "Annual", 5, 4]],
                complete_export=True,
            )
            self._save_workbook(
                new_path,
                [["Acme", "Mebane", "NC", "Annual", 6, 6]],
            )

            report = compare_workbooks(old_path, new_path, output_path)
            with output_path.open(encoding="utf-8-sig", newline="") as csv_file:
                output_rows = list(csv.DictReader(csv_file))

        self.assertEqual(report.changed_count, 1)
        self.assertEqual(len(output_rows), 1)
        self.assertEqual(output_rows[0]["Status"], "PROVISIONED CHANGE; USED SEATS CHANGE")
        self.assertEqual(output_rows[0]["Old Provisioned"], "5")
        self.assertEqual(output_rows[0]["New Provisioned"], "6")

    def test_rejects_workbook_without_required_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "invalid.xlsx"
            workbook = Workbook()
            workbook.active.append(["Customer", "Something Else"])
            workbook.save(path)
            workbook.close()

            with self.assertRaises(TrendMicroFormatError):
                read_trend_micro_workbook(path)

    def test_csv_neutralizes_formula_like_customer_text(self) -> None:
        rows = compare_records(
            [CustomerRecord("Unchanged", "Annual", 1, 1)],
            [
                CustomerRecord("Unchanged", "Annual", 1, 1),
                CustomerRecord("=FORMULA", "Annual", 1, 1),
            ],
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "safe.csv"
            write_comparison_csv(rows, output_path)
            with output_path.open(encoding="utf-8-sig", newline="") as csv_file:
                output_rows = {row["Customer"]: row for row in csv.DictReader(csv_file)}

        self.assertEqual(output_rows["'=FORMULA"]["Customer"], "'=FORMULA")
        self.assertEqual(output_rows["Unchanged"]["Status"], "-")


if __name__ == "__main__":
    unittest.main()
