"""
Test suite for Business Data Cleaner.
Validates input parsing, deduplication, date normalization, error audit, and Excel export.
"""

import sys
from pathlib import Path

# Add src to path
SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

from cleaner import DataCleaner
from excel_exporter import export_clean_workbook
from openpyxl import load_workbook


def run_tests():
    sample_csv = Path(__file__).resolve().parent.parent / "samples" / "sales_dirty.csv"
    output_excel = Path(__file__).resolve().parent.parent / "output" / "cleaned_sales.xlsx"

    cleaner = DataCleaner()
    res = cleaner.clean_csv(sample_csv)

    print("=== CLEANING RESULTS SUMMARY ===")
    print(f"Total Input Rows: {res.total_input_rows}")
    print(f"Valid Clean Rows: {res.valid_rows_count}")
    print(f"Flagged Anomaly Rows: {res.flagged_rows_count}")
    print(f"Totals by Group: {res.totals_by_group}")

    # Assertions
    assert res.total_input_rows == 10, "Should have 10 input lines"
    # A101 (line 2 and line 4 duplicate) -> 1 accepted, 1 duplicate
    # A104 invalid date (2026/11/31 does not exist in November) -> flagged
    # A106 amount missing -> flagged
    assert res.valid_rows_count == 7, f"Expected 7 valid rows, got {res.valid_rows_count}"
    assert res.flagged_rows_count == 3, f"Expected 3 flagged rows, got {res.flagged_rows_count}"

    # Check date normalization: 令和8年11月2日 -> 2026-11-02
    r_a103 = next(r for r in res.valid_records if r["注文ID"] == "A103")
    assert r_a103["日付"] == "2026-11-02", "Reiwa date must convert to ISO 2026-11-02"

    # Check JPY amount parsing: ￥15000 -> 15000, ▲1200 -> -1200
    r_a108 = next(r for r in res.valid_records if r["注文ID"] == "A108")
    assert r_a108["金額"] == -1200, "Accounting negative ▲1200 must convert to -1200"
    r_a109 = next(r for r in res.valid_records if r["注文ID"] == "A109")
    assert r_a109["金額"] == 15000, "￥15000 must convert to 15000"

    # Export to Excel
    export_clean_workbook(res, output_excel)
    assert output_excel.exists(), "Excel file must be generated"

    # Verify Excel readback
    wb = load_workbook(output_excel, read_only=True)
    assert "店舗別集計" in wb.sheetnames, "Summary sheet missing"
    assert "整形済み明細" in wb.sheetnames, "Detail sheet missing"
    assert "要確認データ" in wb.sheetnames, "Issues sheet missing"
    wb.close()

    print(f"\n[PASS] Excel generated successfully at: {output_excel}")
    print(">>> ALL DATA CLEANER TESTS PASSED SUCCESSFULLY! <<<")


if __name__ == "__main__":
    run_tests()
