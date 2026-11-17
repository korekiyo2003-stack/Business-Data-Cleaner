"""
Test suite for Business Data Cleaner.
Validates input parsing, deduplication, date normalization, error audit, and Excel export.
"""

import sys
from pathlib import Path

# Add src to path
SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import tempfile

from cleaner import DataCleaner, parse_japanese_date, parse_jpy_amount
from excel_exporter import export_clean_workbook
from openpyxl import load_workbook


def run_tests():
    sample_csv = Path(__file__).resolve().parent.parent / "samples" / "sales_dirty.csv"
    output_excel = Path(tempfile.mkdtemp()) / "cleaned_sales.xlsx"

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


def run_edge_case_tests():
    # Dates
    cases = {
        "令和元年5月1日": "2019-05-01", "R6.11.1": "2024-11-01", "r8-11-1": "2026-11-01", "H31/4/30": "2019-04-30",
        "平成31年4月30日": "2019-04-30", "昭和64年1月7日": "1989-01-07",
        "2026年11月1日": "2026-11-01", "２０２６／１１／０１": "2026-11-01", "20261101": "2026-11-01",
        "2026/11/31": None, "11/01/2026": None, "": None,
    }
    for raw, want in cases.items():
        assert parse_japanese_date(raw) == want, (raw, parse_japanese_date(raw), want)

    # Amounts
    amounts = {
        "¥12,800": 12800, "12800円": 12800, "▲1,200": -1200, "△500": -500, "-¥500": -500,
        "−300": -300, "(1,200)": -1200, "１２，８００": 12800, "12,800.0": 12800,
        "+700": 700, "12.5": None, "abc": None, "": None, "¥": None,
    }
    for raw, want in amounts.items():
        assert parse_jpy_amount(raw) == want, (raw, parse_jpy_amount(raw), want)

    tmp = Path(tempfile.mkdtemp())
    cleaner = DataCleaner()

    # UTF-8 file where a 3-byte character straddles byte 4096 (the old 4 KB sniff rejected it as UTF-8)
    header = "注文ID,日付,店舗,金額\n"
    key_len = 4095 - len(header.encode("utf-8")) - len(",2026/11/01,")
    text = header + "K" * key_len + ",2026/11/01,横浜,100\n" + "K2,2026/11/02,横浜,200\n"
    raw = text.encode("utf-8")
    try:
        raw[:4096].decode("utf-8")
        raise AssertionError("test setup: byte 4096 should split a character")
    except UnicodeDecodeError:
        pass
    f_utf8 = tmp / "utf8.csv"
    f_utf8.write_bytes(raw)
    res = cleaner.clean_csv(f_utf8)
    assert res.audit_summary["detected_encoding"] == "utf-8", res.audit_summary
    assert res.totals_by_group == {"横浜": 300}, res.totals_by_group

    # Shift-JIS (CP932) file
    f_sjis = tmp / "sjis.csv"
    f_sjis.write_bytes("注文ID,日付,店舗,金額\nC1,令和8年11月2日,東京,▲300\n".encode("cp932"))
    res = cleaner.clean_csv(f_sjis)
    assert res.audit_summary["detected_encoding"] == "cp932"
    assert res.totals_by_group == {"東京": -300}

    # Custom column names flow through to the Excel sheets
    f_custom = tmp / "custom.csv"
    f_custom.write_text("伝票番号,売上日,部門,売上\nD1,2026-11-03,営業部,5000\n", encoding="utf-8")
    res = DataCleaner(key_field="伝票番号", date_field="売上日", group_field="部門", amount_field="売上").clean_csv(f_custom)
    out = export_clean_workbook(res, tmp / "custom.xlsx")
    wb = load_workbook(out, read_only=True)
    detail = list(wb["整形済み明細"].iter_rows(values_only=True))
    summary = list(wb["店舗別集計"].iter_rows(values_only=True))
    wb.close()
    assert detail[0][:3] == ("伝票番号", "売上日", "部門"), detail[0]
    assert detail[1][:4] == ("D1", "2026-11-03", "営業部", 5000), detail[1]
    assert summary[1] == ("営業部", 5000, 1), summary[1]

    print(">>> EDGE CASE TESTS PASSED <<<")


def run_cli_tests():
    import json
    import subprocess
    root = Path(__file__).resolve().parent.parent
    tmp = Path(tempfile.mkdtemp())
    run = lambda *a: subprocess.run([sys.executable, str(root / "clean.py"), *a], capture_output=True, text=True,
                                    encoding="utf-8", env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})
    r = run(str(root / "samples" / "sales_dirty.csv"), "-o", str(tmp / "o.xlsx"), "--json")
    assert r.returncode == 0, r.stderr
    summary = json.loads(r.stdout)
    assert (summary["valid"], summary["flagged"]) == (7, 3) and (tmp / "o.xlsx").exists()
    r = run(str(root / "samples" / "sales_dirty.csv"), "--key", "伝票")
    assert r.returncode == 2 and "伝票" in r.stderr
    r = run(str(tmp / "missing.csv"))
    assert r.returncode == 2
    print(">>> CLI TESTS PASSED <<<")


if __name__ == "__main__":
    run_tests()
    run_edge_case_tests()
    run_cli_tests()
