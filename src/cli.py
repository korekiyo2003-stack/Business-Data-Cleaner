"""Command line: clean a sales CSV into a checked Excel workbook.

    python clean.py 売上.csv                       -> 売上_cleaned.xlsx next to the CSV
    python clean.py 売上.csv -o 結果.xlsx --key 伝票番号 --date 売上日 --group 部門 --amount 売上
    python clean.py 売上.csv --json                 -> summary as JSON on stdout (for scripts)
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from cleaner import DataCleaner


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="clean.py", description="売上CSVを検証・正規化して Excel（店舗別集計 / 整形済み明細 / 要確認データ）に出力します。")
    p.add_argument("csv", type=Path, help="入力CSV（UTF-8 / UTF-8 BOM / Shift-JIS(CP932) を自動判別）")
    p.add_argument("-o", "--output", type=Path, help="出力 .xlsx（省略時は <入力名>_cleaned.xlsx）")
    p.add_argument("--key", default="注文ID", help="主キーの列名（既定: 注文ID）")
    p.add_argument("--date", default="日付", help="日付の列名（既定: 日付）")
    p.add_argument("--group", default="店舗", help="集計単位の列名（既定: 店舗）")
    p.add_argument("--amount", default="金額", help="金額の列名（既定: 金額）")
    p.add_argument("--json", action="store_true", help="結果の要約を JSON で標準出力に出す")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not args.csv.exists():
        print(f"入力ファイルがありません: {args.csv}", file=sys.stderr)
        return 2
    cleaner = DataCleaner(key_field=args.key, date_field=args.date, group_field=args.group, amount_field=args.amount)
    try:
        result = cleaner.clean_csv(args.csv)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 2
    missing = [c for c in (args.key, args.date, args.group, args.amount) if c not in result.audit_summary.get("columns", [])]
    if missing:
        print(f"CSV に列がありません: {', '.join(missing)}（あるのは: {', '.join(result.audit_summary.get('columns', []))}）。"
              f" --key/--date/--group/--amount で列名を指定してください。", file=sys.stderr)
        return 2

    from excel_exporter import export_clean_workbook  # openpyxl is only needed from here on
    out = args.output or args.csv.with_name(args.csv.stem + "_cleaned.xlsx")
    export_clean_workbook(result, out)

    if args.json:
        print(json.dumps({"input": str(args.csv), "output": str(out), "encoding": result.audit_summary["detected_encoding"],
                          "rows": result.total_input_rows, "valid": result.valid_rows_count, "flagged": result.flagged_rows_count,
                          "totals": result.totals_by_group,
                          "issues": [{"line": f["line"], "reason": f["reason"]} for f in result.flagged_records]},
                         ensure_ascii=False, indent=2))
    else:
        print(f"{args.csv.name}（{result.audit_summary['detected_encoding']}）: {result.total_input_rows} 行 → "
              f"有効 {result.valid_rows_count} 行 ＋ 要確認 {result.flagged_rows_count} 行")
        for g, v in result.totals_by_group.items():
            print(f"  {g}: {'-' if v < 0 else ''}¥{abs(v):,}")
        for f in result.flagged_records:
            print(f"  要確認 {f['line']}行目: {f['reason']}")
        print(f"出力: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
