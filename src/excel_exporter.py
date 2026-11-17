"""
Professional Styled Excel Exporter
Creates formatted multi-sheet workbooks with KPI summaries, freeze panes, and audit logs.
"""

from __future__ import annotations

import json
import unicodedata
from pathlib import Path
from typing import Any, Dict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _display_width(value: Any) -> int:
    """Column width in Excel units: full-width (CJK) characters take two."""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in str(value or ""))


def export_clean_workbook(clean_result: Any, output_path: str | Path) -> Path:
    out = Path(output_path)
    fields = clean_result.audit_summary.get("fields", {})
    key_f = fields.get("key", "注文ID")
    date_f = fields.get("date", "日付")
    group_f = fields.get("group", "店舗")
    amount_f = fields.get("amount", "金額")
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="17324D")
    header_font = Font(name="Meiryo", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Meiryo", size=10)
    total_font = Font(name="Meiryo", size=11, bold=True, color="127567")
    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE")
    )

    # Sheet 1: Group Summary (店舗別売上集計)
    ws_summary = wb.active
    ws_summary.title = "店舗別集計"
    ws_summary.append(["集計グループ（店舗）", "売上合計（円）", "件数"])
    for cell in ws_summary[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    total_sum = 0
    total_count = 0
    for group, amount in clean_result.totals_by_group.items():
        count = sum(1 for r in clean_result.valid_records if r.get(group_f) == group)
        ws_summary.append([group, amount, count])
        total_sum += amount
        total_count += count

    ws_summary.append(["合計 (Total)", total_sum, total_count])
    last_row = ws_summary.max_row
    for col in range(1, 4):
        c = ws_summary.cell(row=last_row, column=col)
        c.font = total_font
        c.fill = PatternFill("solid", fgColor="E6F4EA")

    # Sheet 2: Valid Detail Records (整形済み明細)
    ws_detail = wb.create_sheet("整形済み明細")
    detail_headers = [key_f, date_f, group_f, f"{amount_f}（円）", "元CSV行番号"]
    ws_detail.append(detail_headers)
    for cell in ws_detail[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for rec in clean_result.valid_records:
        ws_detail.append([
            rec.get(key_f),
            rec.get(date_f),
            rec.get(group_f),
            rec.get(amount_f),
            rec.get("line")
        ])

    # Sheet 3: Flagged Issues (要確認データ)
    ws_issues = wb.create_sheet("要確認データ")
    ws_issues.append(["元行番号", "除外・フラグ理由", "元データJSON"])
    issue_header_fill = PatternFill("solid", fgColor="8C1D1D")
    for cell in ws_issues[1]:
        cell.fill = issue_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for item in clean_result.flagged_records:
        ws_issues.append([
            item.get("line"),
            item.get("reason"),
            json.dumps(item.get("raw"), ensure_ascii=False)
        ])

    # Format all sheets
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.views.sheetView[0].showGridLines = True
        for col in ws.columns:
            max_len = max(_display_width(cell.value) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
            for cell in col:
                cell.border = thin_border
                if cell.row > 1 and isinstance(cell.value, int) and not isinstance(cell.value, bool) and ws[1][cell.column - 1].value != "元CSV行番号":
                    cell.number_format = "#,##0"

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out
