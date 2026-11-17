"""
Business Data Cleaner & Reconciliation Engine
Automated validation, encoding correction, deduplication, and normalization for Japanese enterprise workflows.
"""

from __future__ import annotations

import csv
import json
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class CleanResult:
    total_input_rows: int
    valid_rows_count: int
    flagged_rows_count: int
    totals_by_group: Dict[str, int]
    valid_records: List[Dict[str, Any]]
    flagged_records: List[Dict[str, Any]]
    audit_summary: Dict[str, Any] = field(default_factory=dict)


def normalize_text(text: str) -> str:
    """Normalize Japanese Zenkaku/Hankaku characters and strip whitespace."""
    if not text:
        return ""
    # Normalize unicode (NFKC turns full-width digits/alphanumeric into half-width)
    normalized = unicodedata.normalize("NFKC", str(text))
    # Replace full-width ideographic spaces
    normalized = normalized.replace("\u3000", " ").strip()
    return normalized


ERA_START = {"令和": 2018, "R": 2018, "平成": 1988, "H": 1988, "昭和": 1925, "S": 1925}


def _make_date(y: int, m: int, d: int) -> Optional[str]:
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def parse_japanese_date(raw_date: str) -> Optional[str]:
    """Parse ISO, slash, dot, 年月日, YYYYMMDD and Wareki (令和/平成/昭和, R/H/S, 元年) dates to YYYY-MM-DD."""
    norm = normalize_text(raw_date).replace(" ", "")
    if not norm:
        return None

    # Wareki: 令和6年11月1日, 令和元年5月1日, R6.11.1, H31/4/30
    m = re.match(r"^(令和|平成|昭和|[RHSrhs])(元|\d{1,2})[年.\-/](\d{1,2})[月.\-/](\d{1,2})日?$", norm)
    if m:
        era = m.group(1).upper() if len(m.group(1)) == 1 else m.group(1)
        year_in_era = 1 if m.group(2) == "元" else int(m.group(2))
        return _make_date(ERA_START[era] + year_in_era, int(m.group(3)), int(m.group(4)))

    # 2026年11月1日
    m = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$", norm)
    if m:
        return _make_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 2026-11-01, 2026/11/1, 2026.11.1
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", norm)
    if m:
        return _make_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 20261101
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", norm)
    if m:
        return _make_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    return None


def parse_jpy_amount(raw_amount: str) -> Optional[int]:
    """Parse currency strings e.g. '¥12,800', '12800円', '▲500', '-¥500', '(1,200)', '12,800.0' to integer yen."""
    norm = normalize_text(raw_amount)
    # U+2212 minus and other dashes survive NFKC; treat them as a minus sign
    norm = re.sub(r"[−‒–—]", "-", norm)
    norm = re.sub(r"[¥円,\s]", "", norm)
    if not norm:
        return None
    negative = False
    if norm.startswith("(") and norm.endswith(")"):
        negative, norm = True, norm[1:-1]
    if norm[:1] in ("-", "▲", "△"):
        negative, norm = not negative, norm[1:]
    elif norm[:1] == "+":
        norm = norm[1:]
    m = re.fullmatch(r"(\d+)(?:\.0+)?", norm)
    if not m:
        return None  # includes fractional yen such as 12.5, which needs a human decision
    val = int(m.group(1))
    return -val if negative else val


class DataCleaner:
    def __init__(self, key_field: str = "注文ID", date_field: str = "日付", group_field: str = "店舗", amount_field: str = "金額"):
        self.key_field = key_field
        self.date_field = date_field
        self.group_field = group_field
        self.amount_field = amount_field

    def detect_encoding(self, file_path: Path) -> str:
        # Decode the whole file: a 4 KB prefix can end in the middle of a
        # multi-byte character and wrongly fail the UTF-8 check.
        raw_bytes = file_path.read_bytes()
        if raw_bytes.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        for enc in ("utf-8", "cp932"):
            try:
                raw_bytes.decode(enc)
                return enc
            except UnicodeDecodeError:
                pass
        raise ValueError(f"{file_path.name}: neither UTF-8 nor CP932 (Shift-JIS); re-save it in one of those")

    def clean_csv(self, csv_path: str | Path) -> CleanResult:
        path = Path(csv_path)
        encoding = self.detect_encoding(path)
        with path.open("r", encoding=encoding, newline="") as f:
            reader = csv.DictReader(f)
            header = reader.fieldnames or []
            rows = list(reader)

        accepted_by_key: Dict[str, Dict[str, Any]] = {}
        flagged_list: List[Dict[str, Any]] = []

        for line_num, row in enumerate(rows, start=2):
            raw_key = normalize_text(row.get(self.key_field, ""))
            raw_date = normalize_text(row.get(self.date_field, ""))
            raw_group = normalize_text(row.get(self.group_field, ""))
            raw_amount = normalize_text(row.get(self.amount_field, ""))

            # Validation
            if not raw_key:
                flagged_list.append({"line": line_num, "raw": row, "reason": "主キー欠落 (Key Missing)"})
                continue

            parsed_date = parse_japanese_date(raw_date)
            if not parsed_date:
                flagged_list.append({"line": line_num, "raw": row, "reason": f"日付形式不正 (Invalid Date: {raw_date})"})
                continue

            parsed_amount = parse_jpy_amount(raw_amount)
            if parsed_amount is None:
                flagged_list.append({"line": line_num, "raw": row, "reason": f"金額形式不正 (Invalid Amount: {raw_amount})"})
                continue

            if not raw_group:
                flagged_list.append({"line": line_num, "raw": row, "reason": "店舗/グループ欠落 (Group Missing)"})
                continue

            clean_record = {
                self.key_field: raw_key,
                self.date_field: parsed_date,
                self.group_field: raw_group,
                self.amount_field: parsed_amount,
                "line": line_num
            }

            # Duplicate check
            if raw_key in accepted_by_key:
                existing = accepted_by_key[raw_key]
                if (existing[self.date_field] == parsed_date and
                    existing[self.group_field] == raw_group and
                    existing[self.amount_field] == parsed_amount):
                    reason = f"完全重複レコード (Duplicate of Line {existing['line']})"
                else:
                    reason = f"注文ID競合・内容不一致 (Conflict with Line {existing['line']})"
                flagged_list.append({"line": line_num, "raw": row, "reason": reason})
                continue

            accepted_by_key[raw_key] = clean_record

        # Aggregate totals
        totals = defaultdict(int)
        for rec in accepted_by_key.values():
            totals[rec[self.group_field]] += rec[self.amount_field]

        return CleanResult(
            total_input_rows=len(rows),
            valid_rows_count=len(accepted_by_key),
            flagged_rows_count=len(flagged_list),
            totals_by_group=dict(sorted(totals.items())),
            valid_records=list(accepted_by_key.values()),
            flagged_records=flagged_list,
            audit_summary={
                "input_file": str(path.name),
                "columns": list(header),
                "fields": {"key": self.key_field, "date": self.date_field,
                           "group": self.group_field, "amount": self.amount_field},
                "detected_encoding": encoding,
                "processed_at": datetime.now().isoformat()
            }
        )
