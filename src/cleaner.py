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


def parse_japanese_date(raw_date: str) -> Optional[str]:
    """Parse dates in ISO, slash, dot, or Wareki formats to standard YYYY-MM-DD."""
    norm = normalize_text(raw_date).replace("/", "-").replace(".", "-")
    # Match Reiwa date e.g. 令和6年11月1日
    m = re.match(r"(?:令和|R)(\d+)年(\d+)月(\d+)日", norm)
    if m:
        year = 2018 + int(m.group(1))
        month = int(m.group(2))
        day = int(m.group(3))
        try:
            return date(year, month, day).isoformat()
        except ValueError:
            return None

    # Match YYYY-MM-DD or YYYY-M-D
    m2 = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", norm)
    if m2:
        try:
            return date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3))).isoformat()
        except ValueError:
            return None

    return None


def parse_jpy_amount(raw_amount: str) -> Optional[int]:
    """Parse currency strings e.g. '¥12,800', '12800円', '▲500' to integer."""
    norm = normalize_text(raw_amount)
    if not norm:
        return None
    # Check negative symbol or accounting triangle ▲ / △
    negative = False
    if norm.startswith("-") or norm.startswith("▲") or norm.startswith("△"):
        negative = True
        norm = norm.lstrip("-▲△")
    # Strip symbols
    cleaned = re.sub(r"[¥円,\s]", "", norm)
    if not cleaned.isdigit():
        return None
    val = int(cleaned)
    return -val if negative else val


class DataCleaner:
    def __init__(self, key_field: str = "注文ID", date_field: str = "日付", group_field: str = "店舗", amount_field: str = "金額"):
        self.key_field = key_field
        self.date_field = date_field
        self.group_field = group_field
        self.amount_field = amount_field

    def detect_encoding(self, file_path: Path) -> str:
        raw_bytes = file_path.read_bytes()[:4096]
        # Check UTF-8 BOM
        if raw_bytes.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        # Try UTF-8
        try:
            raw_bytes.decode("utf-8")
            return "utf-8"
        except UnicodeDecodeError:
            pass
        # Try CP932 / Shift-JIS
        try:
            raw_bytes.decode("cp932")
            return "cp932"
        except UnicodeDecodeError:
            return "utf-8-sig"

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
                "detected_encoding": encoding,
                "processed_at": datetime.now().isoformat()
            }
        )
