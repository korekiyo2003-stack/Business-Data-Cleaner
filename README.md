<p align="center"><img src="docs/cover.png" alt="Business-Data-Cleaner" width="100%"></p>

# Business-Data-Cleaner: Enterprise ETL & Audit Engine

<div align="center">

**Automated Japanese Business Data Cleaning, Normalization & Multi-Sheet Excel Reporting**  
*Shift-JIS/UTF-8 Auto-Detection · Wareki/ISO Date Normalization · Accounting Currency Parser · Verifiable Audit Trail*

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://python.org)
[![Excel Ready](https://img.shields.io/badge/Output-Formatted%20Excel%20.xlsx-success.svg)](#excel-formatting)

[**日本語ドキュメント (README_JA.md)**](README_JA.md) • [**Architecture**](#architecture) • [**Test Results**](#test-results)

</div>

---

## 1. Problem & Business Context
Enterprise sales, procurement, and inventory records in Japan are routinely exported from disparate legacy ERPs, POS terminals, and manual Excel templates. Common data corruptions include:
* Full-width alphanumeric characters (`１２３`, `ＡＢＣ`) and ideographic spaces (`　`).
* Inconsistent Japanese era (Wareki `令和6年`) vs Gregorian dates (`2026/11/01`).
* Accounting currency formatting (`¥12,800`, `▲1,200` negative amounts).
* Encoding discrepancies (CP932/Shift-JIS vs UTF-8).
* Inadvertent duplicate submissions.

**Business-Data-Cleaner** is a turnkey Python processing pipeline that sanitizes input data, computes accurate KPI totals, flags anomalies with line-level audit reasons, and generates formatted, executive-ready Excel workbooks.

---

## 2. Core Capabilities
* **Zenkaku/Hankaku Normalization**: NFKC character folding with custom full-width space handling.
* **Flexible Date Parser**: Standardizes Japanese era (`令和`), slash, and dot formats into ISO `YYYY-MM-DD`. Validates calendar bounds (e.g. catches non-existent dates like `2026/11/31`).
* **Financial Amount Parsing**: Correctly handles currency symbols (`¥`, `円`), thousands separators, and Japanese accounting negative indicators (`▲` / `△`).
* **Audited Multi-Sheet Excel Output**:
  * `店舗別集計`: Executive summary of revenue per store with formatted totals.
  * `整形済み明細`: Sanitized transaction line items with frozen header panes.
  * `要確認データ`: Rejected rows cataloging original line number, full record JSON, and explicit diagnostic reasons.

---

## 3. Quickstart & Verification
```bash
python tests/test_cleaner.py
```
Output:
```text
=== CLEANING RESULTS SUMMARY ===
Total Input Rows: 10
Valid Clean Rows: 7
Flagged Anomaly Rows: 3
Totals by Group: {'名古屋': -1200, '大阪': 5500, '東京': 29600, '横浜': 21900}
[PASS] Excel generated successfully at: output/cleaned_sales.xlsx
```

---

## 4. License
MIT License.
