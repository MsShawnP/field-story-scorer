# datascope

[![PyPI version](https://img.shields.io/pypi/v/datascope-dq)](https://pypi.org/project/datascope-dq/)
[![CI](https://github.com/MsShawnP/datascope/actions/workflows/ci.yml/badge.svg)](https://github.com/MsShawnP/datascope/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/pypi/pyversions/datascope-dq)](https://pypi.org/project/datascope-dq/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/MsShawnP/datascope/blob/main/LICENSE)

**PyPI:** `pip install datascope-dq`

Data created upstream — by manufacturing teams entering UPCs, inventory staff assigning product codes, offshore developers choosing column types — silently breaks systems downstream. A product code with letters where EDI expects numbers. Fifteen "N/A" strings buried in 500 numeric rows that pandas silently drops, skewing every calculation by 3%.

datascope finds these problems, explains what's wrong in plain English, and tells you what to fix. It reads each cell's actual type (not what pandas infers), detects hidden quality issues, classifies their severity by downstream impact, and generates a professional diagnostic report.

---

## What It Finds

| Detection | Example | Severity |
|---|---|---|
| **Mixed types** | 485 numbers + 15 strings in a "numeric" column | Critical |
| **Sentinel values** | "N/A", "TBD", "pending" hiding in numeric data | Critical |
| **Missing values** | Flagged at ≥10% blank; Warning at ≥50%, Info below — aggregations silently exclude those rows | Warning / Info |
| **Leading-zero inconsistency** | "00123" alongside "456" — keys that won't match | Warning |
| **Mixed date formats** | "01/15/2026" and "2026-01-15" in the same column | Warning |
| **Suspected duplicate IDs** | 98% unique in an ID column — the other 2% will fan out joins | Warning |
| **Near-constant columns** | 1 distinct value across 10,000 rows | Info |

Each finding is expressed as **assumption vs. reality**: what the data *appears* to be vs. what it *actually contains*. Every finding includes a downstream impact explanation, a fix recommendation, and a prevention rule.

---

## Installation

```bash
pip install datascope-dq
```

For Parquet file support:

```bash
pip install datascope-dq[parquet]
```

Or install from source:

```bash
git clone https://github.com/MsShawnP/datascope.git
cd datascope
pip install -e .
```

---

## Usage

```bash
# Analyze an Excel file
datascope data.xlsx

# Analyze a CSV
datascope sales_export.csv

# Analyze a Parquet file (requires pyarrow)
datascope data.parquet

# Specify a sheet and output directory
datascope data.xlsx --sheet Revenue --output-dir ./client_reports
```

### Output Formats

```bash
# PDF report (default)
datascope data.xlsx

# Structured JSON for pipeline integration
datascope data.xlsx --format json

# Self-contained HTML report
datascope data.xlsx --format html

# Annotated Excel with highlighted problem cells
datascope data.xlsx --format annotated-excel

# PDF + JSON together
datascope data.xlsx --format both
```

### CLI Flags

```bash
# Quiet mode — exit code only (0 = no critical, 1 = critical findings)
datascope data.xlsx --quiet

# Verbose mode — full tracebacks on analyzer failures
datascope data.xlsx --verbose

# Limit row count (default: warn at 500K cells, abort at 5M)
datascope huge_file.csv --max-rows 100000
```

### Example Output

```
datascope: Analyzing sample_mixed_types.xlsx...
  200 rows x 4 columns

Found 2 findings:
  2 Critical  ########

Top critical findings:
  * revenue_mixed: However, 15 str values were found among 200 non-null values (the majority type covers 92.5%). Examples of unexpected values: 'N/A', 'N/A', 'N/A', and 2 more.
  * revenue_mixed: However, 7.5% of values (1 distinct sentinel string) are placeholder text rather than real data: 'N/A' (15 times).

Report saved: reports/sample_mixed_types_diagnostic.pdf
```

---

## The Report

Reports are structured for non-technical readers — no jargon, no composite scores, no unexplained metrics.

**Executive Summary** — overall health assessment, finding counts by severity, top critical issues highlighted.

**Findings by Severity** — each finding presented as a card:
- **Assumption**: what the data appears to be
- **Reality**: what it actually contains
- **Impact**: what breaks downstream
- **Recommended Fix**: what to do now
- **Prevention Rule**: what right looks like going forward

**Field Inventory** — summary table of all columns with their detected issue types and severity.

Findings are color-coded (red/amber/blue) and grouped by severity so readers know what to fix first.

---

## How It Works

Most tools let pandas (or the SQL driver, or Excel) decide column types. A column with 485 numbers and 15 strings becomes `float64` — the strings become `NaN`, the type problem disappears, and every downstream calculation is quietly wrong.

datascope reads each cell's actual Python type via openpyxl (for Excel) or raw-string inference (for CSV). This cell-level type detection is always on — there's no flag to enable it because skipping it defeats the purpose.

The analysis pipeline:

1. **Load** — read with cell-level type preservation (no silent coercion)
2. **Detect** — seven analyzers scan for type inconsistencies, sentinels, missing values, format issues, and cardinality anomalies
3. **Classify** — severity assigned by downstream impact (critical = silent data loss, warning = likely misinterpretation, info = worth noting)
4. **Compose** — plain-English narrative generated for each finding
5. **Report** — output as PDF, HTML, JSON, or annotated Excel

---

## Severity Model

| Level | Meaning | Examples |
|---|---|---|
| **Critical** | Silent data loss or incorrect calculations will occur | Mixed types in numeric columns; sentinel values pandas drops without warning |
| **Warning** | Key mismatches or misinterpretation likely | Leading-zero stripping; ambiguous date formats; duplicate IDs; high null rates |
| **Info** | Worth noting, no direct downstream breakage | Near-constant columns; moderate missing values |

---

## Project Structure

```
datascope/
├── loaders/            # Excel, CSV, and Parquet with cell-level type tracking
│   ├── excel.py        # openpyxl-based, preserves per-cell Python types
│   ├── csv_loader.py   # Raw string inference with regex-accelerated datetime detection
│   ├── parquet.py      # Arrow schema → Python type mapping (optional pyarrow)
│   └── base.py         # Extension-based dispatch
├── analyzers/          # Seven detectors, each returns list[Finding]
│   ├── type_consistency.py
│   ├── sentinel.py
│   ├── format_check.py # Leading zeros + mixed dates
│   ├── cardinality.py  # Near-constant + duplicate IDs
│   └── missing_values.py
├── findings/           # Severity classifier + NL template engine
│   ├── severity.py     # Impact-based classification rules
│   ├── templates.py    # Plain-English templates per finding type
│   ├── composer.py     # Template dispatch
│   └── pipeline.py     # classify → compose → sort
├── reports/
│   ├── pdf.py          # Professional PDF with reportlab
│   ├── html.py         # Self-contained HTML with inline CSS
│   └── annotated_excel.py  # Highlighted cells + findings sheet
└── cli.py              # argparse CLI, pipeline orchestration
```

---

## Requirements

- Python 3.10+
- pandas >= 2.0
- openpyxl >= 3.1
- reportlab >= 4.0
- pyarrow >= 12.0 *(optional, for Parquet support)*

---

## License

MIT

---
Built by [Lailara LLC](https://lailarallc.com) — data hygiene and analytics consulting for specialty food brands scaling into national retail.
