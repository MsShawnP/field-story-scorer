# field-story-scorer

Built for the reality of consulting work: you get handed a mystery `.xlsx`, you need to know which fields are trustworthy before you build anything on top of them.

field-story-scorer is a command-line tool that profiles every column in an Excel file, scores each field across five data quality dimensions, and outputs formatted Excel and PDF reports ready for stakeholder review.

---

## Features

- **Five scoring dimensions** with a weighted composite score per field
- **Excel report** with color-scale conditional formatting, frozen panes, and data bars
- **PDF report** in landscape format with color-coded score cells
- **`--strict-types` mode** — catches mixed-type columns that pandas silently coerces (see below)
- Single-file tool, minimal dependencies, no configuration required

---

## Scoring Methodology

Each field is scored 0–1 across five dimensions. Scores are combined into a weighted composite.

| Dimension | Weight | Description |
|---|---|---|
| **Completeness** | 30% | Ratio of non-null values to total rows |
| **Type Consistency** | 25% | Proportion of values sharing the majority Python type |
| **Cardinality** | 15% | Unique value ratio (identifies constants and pure ID columns) |
| **Distribution** | 15% | Coefficient of variation (numeric) or normalized entropy (categorical) |
| **Correlation** | 15% | Mean absolute Pearson correlation with other numeric fields |

### Composite Score Thresholds

| Score | Meaning |
|---|---|
| ≥ 0.75 | Field is analytically strong — high completeness, consistent types, meaningful variance |
| 0.45 – 0.74 | Usable with caveats — review dimension breakdown for weak points |
| < 0.45 | Problematic — likely sparse, constant, or type-inconsistent |

---

## Output Reports

Both formats contain the same four sections:

**Tab 1 — Field Rankings**
All fields sorted by composite score. Includes field type classification, null count, and unique value count.

**Tab 2 — Field Profiles**
Per-field breakdown of all five dimension scores. Data bars in Excel, color-coded cells in PDF.

**Tab 3 — Chart Recommendations**
Suggested visualization types derived from inferred field type, cardinality, and distribution score. Flags identifier columns and near-constant fields.

**Tab 4 — Correlation Matrix**
Pearson correlation matrix for all numeric fields. Diverging red/white/green color scale.

---

## Installation

```bash
git clone https://github.com/MsShawnP/field-story-scorer.git
cd field-story-scorer
pip install -r requirements.txt
```

---

## Usage

```bash
# Basic run — first sheet, reports saved to ./reports/
python scorer.py --input data.xlsx

# Specify sheet by name or index
python scorer.py --input data.xlsx --sheet Sheet1
python scorer.py --input data.xlsx --sheet 0

# Custom output directory
python scorer.py --input data.xlsx --sheet Sales --output-dir ./client_reports

# Strict type detection (see below)
python scorer.py --input data.xlsx --sheet Sales --strict-types
```

Output files are named `{input_stem}_field_report.xlsx` and `{input_stem}_field_report.pdf`. Strict-types runs append `_strict` to avoid overwriting standard output.

---

## `--strict-types` Flag

By default, pandas infers column dtypes on load. A column containing 485 floats and 15 cells with the string `"N/A"` will be read as `float64` — the strings become `NaN`, the type inconsistency disappears, and the column scores nearly identically to a clean numeric field.

`--strict-types` bypasses pandas inference entirely, reading each cell's native Python type via openpyxl. The same column now reveals its actual composition:

```
Standard mode:   revenue_mixed   score=0.9775   type=numeric_continuous
Strict mode:     revenue_mixed   score=0.8170   type=identifier   mix=[numeric:185, str:15]
```

The `type_mix` column is added to the Field Rankings and Chart Recommendations tabs in strict-mode output, showing the exact type breakdown for every field.

**When to use it:** Any time a dataset has been manually edited in Excel, exported from a system that emits sentinel strings (`"N/A"`, `"TBD"`, `"—"`, `"NULL"`), or assembled from multiple sources. Standard mode is faster and sufficient for clean, system-generated data.

---

## See It In Action

Real reports produced by running the scorer on the bundled sample inputs — committed to the repo so you can preview the output without installing anything. See [samples/README.md](samples/README.md) for the full breakdown.

**Inputs**
- [samples/input/sample_sales.xlsx](samples/input/sample_sales.xlsx) — 500-row clean dataset
- [samples/input/sample_mixed_types.xlsx](samples/input/sample_mixed_types.xlsx) — 200-row dataset with 15 sentinel-string cells in `revenue_mixed`

**Outputs**
- Clean dataset:
  [xlsx](samples/output/sample_sales_field_report.xlsx) ·
  [pdf](samples/output/sample_sales_field_report.pdf)
- Mixed-types, standard mode (string contamination hidden — `revenue_mixed` scores 0.9775):
  [xlsx](samples/output/sample_mixed_types_field_report.xlsx) ·
  [pdf](samples/output/sample_mixed_types_field_report.pdf)
- Mixed-types, `--strict-types` (string contamination exposed — `revenue_mixed` drops to 0.8170, type flips to `identifier`):
  [xlsx](samples/output/sample_mixed_types_field_report_strict.xlsx) ·
  [pdf](samples/output/sample_mixed_types_field_report_strict.pdf)

The two mixed-types reports are the headline — same input, scored both ways. Open them side-by-side to see exactly what `--strict-types` catches.

---

## Sample Data

Pre-generated samples live in [samples/](samples/) (see [samples/README.md](samples/README.md)). To regenerate them:

```bash
python generate_sample.py
```

This writes two files into `samples/input/`:
- `sample_sales.xlsx` — 500-row dataset with a range of field types (categorical, numeric, boolean, sparse, constant)
- `sample_mixed_types.xlsx` — 200-row dataset with 15 genuine string cells in an otherwise numeric column, designed to demonstrate `--strict-types`

---

## Field Type Classification

The tool infers a field type for each column used in chart recommendations and report labeling.

| Type | Criteria |
|---|---|
| `numeric_continuous` | Numeric dtype, > 20 unique values |
| `numeric_discrete` | Numeric dtype, ≤ 20 unique values |
| `categorical_low` | Object/string, ≤ 20 unique values |
| `categorical_high` | Object/string, > 20 unique values, < 90% unique ratio |
| `identifier` | > 90% unique values — likely a key or ID column |
| `boolean` | Boolean dtype |
| `datetime` | Datetime dtype |
| `unknown` | No non-null values |

---

## Project Structure

```
field-story-scorer/
├── scorer.py               # Main CLI tool — single file, no submodules
├── generate_sample.py      # Generates the bundled sample inputs
├── requirements.txt
├── samples/                # Committed sample inputs and rendered outputs
│   ├── README.md
│   ├── input/
│   └── output/
└── README.md
```

---

## Requirements

- Python 3.10+
- pandas ≥ 2.0
- openpyxl ≥ 3.1
- reportlab ≥ 4.0
- numpy ≥ 1.24
- scipy ≥ 1.10

---

## License

MIT
