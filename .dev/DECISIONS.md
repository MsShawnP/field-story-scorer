# Decisions

## 2026-05-15: PyPI distribution name → datascope-dq

`datascope` is taken on PyPI by an ETH Zurich ML tool (Shapley-value data importance scores). Chose `datascope-dq` as the distribution name — short, communicates "data quality." The Python import stays `import datascope`. All install docs, README, and branding reference `pip install datascope-dq`.

Alternatives considered: `datascope-diagnostics` (too verbose), `datascope-cli` (undersells the library use case).

## 2026-05-15: FindingType sub-types promoted to first-class enum values

Promoted LEADING_ZEROS, MIXED_DATES, NEAR_CONSTANT, DUPLICATE_IDS from implicit evidence-key conventions to explicit `FindingType` enum members. Removed the old generic `FORMAT_INCONSISTENCY` and `CARDINALITY_ANOMALY` types.

This eliminated 6 dispatch sites across severity.py, composer.py, and pdf.py that inspected magic dict keys like `"leading_zero_count" in evidence`. New finding types must be added as enum values with corresponding template, severity rule, and PDF label — not smuggled through evidence keys.

## 2026-05-15: HTML reports use inline CSS, no Jinja2

HTML report is a single self-contained file with all CSS inlined. No template engine dependency — just f-strings and `html.escape()`. This keeps the dependency footprint minimal (Jinja2 comes via pandas but we don't rely on it) and means the HTML file works offline, in email attachments, or anywhere a browser exists.

## 2026-05-15: Parquet support as optional extra, not core dependency

pyarrow is large (~200MB installed). Making it a core dependency would bloat install for users who only work with CSV/Excel. Instead, it's behind `pip install datascope-dq[parquet]`. The loader raises a clear `ImportError` with install instructions if pyarrow is missing.

## 2026-05-15: PEP 639 license format — drop legacy classifier

Modern setuptools (isolated build env) rejects the `License :: OSI Approved :: MIT License` classifier when `license = "MIT"` is also present. Removed the classifier, keeping only the PEP 639 `license` string field. Future classifiers should not include license entries.

## 2026-05-22: defusedxml removed; openpyxl handles XML safety internally

- **Why:** `defusedxml` was listed as a dependency but never imported. It only works when explicitly imported before XML parsing (monkey-patches stdlib). Modern openpyxl (3.1+) handles XML parsing safely without it. The dependency added install weight with no protection.
- **Scope:** datascope dependency management
- **Do not:** Re-add defusedxml unless datascope starts parsing user-supplied XML outside of openpyxl (e.g., raw lxml usage).

## 2026-07-27: UI review targets the generated HTML reports served locally

- **Why:** datascope has no live web app — its "UI" is the self-contained HTML diagnostic report. The ui-review tool runs against `samples/output/*.html` served with `python -m http.server`, configured by `review.yaml`. `design_system: true` is enabled even though this isn't a `lailarallc.com` deploy, because the reports are built to the Lailara Design System and the brand-token checks are exactly what matters.
- **Scope:** How `/ui-review` is run for datascope (and similar report-generating tools).
- **Do not:** Treat a UI-review "heading typeface" warning as real without checking — it measures the container element, not the `h1` (the `h1` correctly uses Playfair). Do not commit the `screenshots/` the tool regenerates (gitignored); do keep `review.yaml`.

## 2026-05-16: Stay in the file-audit niche; do not compete with pipeline tools

- **Why:** GX owns rules, Pandera owns schemas, Soda owns databases, ydata owns stats. datascope's moat is cell-level detection + professional narrative reports for non-technical readers. Competing on their turf dilutes the positioning.
- **Scope:** All future feature decisions for datascope
- **Do not:** Add custom validation rules, database connectors, statistical profiling, Polars backend, drift detection, or web UI/SaaS

## 2026-07-31: Committed samples must match current code output; HTML enforced by a test

- **Why:** `samples/output/` files are portfolio artifacts shown to prospects. They silently went stale twice (frozen at v2.2.0, then again at v2.3.3) because nothing regenerated and compared them against current code.
- **Scope:** All committed sample outputs. `tests/test_samples_fidelity.py` regenerates the two HTML samples via `cli.main()` and asserts a byte match (timestamp-normalized). Any intentional change to report rendering, `_palette` narrative/colors, or the package version must be accompanied by regenerated samples in the same change.
- **Do not:** Ship a report/palette/version change without regenerating `samples/output/`. Do not attempt to byte-diff the PDF or annotated-Excel samples in a test — they embed nondeterministic metadata; regenerate those by hand.
