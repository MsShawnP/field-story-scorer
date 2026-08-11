"""Fidelity check: committed sample reports must match current code output.

The reports in ``samples/output/`` are portfolio artifacts shown to prospects.
They silently went stale once -- frozen at v2.2.0 while the code moved on to
v2.3.x -- because nothing regenerated them and compared against the committed
copies. This test closes that gap: it regenerates each HTML sample through the
real CLI path (the exact code that produces the committed files) and asserts a
byte-for-byte match, ignoring only the embedded generation timestamp.

If this fails after an *intentional* report change, the samples are simply out
of date -- regenerate and commit them:

    datascope samples/input/sample_sales.xlsx --format html --output-dir samples/output
    datascope samples/input/sample_mixed_types.xlsx --format html --output-dir samples/output

Only the two HTML samples are checked. PDFs embed nondeterministic ids and are
not text-diffable; the annotated Excel sample is a binary workbook.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from datascope.cli import main

_SAMPLES = Path(__file__).resolve().parent.parent / "samples"

# Report timestamp: "%Y-%m-%d %H:%M", e.g. "2026-07-28 23:19". The only
# nondeterministic content in the HTML output (html.py uses datetime.now()).
_TIMESTAMP_RE = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")

_HTML_SAMPLE_STEMS = ["sample_sales", "sample_mixed_types"]


def _normalize(html: str) -> str:
    return _TIMESTAMP_RE.sub("<TIMESTAMP>", html)


@pytest.mark.parametrize("stem", _HTML_SAMPLE_STEMS)
def test_committed_html_sample_is_current(stem: str, tmp_path: Path, capsys):
    committed = _SAMPLES / "output" / f"{stem}_diagnostic.html"
    input_file = _SAMPLES / "input" / f"{stem}.xlsx"
    assert committed.exists(), f"missing committed sample: {committed}"
    assert input_file.exists(), f"missing sample input: {input_file}"

    main([str(input_file), "--format", "html", "--output-dir", str(tmp_path)])
    capsys.readouterr()  # discard the CLI stdout summary

    regenerated = tmp_path / f"{stem}_diagnostic.html"
    assert regenerated.exists(), "CLI did not produce the HTML report"

    assert _normalize(regenerated.read_text(encoding="utf-8")) == _normalize(
        committed.read_text(encoding="utf-8")
    ), (
        f"{committed.name} is stale -- current code produces different output. "
        f"Regenerate it: datascope samples/input/{stem}.xlsx --format html "
        f"--output-dir samples/output"
    )
