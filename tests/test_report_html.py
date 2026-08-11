"""Tests for datascope.reports.html -- HTML report generation."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from datascope.findings.composer import compose_finding
from datascope.findings.pipeline import process_findings
from datascope.findings.severity import classify_severity
from datascope.models import Finding, FindingType
from datascope.reports._palette import CANVAS_HEX, LONDON_95_HEX
from datascope.reports.html import _e, _health_assessment, write_html

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DEFAULT_METADATA = {
    "filename": "test_data.csv",
    "row_count": 1000,
    "column_count": 12,
}


def _make_finding(
    finding_type: FindingType,
    evidence: dict,
    field_name: str = "test_col",
) -> Finding:
    f = Finding(field_name=field_name, finding_type=finding_type, evidence=evidence)
    f.severity = classify_severity(f)
    compose_finding(f)
    return f


def _critical_finding(field_name: str = "revenue") -> Finding:
    return _make_finding(
        FindingType.TYPE_INCONSISTENCY,
        {
            "majority_type": "numeric",
            "majority_count": 485,
            "minority_types": [
                {"type_name": "str", "count": 15, "examples": ["N/A", "pending", "TBD"]},
            ],
            "total_non_null": 500,
            "majority_pct": 97.0,
        },
        field_name=field_name,
    )


def _warning_finding(field_name: str = "zip_code") -> Finding:
    return _make_finding(
        FindingType.LEADING_ZEROS,
        {
            "leading_zero_count": 10,
            "no_leading_zero_count": 40,
            "examples_with_zeros": ["00123", "00456", "00789"],
            "examples_without_zeros": ["123", "456"],
            "total_checked": 50,
        },
        field_name=field_name,
    )


def _info_finding(field_name: str = "status") -> Finding:
    return _make_finding(
        FindingType.NEAR_CONSTANT,
        {
            "unique_count": 2,
            "total_count": 5000,
            "uniqueness_ratio": 0.0004,
            "top_values": [
                {"value": "Active", "count": 4990},
                {"value": "Inactive", "count": 10},
            ],
        },
        field_name=field_name,
    )


# ---------------------------------------------------------------------------
# Unit tests for _e (HTML escape)
# ---------------------------------------------------------------------------

class TestHtmlEscape:
    def test_escapes_angle_brackets(self):
        assert _e("<script>alert('xss')</script>") == "&lt;script&gt;alert(&#x27;xss&#x27;)&lt;/script&gt;"

    def test_none_returns_empty(self):
        assert _e(None) == ""

    def test_plain_text_unchanged(self):
        assert _e("hello world") == "hello world"

    def test_ampersand_escaped(self):
        assert _e("A & B") == "A &amp; B"


# ---------------------------------------------------------------------------
# Unit tests for _health_assessment
# ---------------------------------------------------------------------------

class TestHealthAssessment:
    def test_critical_findings(self):
        findings = [_critical_finding(), _critical_finding("cost")]
        result = _health_assessment(findings)
        assert "2 critical issues" in result
        assert "silent data loss" in result

    def test_single_critical(self):
        result = _health_assessment([_critical_finding()])
        assert "1 critical issue" in result

    def test_warnings_only(self):
        result = _health_assessment([_warning_finding()])
        assert "No critical issues" in result
        assert "1 warning" in result

    def test_info_only(self):
        result = _health_assessment([_info_finding()])
        assert "1 informational observation" in result
        assert "good shape" in result

    def test_multiple_info(self):
        result = _health_assessment([_info_finding(), _info_finding("region")])
        assert "2 informational observations" in result

    def test_no_findings(self):
        result = _health_assessment([])
        assert "No data quality issues were detected" in result

    def test_critical_takes_precedence(self):
        findings = [_critical_finding(), _warning_finding(), _info_finding()]
        result = _health_assessment(findings)
        assert "critical" in result


# ---------------------------------------------------------------------------
# Happy path -- mixed findings
# ---------------------------------------------------------------------------

class TestHappyPathMixedFindings:
    @pytest.fixture()
    def html_path(self, tmp_path: Path) -> Path:
        findings = [
            _critical_finding("revenue"),
            _critical_finding("cost"),
            _warning_finding("zip_code"),
            _info_finding("status"),
        ]
        findings = process_findings(findings)
        out = tmp_path / "report.html"
        write_html(findings, _DEFAULT_METADATA, out)
        return out

    def test_file_created(self, html_path: Path):
        assert html_path.exists()

    def test_file_non_zero_size(self, html_path: Path):
        assert html_path.stat().st_size > 0

    def test_valid_html_structure(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert content.startswith("<!DOCTYPE html>")
        assert "</html>" in content

    def test_contains_title(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "Data Quality Diagnostic" in content

    def test_contains_filename(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "test_data.csv" in content

    def test_contains_row_count(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "1000" in content

    def test_contains_all_severity_sections(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "Critical Findings" in content
        assert "Warning Findings" in content
        assert "Info Findings" in content

    def test_contains_field_names(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "revenue" in content
        assert "zip_code" in content
        assert "status" in content

    def test_contains_severity_badges(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "Critical" in content
        assert "Warning" in content
        assert "Info" in content

    def test_contains_finding_count(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "4 findings" in content

    def test_palette_colors_present(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "#a80d08" in content
        assert "#a05a1a" in content
        assert "#3348a8" in content

    def test_lailara_design_tokens(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")

        # Card/panel surfaces render the London-95 token -- pin selector -> token
        # so a revert to bare #ffffff (or a rename that misses a surface) fails.
        for selector in (".summary-card", ".health", ".finding-card"):
            assert re.search(
                rf"{re.escape(selector)}\s*\{{[^}}]*background:\s*{re.escape(LONDON_95_HEX)}",
                content,
            ), f"{selector} does not use the London-95 surface token {LONDON_95_HEX}"

        # Page background and table base render the canvas token.
        for selector in ("body", "table"):
            assert re.search(
                rf"(?:^|\s){re.escape(selector)}\s*\{{[^}}]*background:\s*{re.escape(CANVAS_HEX)}",
                content,
            ), f"{selector} does not use the canvas token {CANVAS_HEX}"

        # Zebra striping: even rows take London-95 over the canvas table base.
        # The two tokens must differ or the stripe collapses to a flat fill.
        assert CANVAS_HEX != LONDON_95_HEX
        assert re.search(
            rf"tr:nth-child\(even\)\s*\{{[^}}]*background:\s*{re.escape(LONDON_95_HEX)}",
            content,
        ), "zebra even-row does not use the London-95 token"

        assert "#1f2e7a" in content
        assert "@font-face" in content, (
            "No @font-face block: the woff2 files did not ship with the package. "
            "Naming the family in the font stack is not evidence they loaded — "
            "check [tool.setuptools.package-data] covers datascope.reports.fonts."
        )
        assert "Playfair Display" in content
        assert "Source Sans 3" in content


# ---------------------------------------------------------------------------
# Zero findings
# ---------------------------------------------------------------------------

class TestZeroFindings:
    @pytest.fixture()
    def html_path(self, tmp_path: Path) -> Path:
        out = tmp_path / "empty.html"
        write_html([], _DEFAULT_METADATA, out)
        return out

    def test_file_created(self, html_path: Path):
        assert html_path.exists()

    def test_valid_html(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "<!DOCTYPE html>" in content

    def test_health_says_clean(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "No data quality issues were detected" in content

    def test_zero_counts(self, html_path: Path):
        content = html_path.read_text(encoding="utf-8")
        assert "0 findings" in content or "finding" not in content.split("footer")[0]


# ---------------------------------------------------------------------------
# Single finding
# ---------------------------------------------------------------------------

class TestSingleFinding:
    def test_single_critical(self, tmp_path: Path):
        findings = process_findings([_critical_finding()])
        out = tmp_path / "single.html"
        write_html(findings, _DEFAULT_METADATA, out)
        content = out.read_text(encoding="utf-8")
        assert out.exists()
        assert "1 finding" in content

    def test_single_info(self, tmp_path: Path):
        findings = process_findings([_info_finding()])
        out = tmp_path / "single_info.html"
        write_html(findings, _DEFAULT_METADATA, out)
        assert out.exists()


# ---------------------------------------------------------------------------
# Special characters in field names (XSS safety)
# ---------------------------------------------------------------------------

class TestXssSafety:
    def test_html_injection_in_field_name(self, tmp_path: Path):
        f = _critical_finding('<script>alert("xss")</script>')
        findings = process_findings([f])
        out = tmp_path / "xss.html"
        write_html(findings, _DEFAULT_METADATA, out)
        content = out.read_text(encoding="utf-8")
        assert "<script>" not in content
        assert "&lt;script&gt;" in content

    def test_html_injection_in_filename(self, tmp_path: Path):
        metadata = {**_DEFAULT_METADATA, "filename": '<img src=x onerror="alert(1)">'}
        out = tmp_path / "xss_meta.html"
        write_html([], metadata, out)
        content = out.read_text(encoding="utf-8")
        assert "<img " not in content


# ---------------------------------------------------------------------------
# Auto-create parent directory
# ---------------------------------------------------------------------------

class TestAutoCreateDirectory:
    def test_creates_nested_directory(self, tmp_path: Path):
        out = tmp_path / "sub" / "deep" / "report.html"
        write_html([], _DEFAULT_METADATA, out)
        assert out.exists()


# ---------------------------------------------------------------------------
# Metadata edge cases
# ---------------------------------------------------------------------------

class TestMetadataEdgeCases:
    def test_missing_filename(self, tmp_path: Path):
        out = tmp_path / "no_name.html"
        write_html([], {}, out)
        content = out.read_text(encoding="utf-8")
        assert "unknown" in content

    def test_missing_row_count(self, tmp_path: Path):
        out = tmp_path / "no_rows.html"
        write_html([], {"filename": "test.csv"}, out)
        content = out.read_text(encoding="utf-8")
        assert "?" in content


# ---------------------------------------------------------------------------
# Long text content
# ---------------------------------------------------------------------------

class TestLongTextContent:
    def test_long_narrative_fields(self, tmp_path: Path):
        f = _critical_finding()
        f.assumption = "Very long assumption. " * 50
        f.reality = "Very long reality. " * 50
        f.impact = "Very long impact. " * 50
        f.fix_recommendation = "Very long fix. " * 50
        f.prevention_rule = "Very long prevention. " * 50
        out = tmp_path / "long.html"
        write_html([f], _DEFAULT_METADATA, out)
        assert out.exists()
        assert out.stat().st_size > 0


# ---------------------------------------------------------------------------
# Field inventory table
# ---------------------------------------------------------------------------

class TestFieldInventory:
    def test_inventory_lists_all_findings(self, tmp_path: Path):
        findings = process_findings([
            _critical_finding("amount"),
            _warning_finding("code"),
            _info_finding("region"),
        ])
        out = tmp_path / "inventory.html"
        write_html(findings, _DEFAULT_METADATA, out)
        content = out.read_text(encoding="utf-8")
        assert "amount" in content
        assert "code" in content
        assert "region" in content
        assert "Field Inventory" in content

    def test_finding_type_labels(self, tmp_path: Path):
        findings = process_findings([_critical_finding()])
        out = tmp_path / "types.html"
        write_html(findings, _DEFAULT_METADATA, out)
        content = out.read_text(encoding="utf-8")
        assert "Type Inconsistency" in content
