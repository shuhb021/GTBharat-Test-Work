"""
FAR Automation Tool — Validation Layer
Pre-export validation checks to catch common data inconsistencies
before they produce silently incorrect or zero-valued ratios.
"""

import logging
import math

logger = logging.getLogger(__name__)


class ValidationReport:
    """Container for validation results."""

    def __init__(self):
        self.passed = []
        self.failed = []
        self.warnings = []
        self.auto_corrected = []

    def add_pass(self, check, detail=''):
        msg = f"✓ {check}" + (f": {detail}" if detail else '')
        self.passed.append(msg)
        logger.info("VALIDATION PASS - %s", f"{check}" + (f": {detail}" if detail else ''))

    def add_fail(self, check, detail=''):
        msg = f"✗ {check}" + (f": {detail}" if detail else '')
        self.failed.append(msg)
        logger.error("VALIDATION FAIL - %s", f"{check}" + (f": {detail}" if detail else ''))

    def add_warning(self, check, detail=''):
        msg = f"⚠ {check}" + (f": {detail}" if detail else '')
        self.warnings.append(msg)
        logger.warning("VALIDATION WARN - %s", f"{check}" + (f": {detail}" if detail else ''))

    def add_auto_corrected(self, check, detail=''):
        msg = f"↺ {check}" + (f": {detail}" if detail else '')
        self.auto_corrected.append(msg)
        logger.info("VALIDATION AUTO-CORRECTED - %s", f"{check}" + (f": {detail}" if detail else ''))

    @property
    def has_failures(self):
        return len(self.failed) > 0

    def summary_text(self):
        parts = []
        if self.passed:
            parts.append(f"✓ {len(self.passed)} passed")
        if self.warnings:
            parts.append(f"⚠ {len(self.warnings)} warnings")
        if self.failed:
            parts.append(f"✗ {len(self.failed)} failed")
        if self.auto_corrected:
            parts.append(f"↺ {len(self.auto_corrected)} auto-corrected")
        return "  |  ".join(parts) if parts else "No checks run"

    def detailed_report(self):
        lines = ["=" * 60, "FAR VALIDATION REPORT", "=" * 60]
        if self.passed:
            lines.append("\nPASSED CHECKS:")
            lines.extend(f"  {m}" for m in self.passed)
        if self.warnings:
            lines.append("\nWARNINGS:")
            lines.extend(f"  {m}" for m in self.warnings)
        if self.failed:
            lines.append("\nFAILED CHECKS:")
            lines.extend(f"  {m}" for m in self.failed)
        if self.auto_corrected:
            lines.append("\nAUTO-CORRECTED:")
            lines.extend(f"  {m}" for m in self.auto_corrected)
        lines.append("=" * 60)
        return "\n".join(lines)


def validate_far_data(bs_data, pl_data, bs_summary, pl_summary, ratios):
    """
    Run all pre-export validation checks.

    Args:
        bs_data:    list of raw BS row dicts
        pl_data:    list of raw PL row dicts
        bs_summary: dict from extract_bs_summary
        pl_summary: dict from extract_pl_summary
        ratios:     list of computed ratio dicts

    Returns:
        ValidationReport
    """
    report = ValidationReport()

    _check_balance_sheet_identity(bs_data, bs_summary, report)
    _check_revenue_nonzero(pl_summary, ratios, report)
    _check_receivables_nonzero(bs_summary, ratios, report)
    _check_ratios_not_silently_zero(ratios, bs_summary, pl_summary, report)
    _check_missing_pl_items(pl_summary, report)
    _check_missing_bs_items(bs_summary, report)
    _check_data_row_counts(bs_data, pl_data, report)

    logger.info("Validation complete: %d passed, %d warnings, %d failed, %d auto-corrected",
                len(report.passed), len(report.warnings),
                len(report.failed), len(report.auto_corrected))
    
    safe_report = report.detailed_report().encode('ascii', 'replace').decode('ascii')
    logger.info(safe_report)

    return report


def _safe(d, key):
    v = d.get(key, 0)
    try:
        f = float(v)
        return 0.0 if (math.isnan(f) or math.isinf(f)) else f
    except (TypeError, ValueError):
        return 0.0


def _check_balance_sheet_identity(bs_data, bs_summary, report):
    """Check: Total Assets ≈ Total Equity + Total Liabilities."""
    total_assets = _safe(bs_summary, 'total_assets_cy')
    total_equity = _safe(bs_summary, 'total_equity_cy')
    total_cl = _safe(bs_summary, 'total_cl_cy')
    total_ncl = _safe(bs_summary, 'total_ncl_cy')

    if total_assets == 0 and total_equity == 0:
        report.add_warning(
            "Balance Sheet identity",
            "Total Assets and Total Equity are both 0 — cannot verify BS identity. "
            "Check that the BS sheet has numeric values in the correct columns.")
        return

    total_liabilities = total_cl + total_ncl
    rhs = total_equity + total_liabilities

    if total_assets == 0:
        # Try to derive from parsed rows
        for row in bs_data:
            p = row.get('particulars', '').lower()
            if 'total assets' in p:
                total_assets = float(row.get('cy', 0) or 0)
                break

    if total_assets == 0:
        report.add_warning(
            "Balance Sheet identity",
            "Could not find 'Total Assets' row — skipping identity check")
        return

    tolerance = max(abs(total_assets) * 0.001, 1.0)  # 0.1% or 1 unit
    diff = abs(total_assets - rhs)

    if diff <= tolerance:
        report.add_pass(
            "Balance Sheet identity",
            f"Assets={total_assets:,.0f} ≈ Equity+Liabilities={rhs:,.0f} "
            f"(diff={diff:,.0f}, within tolerance {tolerance:,.0f})")
    else:
        report.add_fail(
            "Balance Sheet identity",
            f"Assets={total_assets:,.0f} ≠ Equity+Liabilities={rhs:,.0f} "
            f"(diff={diff:,.0f} exceeds tolerance {tolerance:,.0f}). "
            "Possible missing notes or mis-mapped rows.")


def _check_revenue_nonzero(pl_summary, ratios, report):
    """Check: Revenue > 0 when profitability ratios are computed."""
    revenue = _safe(pl_summary, 'revenue_ops_cy')

    if revenue > 0:
        report.add_pass(
            "Revenue from Operations > 0",
            f"Revenue = {revenue:,.0f}")
    else:
        report.add_fail(
            "Revenue from Operations is ZERO",
            "Net Profit Ratio and Debtor Turnover Ratio will be 0. "
            "Verify P&L parsing: check sheet name, date columns, "
            "and 'Revenue from Operations' row label in source file.")

        # Mark affected ratios
        for r in ratios:
            if r.get('key') in ('Net Profit Ratio', 'Debtor Turnover Ratio', 'GCA Days'):
                if r.get('cy', None) == 0:
                    report.add_warning(
                        f"Ratio '{r['key']}' = 0",
                        "Caused by zero revenue denominator — fix revenue mapping first")


def _check_receivables_nonzero(bs_summary, ratios, report):
    """Check: Trade Receivables > 0 when Debtor Turnover is computed."""
    debtors_cy = _safe(bs_summary, 'debtors_cy')
    debtors_py = _safe(bs_summary, 'debtors_py')
    avg = (debtors_cy + debtors_py) / 2.0

    if avg > 0:
        report.add_pass(
            "Average Trade Receivables > 0",
            f"CY={debtors_cy:,.0f}, PY={debtors_py:,.0f}, Avg={avg:,.0f}")
    else:
        if debtors_cy == 0 and debtors_py == 0:
            report.add_fail(
                "Trade Receivables are ZERO",
                "Debtor Turnover Ratio denominator is 0. "
                "Verify BS parsing: the 'Trade Receivables' row may contain "
                "a Note reference but no numeric value — check Notes sheet instead.")
        else:
            report.add_warning(
                "Average Trade Receivables uses only closing debtors",
                f"CY={debtors_cy:,.0f}, PY={debtors_py:,.0f}")


def _check_ratios_not_silently_zero(ratios, bs_summary, pl_summary, report):
    """
    For each ratio, check if it's 0 when it shouldn't be.
    A ratio of 0 is only correct if the numerator is genuinely 0.
    """
    numerator_keys = {
        'Current Ratio': ('total_ca_cy', bs_summary),
        'Net Profit Ratio': ('pat_cy', pl_summary),
        'Debtor Turnover Ratio': ('revenue_ops_cy', pl_summary),
        'Return on Equity (ROE)': ('pbt_cy', pl_summary),
        'Return on Capital Employed (ROCE)': ('pbt_cy', pl_summary),
    }

    for ratio in ratios:
        key = ratio.get('key', '')
        cy_val = ratio.get('cy', 0)

        if key in numerator_keys:
            num_key, summary_dict = numerator_keys[key]
            numerator = _safe(summary_dict, num_key)

            if cy_val == 0 and numerator != 0:
                report.add_fail(
                    f"Ratio '{key}' = 0 but numerator ≠ 0",
                    f"Numerator ({num_key}) = {numerator:,.2f} — "
                    "this indicates the denominator is 0. Check source mapping.")
            elif cy_val == 0 and numerator == 0:
                report.add_warning(
                    f"Ratio '{key}' = 0",
                    f"Both numerator and denominator are 0 — "
                    f"verify source data has actual values for '{key}'")
            else:
                report.add_pass(
                    f"Ratio '{key}'",
                    f"CY={cy_val:.4f} (non-zero)")


def _check_missing_pl_items(pl_summary, report):
    """Check that key P&L items were successfully extracted."""
    checks = [
        ('revenue_ops_cy', 'Revenue from Operations'),
        ('pbt_cy', 'Profit Before Tax'),
        ('pat_cy', 'Profit After Tax'),
    ]
    for key, label in checks:
        val = _safe(pl_summary, key)
        if val != 0:
            report.add_pass(f"P&L mapping: {label}", f"= {val:,.0f}")
        else:
            report.add_warning(
                f"P&L mapping: {label} = 0",
                f"Key '{key}' not found or is genuinely zero. "
                "Check P&L sheet for the correct row label.")


def _check_missing_bs_items(bs_summary, report):
    """Check that key BS items were successfully extracted."""
    checks = [
        ('total_ca_cy', 'Total Current Assets'),
        ('total_cl_cy', 'Total Current Liabilities'),
        ('total_equity_cy', 'Total Equity'),
        ('debtors_cy', 'Trade Receivables'),
    ]
    for key, label in checks:
        val = _safe(bs_summary, key)
        if val != 0:
            report.add_pass(f"BS mapping: {label}", f"= {val:,.0f}")
        else:
            report.add_warning(
                f"BS mapping: {label} = 0",
                f"Key '{key}' not found or is genuinely zero. "
                "Check BS sheet for the correct row label.")


def _check_data_row_counts(bs_data, pl_data, report):
    """Sanity check that BS and PL have a reasonable number of rows."""
    bs_count = len(bs_data)
    pl_count = len(pl_data)

    if bs_count >= 5:
        report.add_pass("BS row count", f"{bs_count} rows parsed")
    else:
        report.add_fail(
            "BS row count is too low",
            f"Only {bs_count} rows parsed — BS sheet may not have been found correctly")

    if pl_count >= 5:
        report.add_pass("P&L row count", f"{pl_count} rows parsed")
    else:
        report.add_fail(
            "P&L row count is too low",
            f"Only {pl_count} rows parsed — P&L sheet may not have been found correctly")
