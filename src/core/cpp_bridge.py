"""
FAR Automation Tool — C++ Engine Bridge
Provides a Python ctypes bridge to far_engine.dll for high-performance
variance and ratio calculations. Falls back to pure-Python if DLL not found.
"""

import ctypes
import json
import os
import math
import logging

logger = logging.getLogger(__name__)

# Try to load C++ DLL
_dll = None
_dll_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                         'cpp_engine', 'build', 'far_engine.dll')

try:
    if os.path.exists(_dll_path):
        _dll = ctypes.CDLL(_dll_path)
        _dll.compute_variances.restype = ctypes.c_char_p
        _dll.compute_variances.argtypes = [ctypes.c_char_p]
        _dll.compute_ratios.restype = ctypes.c_char_p
        _dll.compute_ratios.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
        try:
            _dll.compute_ratios_from_raw.restype = ctypes.c_char_p
            _dll.compute_ratios_from_raw.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
        except AttributeError:
            pass  # old dll version
        _dll.process_bulk_data.restype = ctypes.c_char_p
        _dll.process_bulk_data.argtypes = [ctypes.c_char_p, ctypes.c_int]
        _dll.free_result.restype = None
        _dll.free_result.argtypes = [ctypes.c_char_p]
        logger.info("C++ engine loaded from %s", _dll_path)
    else:
        logger.info("C++ DLL not found at %s — using Python fallback", _dll_path)
except Exception as e:
    logger.warning("Failed to load C++ DLL: %s — using Python fallback", e)
    _dll = None


def _using_cpp():
    """Check whether C++ engine is available."""
    return _dll is not None


# ── Pure-Python fallback implementations ──────────────────────────

def _py_compute_variances(data):
    """Python fallback for variance computation."""
    results = []
    for row in data:
        cy = float(row.get('cy', 0) or 0)
        py = float(row.get('py', 0) or 0)
        variance_abs = cy - py

        if py == 0:
            if cy == 0:
                variance_pct = 0.0
                display_pct = "0.0%"
            else:
                variance_pct = float('inf')
                display_pct = "N/A"
        else:
            variance_pct = (variance_abs / abs(py)) * 100.0
            display_pct = f"{variance_pct:.1f}%"

        abs_pct = 0.0 if (py == 0 and cy == 0) else (100.0 if py == 0 else abs(variance_pct))
        flag = abs_pct >= 10.0

        result = dict(row)
        result['variance_abs'] = round(variance_abs, 2)
        result['variance_pct'] = round(variance_pct, 2) if not math.isinf(variance_pct) else None
        result['display_pct'] = display_pct
        result['flag'] = flag
        results.append(result)
    return results


def _py_compute_ratios(bs_summary, pl_summary):
    """
    Python fallback for ratio computation.

    Uses:
    - Current Ratio  = Total Current Assets / Total Current Liabilities
    - Net Profit Ratio = PAT / Revenue from Ops × 100  (uses Profit After Tax)
    - Debtor Turnover = Revenue from Ops / Average Trade Receivables
    """
    def safe_get(d, key, default=0.0):
        v = d.get(key, default)
        return float(v) if v is not None else default

    # Extract BS values
    total_equity_cy = safe_get(bs_summary, 'total_equity_cy')
    total_equity_py = safe_get(bs_summary, 'total_equity_py')
    ltb_cy = safe_get(bs_summary, 'ltb_cy')
    ltb_py = safe_get(bs_summary, 'ltb_py')
    stb_cy = safe_get(bs_summary, 'stb_cy')
    stb_py = safe_get(bs_summary, 'stb_py')
    total_ca_cy = safe_get(bs_summary, 'total_ca_cy')
    total_ca_py = safe_get(bs_summary, 'total_ca_py')
    total_cl_cy = safe_get(bs_summary, 'total_cl_cy')
    total_cl_py = safe_get(bs_summary, 'total_cl_py')
    total_ncl_cy = safe_get(bs_summary, 'total_ncl_cy')
    total_ncl_py = safe_get(bs_summary, 'total_ncl_py')
    debtors_cy = safe_get(bs_summary, 'debtors_cy')
    debtors_py = safe_get(bs_summary, 'debtors_py')

    # Extract PL values
    revenue_ops_cy = safe_get(pl_summary, 'revenue_ops_cy')
    revenue_ops_py = safe_get(pl_summary, 'revenue_ops_py')
    total_revenue_cy = safe_get(pl_summary, 'total_revenue_cy')
    total_revenue_py = safe_get(pl_summary, 'total_revenue_py')
    pbt_cy = safe_get(pl_summary, 'pbt_cy')
    pbt_py = safe_get(pl_summary, 'pbt_py')
    pat_cy = safe_get(pl_summary, 'pat_cy')
    pat_py = safe_get(pl_summary, 'pat_py')
    finance_cost_cy = safe_get(pl_summary, 'finance_cost_cy')
    finance_cost_py = safe_get(pl_summary, 'finance_cost_py')

    # Average trade receivables: (CY + PY) / 2
    # CY = closing, PY = opening (prior year closing)
    avg_debtors_cy = (debtors_cy + debtors_py) / 2.0 if (debtors_cy + debtors_py) > 0 else 0.0
    # For PY ratio we only have one year of history, use PY closing as proxy
    avg_debtors_py = debtors_py

    # Log key inputs for traceability
    logger.info(
        "Ratio inputs — revenue_ops: CY=%.0f PY=%.0f | pat: CY=%.0f PY=%.0f | "
        "debtors: CY=%.0f PY=%.0f (avg_cy=%.0f) | "
        "total_ca: CY=%.0f PY=%.0f | total_cl: CY=%.0f PY=%.0f",
        revenue_ops_cy, revenue_ops_py,
        pat_cy, pat_py,
        debtors_cy, debtors_py, avg_debtors_cy,
        total_ca_cy, total_ca_py,
        total_cl_cy, total_cl_py
    )

    def calc_ratio(cy_num, cy_den, py_num, py_den, is_pct=False, ratio_name=''):
        """Compute a ratio with division-by-zero guard and logging."""
        if cy_den == 0:
            cy_val = 0.0
            logger.warning(
                "RATIO '%s': CY denominator is ZERO — ratio set to 0. "
                "Numerator was %.2f", ratio_name, cy_num)
        else:
            cy_val = cy_num / cy_den

        if py_den == 0:
            py_val = 0.0
            logger.warning(
                "RATIO '%s': PY denominator is ZERO — ratio set to 0. "
                "Numerator was %.2f", ratio_name, py_num)
        else:
            py_val = py_num / py_den

        if is_pct:
            cy_val *= 100
            py_val *= 100

        if py_val == 0:
            change = 0.0 if cy_val == 0 else 100.0
            display = "0.0%" if cy_val == 0 else "N/A"
        else:
            change = ((cy_val - py_val) / abs(py_val)) * 100.0
            display = f"{change:.1f}%"

        logger.info("RATIO '%s': CY=%.4f, PY=%.4f, Change=%.2f%%",
                    ratio_name, cy_val, py_val, change)

        return {
            'cy': round(cy_val, 4),
            'py': round(py_val, 4),
            'change': round(change, 2),
            'display_change': display,
            'flag': abs(change) >= 10.0
        }

    ratios = [
        {
            'key': 'Debt Equity Ratio',
            'formula': '(Long-Term Borrowings + Short-Term Borrowings) / Total Equity',
            **calc_ratio(ltb_cy + stb_cy, total_equity_cy,
                         ltb_py + stb_py, total_equity_py,
                         ratio_name='Debt Equity Ratio')
        },
        {
            # FIXED: Label was "Total Assets / Total Liabilities" — now corrected
            'key': 'Current Ratio',
            'formula': 'Current Assets / Current Liabilities',
            **calc_ratio(total_ca_cy, total_cl_cy,
                         total_ca_py, total_cl_py,
                         ratio_name='Current Ratio')
        },
        {
            'key': 'GCA Days',
            'formula': '(Total Current Assets / Total Revenue) × 365',
            **calc_ratio(total_ca_cy * 365, total_revenue_cy,
                         total_ca_py * 365, total_revenue_py,
                         ratio_name='GCA Days')
        },
        {
            'key': 'Total Outside Liab. vs Total Equity',
            'formula': '(Total Current Liabilities + Total Non-Current Liabilities) / Total Equity',
            **calc_ratio(total_cl_cy + total_ncl_cy, total_equity_cy,
                         total_cl_py + total_ncl_py, total_equity_py,
                         ratio_name='Total Outside Liab. vs Total Equity')
        },
        {
            # FIXED: Now uses PAT (Profit After Tax) instead of PBT
            'key': 'Net Profit Ratio',
            'formula': 'Net Profit After Tax / Revenue from Operations × 100',
            **calc_ratio(pat_cy, revenue_ops_cy,
                         pat_py, revenue_ops_py,
                         is_pct=True, ratio_name='Net Profit Ratio')
        },
        {
            'key': 'Return on Equity (ROE)',
            'formula': 'Profit Before Tax / Average Total Equity × 100',
            **calc_ratio(pbt_cy, (total_equity_cy + total_equity_py) / 2.0,
                         pbt_py, total_equity_py,
                         is_pct=True, ratio_name='Return on Equity (ROE)')
        },
        {
            'key': 'Return on Capital Employed (ROCE)',
            'formula': '(PBT + Finance Cost) / (Total Equity + LTB + STB) × 100',
            **calc_ratio(pbt_cy + finance_cost_cy,
                         total_equity_cy + ltb_cy + stb_cy,
                         pbt_py + finance_cost_py,
                         total_equity_py + ltb_py + stb_py,
                         is_pct=True, ratio_name='Return on Capital Employed (ROCE)')
        },
        {
            # FIXED: Uses Average Trade Receivables = (CY + PY) / 2 instead of closing only
            'key': 'Debtor Turnover Ratio',
            'formula': 'Revenue from Operations / Average Trade Receivables',
            **calc_ratio(revenue_ops_cy, avg_debtors_cy,
                         revenue_ops_py, avg_debtors_py,
                         ratio_name='Debtor Turnover Ratio')
        }
    ]
    return ratios


def _py_process_bulk_data(data):
    """Python fallback for bulk data processing (same as variances)."""
    return _py_compute_variances(data)


# ── Public API ────────────────────────────────────────────────────

def compute_variances(data):
    """
    Compute absolute and percentage variances for each row.

    Args:
        data: list of dicts with keys: particulars, note, cy, py, ...
    Returns:
        list of dicts with added keys: variance_abs, variance_pct, display_pct, flag
    """
    if _dll:
        try:
            json_in = json.dumps(data).encode('utf-8')
            result_ptr = _dll.compute_variances(json_in)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            logger.error("C++ compute_variances failed: %s — falling back", e)

    return _py_compute_variances(data)


def compute_ratios(bs_summary, pl_summary):
    """
    Compute all financial ratios.

    Args:
        bs_summary: dict with BS aggregate values (total_equity_cy, etc.)
        pl_summary: dict with PL aggregate values (revenue_ops_cy, etc.)
    Returns:
        list of ratio dicts with keys: key, formula, cy, py, change, display_change, flag
    """
    if _dll:
        try:
            bs_json = json.dumps(bs_summary).encode('utf-8')
            pl_json = json.dumps(pl_summary).encode('utf-8')
            result_ptr = _dll.compute_ratios(bs_json, pl_json)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            logger.error("C++ compute_ratios failed: %s — falling back", e)

    return _py_compute_ratios(bs_summary, pl_summary)


def compute_ratios_from_raw(bs_data, pl_data):
    """
    Compute financial ratios and extract summaries directly from raw parsed rows.

    Args:
        bs_data: list of raw Balance Sheet row dicts
        pl_data: list of raw P&L row dicts
    Returns:
        dict with keys: ratios, bs_summary, pl_summary
    """
    if _dll and hasattr(_dll, 'compute_ratios_from_raw'):
        try:
            bs_json = json.dumps(bs_data).encode('utf-8')
            pl_json = json.dumps(pl_data).encode('utf-8')
            result_ptr = _dll.compute_ratios_from_raw(bs_json, pl_json)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            logger.error("C++ compute_ratios_from_raw failed: %s — falling back", e)

    # Pure-Python fallback using existing parsing logic
    from src.core.excel_parser import extract_bs_summary, extract_pl_summary
    bs_summary = extract_bs_summary(bs_data)
    pl_summary = extract_pl_summary(pl_data)
    ratios = _py_compute_ratios(bs_summary, pl_summary)

    return {
        "ratios": ratios,
        "bs_summary": bs_summary,
        "pl_summary": pl_summary
    }


def process_bulk_data(data):
    """
    Process large datasets with optimized bulk computation.

    Args:
        data: list of dicts with cy, py values
    Returns:
        list of dicts with variance results
    """
    if _dll:
        try:
            json_in = json.dumps(data).encode('utf-8')
            result_ptr = _dll.process_bulk_data(json_in, len(data))
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            logger.error("C++ process_bulk_data failed: %s — falling back", e)

    return _py_process_bulk_data(data)
