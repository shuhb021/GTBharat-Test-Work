"""
FAR Automation Tool — C++ Engine Bridge
Provides a Python ctypes bridge to far_engine.dll for high-performance
variance and ratio calculations. Falls back to pure-Python if DLL not found.

PHASE 5 OPTIMIZATION: Zero-copy C++ struct bindings implemented to avoid JSON serialization overhead.
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

# PHASE 5: Zero-Copy C++ Structure Binding
class FinRow(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ("cy", ctypes.c_double),
        ("py", ctypes.c_double),
        ("variance_abs", ctypes.c_double),
        ("variance_pct", ctypes.c_double),
        ("flag", ctypes.c_bool)
    ]

try:
    if os.path.exists(_dll_path):
        _dll = ctypes.CDLL(_dll_path)
        _dll.compute_variances.restype = ctypes.c_char_p
        _dll.compute_variances.argtypes = [ctypes.c_char_p]
        _dll.compute_ratios.restype = ctypes.c_char_p
        _dll.compute_ratios.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
        
        try:
            # Phase 5: Fast zero-copy bindings via memory array
            _dll.compute_variances_fast.restype = None
            _dll.compute_variances_fast.argtypes = [ctypes.POINTER(FinRow), ctypes.c_int]
            logger.info("C++ Fast Zero-Copy Binding initialized.")
        except AttributeError:
            pass  # old dll version
            
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
    """
    def safe_get(d, key, default=0.0):
        v = d.get(key, default)
        return float(v) if v is not None else default

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

    avg_debtors_cy = (debtors_cy + debtors_py) / 2.0 if (debtors_cy + debtors_py) > 0 else 0.0
    avg_debtors_py = debtors_py

    def calc_ratio(cy_num, cy_den, py_num, py_den, is_pct=False, ratio_name=''):
        if cy_den == 0: cy_val = 0.0
        else: cy_val = cy_num / cy_den
        if py_den == 0: py_val = 0.0
        else: py_val = py_num / py_den

        if is_pct:
            cy_val *= 100
            py_val *= 100

        if py_val == 0:
            change = 0.0 if cy_val == 0 else 100.0
            display = "0.0%" if cy_val == 0 else "N/A"
        else:
            change = ((cy_val - py_val) / abs(py_val)) * 100.0
            display = f"{change:.1f}%"

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
            'key': 'Debtor Turnover Ratio',
            'formula': 'Revenue from Operations / Average Trade Receivables',
            **calc_ratio(revenue_ops_cy, avg_debtors_cy,
                         revenue_ops_py, avg_debtors_py,
                         ratio_name='Debtor Turnover Ratio')
        }
    ]
    return ratios


def _py_process_bulk_data(data):
    """Python fallback for bulk data processing."""
    return _py_compute_variances(data)


# ── Public API ────────────────────────────────────────────────────

def compute_variances(data):
    """
    Compute absolute and percentage variances for each row.
    Uses Phase 5 Zero-Copy Bindings if available.
    """
    if _dll and hasattr(_dll, 'compute_variances_fast'):
        try:
            # PHASE 5: Batch data movement to flat memory array (Zero-Copy)
            count = len(data)
            row_array = (FinRow * count)()
            
            for i, row in enumerate(data):
                row_array[i].cy = float(row.get('cy', 0) or 0)
                row_array[i].py = float(row.get('py', 0) or 0)
                
            # Execute natively in C++ via flat memory pointer
            _dll.compute_variances_fast(row_array, count)
            
            # Map results back seamlessly
            results = []
            for i, row in enumerate(data):
                res = dict(row)
                res['variance_abs'] = row_array[i].variance_abs
                v_pct = row_array[i].variance_pct
                
                if v_pct == -999999.0:
                    res['variance_pct'] = None
                    res['display_pct'] = "N/A"
                else:
                    res['variance_pct'] = v_pct
                    res['display_pct'] = f"{v_pct:.1f}%"
                    
                res['flag'] = row_array[i].flag
                results.append(res)
            return results
        except Exception as e:
            logger.error("C++ compute_variances_fast zero-copy failed: %s — falling back", e)
            
    # Legacy JSON method fallback
    if _dll:
        try:
            json_in = json.dumps(data).encode('utf-8')
            result_ptr = _dll.compute_variances(json_in)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            pass

    return _py_compute_variances(data)


def compute_ratios(bs_summary, pl_summary):
    if _dll:
        try:
            bs_json = json.dumps(bs_summary).encode('utf-8')
            pl_json = json.dumps(pl_summary).encode('utf-8')
            result_ptr = _dll.compute_ratios(bs_json, pl_json)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            pass

    return _py_compute_ratios(bs_summary, pl_summary)


def compute_ratios_from_raw(bs_data, pl_data):
    if _dll and hasattr(_dll, 'compute_ratios_from_raw'):
        try:
            bs_json = json.dumps(bs_data).encode('utf-8')
            pl_json = json.dumps(pl_data).encode('utf-8')
            result_ptr = _dll.compute_ratios_from_raw(bs_json, pl_json)
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            pass

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
    if _dll:
        try:
            json_in = json.dumps(data).encode('utf-8')
            result_ptr = _dll.process_bulk_data(json_in, len(data))
            result = json.loads(result_ptr.decode('utf-8'))
            _dll.free_result(result_ptr)
            return result
        except Exception as e:
            pass

    return _py_process_bulk_data(data)
