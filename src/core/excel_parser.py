"""
FAR Automation Tool — Excel Parser
Parses client financial statements (Balance Sheet, P&L, Notes to Accounts)
from Excel files using openpyxl.
"""

import re
import logging
from datetime import datetime
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def _is_bold(cell):
    """Check if a cell's font is bold."""
    try:
        return cell.font and cell.font.bold
    except Exception:
        return False


def _get_value(cell):
    """Get numeric value from cell, return None if not a number."""
    v = cell.value
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        # Try to parse as number after removing commas
        cleaned = v.replace(',', '').replace(' ', '').strip()
        if cleaned.startswith('(') and cleaned.endswith(')'):
            cleaned = '-' + cleaned[1:-1]
        try:
            return float(cleaned)
        except (ValueError, TypeError):
            return None
    return None


def _get_text(cell):
    """Get text value from cell, return empty string if None."""
    v = cell.value
    if v is None:
        return ''
    return str(v).strip()


def _find_date_columns(ws, search_rows=range(5, 12), default_cy_year=2025, default_py_year=2024):
    """
    Find CY and PY date header columns by scanning for datetime(YYYY, 3, 31) values or year text.
    Returns (cy_col, py_col, cy_year, py_year, header_row).
    """
    date_cells = []
    for row_idx in search_rows:
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            v = cell.value
            if isinstance(v, datetime):
                date_cells.append((row_idx, col_idx, v.year))
            elif isinstance(v, str):
                # Check for text like "As at 31 March 2025", "31st March, 2025", "20XX"
                m = re.search(r'31\s*(?:st|th)?\s*Mar(?:ch)?\s*[,.]?\s*(\d{4}|20XX)', v, re.IGNORECASE)
                if m:
                    yr_str = m.group(1).upper()
                    if yr_str == '20XX':
                        date_cells.append((row_idx, col_idx, '20XX'))
                    else:
                        date_cells.append((row_idx, col_idx, int(yr_str)))
                else:
                    # Look for standalone "2024" or "2025"
                    m2 = re.search(r'\b(20\d{2})\b', v)
                    if m2:
                        date_cells.append((row_idx, col_idx, int(m2.group(1))))

    if len(date_cells) < 2:
        logger.warning("Could not find two date header columns, found: %s", date_cells)
        return None, None, None, None, None

    # Group by row, we need a row with at least 2 dates
    row_groups = {}
    for r, c, y in date_cells:
        row_groups.setdefault(r, []).append((c, y))
    
    valid_rows = {r: dates for r, dates in row_groups.items() if len(dates) >= 2}
    if not valid_rows:
        logger.warning("No single row had at least two date headers.")
        return None, None, None, None, None
        
    # Take the first valid row
    header_row = min(valid_rows.keys())
    dates_in_row = valid_rows[header_row]
    
    # Sort columns from left to right
    dates_in_row.sort(key=lambda x: x[0])
    
    # Assume left is CY and right is PY if it's 20XX
    if dates_in_row[0][1] == '20XX' or dates_in_row[1][1] == '20XX':
        cy_col = dates_in_row[0][0]
        cy_year = default_cy_year
        py_col = dates_in_row[1][0]
        py_year = default_py_year
    else:
        # Sort by year descending — highest year is CY
        dates_in_row.sort(key=lambda x: x[1], reverse=True)
        cy_col = dates_in_row[0][0]
        cy_year = dates_in_row[0][1]
        py_col = dates_in_row[1][0]
        py_year = dates_in_row[1][1]

    logger.info("Date headers found: CY=%s (col %d, row %d), PY=%s (col %d, row %d)",
                cy_year, cy_col, header_row, py_year, py_col, header_row)
    return cy_col, py_col, cy_year, py_year, header_row


def _extract_client_name(ws):
    """
    Scan the first 5 rows of column A for a non-empty text cell that looks like
    a company name. Returns the first non-empty string found.
    """
    for row_idx in range(1, 6):
        text = _get_text(ws.cell(row=row_idx, column=1))
        if text and len(text) > 3:
            logger.info("Client name extracted from row %d: %s", row_idx, text)
            return text
    # Also check first few merged-cell rows in column B if A is blank
    for row_idx in range(1, 6):
        for col_idx in range(1, 4):
            text = _get_text(ws.cell(row=row_idx, column=col_idx))
            if text and len(text) > 3:
                logger.info("Client name extracted from row %d col %d: %s",
                            row_idx, col_idx, text)
                return text
    logger.warning("Could not extract client name from workbook — field may be in a merged cell")
    return ''


def parse_balance_sheet(filepath):
    """
    Parse Balance Sheet from client Excel file.

    Expected structure (from sample):
    - Row 1: Client name
    - Row 3: "Balance Sheet as at 31 March YYYY"
    - Row 7: Headers with datetime(YYYY, 3, 31) in cols G (CY) and I (PY)
    - Col A: Particulars
    - Col C: Note numbers
    - Data from "Assets" row to "Total equity and liabilities" row

    Returns:
        dict with keys:
            'data': list of row dicts
            'cy_year': int
            'py_year': int
            'client_name': str
    """
    wb = load_workbook(filepath, data_only=True)

    # Try to find 'BS' sheet or first sheet
    ws = None
    for name in ['BS', 'Balance Sheet', 'BalanceSheet']:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]
        logger.info("No 'BS' sheet found, using first sheet: %s", ws.title)

    # Get client name — enhanced scan across first 5 rows
    client_name = _extract_client_name(ws)

    # Find date columns
    cy_col, py_col, cy_year, py_year, header_row = _find_date_columns(ws)
    if cy_col is None:
        raise ValueError("Could not find year headers (e.g. '31 March 2025') in the Balance Sheet.\n"
                         "Please load your actual client financial statement, not the blank template.")

    # Determine particulars and notes columns
    # In sample: Col A = Particulars, Col C = Notes
    part_col = 1   # Column A
    note_col = 3   # Column C

    # Parse data rows
    data = []
    start_row = header_row + 1
    in_data = False

    for row_idx in range(start_row, ws.max_row + 1):
        part_text = _get_text(ws.cell(row=row_idx, column=part_col))

        # Start parsing when we find "Assets"
        if not in_data and part_text.lower().strip() in ('assets', 'asset'):
            in_data = True

        if not in_data:
            continue

        # Get values
        note_text = _get_text(ws.cell(row=row_idx, column=note_col))
        cy_val = _get_value(ws.cell(row=row_idx, column=cy_col))
        py_val = _get_value(ws.cell(row=row_idx, column=py_col))
        is_bold = _is_bold(ws.cell(row=row_idx, column=part_col))

        # Skip completely empty rows
        if not part_text and cy_val is None and py_val is None:
            continue

        row_data = {
            'particulars': part_text,
            'note': note_text,
            'cy': cy_val if cy_val is not None else 0,
            'py': py_val if py_val is not None else 0,
            'is_bold': is_bold,
            'row_num': row_idx
        }
        data.append(row_data)

        # Stop at "Total equity and liabilities"
        if 'total equity and liabilities' in part_text.lower():
            break

    wb.close()
    logger.info("Parsed BS: %d rows, CY=%d, PY=%d", len(data), cy_year, py_year)

    return {
        'data': data,
        'cy_year': cy_year,
        'py_year': py_year,
        'client_name': client_name,
        'cy_col_letter': _col_letter(cy_col),
        'py_col_letter': _col_letter(py_col)
    }


def parse_profit_loss(filepath):
    """
    Parse Profit & Loss statement from client Excel file.

    Expected structure (from sample):
    - Col A: Particulars
    - Col D: Note numbers
    - Col E: CY values, Col G: PY values (row 7 has dates)
    - Data from "Revenue from operations" to end

    Returns:
        dict with keys: 'data', 'cy_year', 'py_year', 'client_name'
    """
    wb = load_workbook(filepath, data_only=True)

    # Try to find 'PL' or 'P&L' sheet
    ws = None
    for name in ['PL', 'P&L', 'Profit & Loss', 'Profit and Loss', 'ProfitLoss']:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        # Check if there's a second sheet
        if len(wb.sheetnames) > 1:
            ws = wb.worksheets[1]
        else:
            ws = wb.worksheets[0]
        logger.info("No 'PL' sheet found, using sheet: %s", ws.title)

    client_name = _extract_client_name(ws)

    # Find date columns
    cy_col, py_col, cy_year, py_year, header_row = _find_date_columns(ws)
    if cy_col is None:
        raise ValueError("Could not find year headers (e.g. '31 March 2025') in the P&L statement.\n"
                         "Please load your actual client financial statement, not the blank template.")

    # In sample PL: Col A = Particulars, Col D = Notes
    part_col = 1   # Column A
    note_col = 4   # Column D

    data = []
    start_row = header_row + 1

    for row_idx in range(start_row, ws.max_row + 1):
        part_text = _get_text(ws.cell(row=row_idx, column=part_col))
        note_text = _get_text(ws.cell(row=row_idx, column=note_col))
        cy_val = _get_value(ws.cell(row=row_idx, column=cy_col))
        py_val = _get_value(ws.cell(row=row_idx, column=py_col))
        is_bold = _is_bold(ws.cell(row=row_idx, column=part_col))

        if not part_text and cy_val is None and py_val is None:
            continue

        row_data = {
            'particulars': part_text,
            'note': note_text,
            'cy': cy_val if cy_val is not None else 0,
            'py': py_val if py_val is not None else 0,
            'is_bold': is_bold,
            'row_num': row_idx
        }
        data.append(row_data)

    wb.close()
    logger.info("Parsed PL: %d rows, CY=%d, PY=%d", len(data), cy_year, py_year)

    return {
        'data': data,
        'cy_year': cy_year,
        'py_year': py_year,
        'client_name': client_name,
        'cy_col_letter': _col_letter(cy_col),
        'py_col_letter': _col_letter(py_col)
    }


def parse_notes(filepath, cy_year=2025, py_year=2024):
    """
    Parse Notes to Accounts from client Excel file.
    Handles multiple sheets (3-4, 5-9, 10-17, 18-26) each containing
    multiple notes with their own headers.

    Returns:
        dict mapping sheet_name → list of note groups:
        {
            '3-4': [
                {
                    'note_num': '3',
                    'note_heading': 'Property and equipment',
                    'data': [...],
                    'is_asset_note': True/False
                },
                ...
            ],
            ...
        }
    """
    wb = load_workbook(filepath, data_only=True)

    # Known notes sheet names
    notes_sheet_names = ['3-4', '5-9', '10-17', '18-26']
    result = {}

    for sheet_name in notes_sheet_names:
        if sheet_name not in wb.sheetnames:
            logger.debug("Notes sheet '%s' not found in workbook", sheet_name)
            continue

        ws = wb[sheet_name]
        notes_groups = _parse_notes_sheet(ws, cy_year, py_year)
        if notes_groups:
            result[sheet_name] = notes_groups
        else:
            logger.warning("Notes sheet '%s' was found but yielded no parsed groups", sheet_name)

    wb.close()
    if not result:
        logger.warning("No notes were parsed. This usually happens if you load a blank template instead of actual client data.")
    else:
        logger.info("Parsed notes from %d sheets", len(result))
    return result


def _is_note_heading(ws, row_idx):
    """Check if a row starts a new note group and return (note_num, note_heading) or (None, None)."""
    a_cell = ws.cell(row=row_idx, column=1)
    b_cell = ws.cell(row=row_idx, column=2)

    note_num = None
    note_heading = ''

    # 1. Check if column A has a note number (can be bold or not, but must match pattern)
    if a_cell.value is not None:
        a_val = str(a_cell.value).strip()
        if re.match(r'^\d+[A-Za-z]?$', a_val):
            note_num = a_val
            note_heading = _get_text(b_cell)
            return note_num, note_heading

    # 2. Check if column B contains "Note X:" pattern
    if b_cell.value is not None:
        b_val = str(b_cell.value).strip()
        m = re.match(r'^Note\s*(\d+[A-Za-z]?)\s*[:\-]?\s*(.*)$', b_val, re.IGNORECASE)
        if m:
            note_num = m.group(1)
            note_heading = m.group(2).strip()
            return note_num, note_heading

    # 3. Fallback for known headings in column B (must be bold to avoid false positives)
    if b_cell.value is not None and _is_bold(b_cell):
        b_val = str(b_cell.value).strip()
        lower_b = b_val.lower()
        if any(kw in lower_b for kw in ['property', 'plant', 'equipment', 'intangible', 'right of use', 'right-of-use', 'rou', 'ppe']):
            note_heading = b_val
            if 'intangible' in lower_b:
                note_num = '3'
            elif 'right' in lower_b or 'rou' in lower_b:
                note_num = '3A'
            else:
                note_num = '3'
            return note_num, note_heading

    return None, None


def _parse_notes_sheet(ws, cy_year=2025, py_year=2024):
    """Parse a single notes sheet, extracting all note groups."""
    notes = []

    row = 1
    while row <= ws.max_row:
        note_num, note_heading = _is_note_heading(ws, row)

        if note_num is None:
            row += 1
            continue

        # Found a note heading. Now find the date headers below it.
        cy_col, py_col, cy_yr, py_yr, date_row = _find_date_columns(
            ws, search_rows=range(row, min(row + 10, ws.max_row + 1)),
            default_cy_year=cy_year, default_py_year=py_year)

        # Detect if this is an asset note (PPE)
        heading_clean = note_heading.lower().replace('-', ' ')
        is_asset_note = any(kw in heading_clean for kw in
                           ['property', 'plant', 'equipment', 'intangible',
                            'right of use', 'rou', 'ppe'])

        if cy_col is None:
            if is_asset_note:
                # For PPE schedules, find the "Total" column
                found_total = False
                for r in range(row, min(row + 10, ws.max_row + 1)):
                    for c in range(1, ws.max_column + 1):
                        val = _get_text(ws.cell(row=r, column=c)).lower()
                        if val == 'total':
                            cy_col = c
                            py_col = c
                            cy_yr = cy_year
                            py_yr = py_year
                            date_row = r
                            logger.info("Note %s: Found 'Total' column at col %d for Asset note", note_num, c)
                            found_total = True
                            break
                    if found_total:
                        break
            
            if cy_col is None:
                # Fallback to default columns C (3) and D (4)
                cy_col = 3
                py_col = 4
                cy_yr = cy_year
                py_yr = py_year
                date_row = row + 1
                logger.info("Note %s: Could not find date columns — falling back to columns 3 and 4", note_num)
        else:
            cy_year = cy_yr
            py_year = py_yr

        # Parse data rows until next note heading or end
        data_rows = []
        part_col = 2  # Notes typically use col B for particulars
        data_start = date_row + 1 if date_row else row + 3

        for data_row in range(data_start, ws.max_row + 1):
            # Check if we hit the next note heading
            next_num, next_heading = _is_note_heading(ws, data_row)
            if next_num is not None:
                break  # Next note starts here

            part_text = _get_text(ws.cell(row=data_row, column=part_col))
            cy_val = _get_value(ws.cell(row=data_row, column=cy_col))
            py_val = _get_value(ws.cell(row=data_row, column=py_col))
            is_bold_row = _is_bold(ws.cell(row=data_row, column=part_col))

            if not part_text and cy_val is None and py_val is None:
                continue

            data_rows.append({
                'particulars': part_text,
                'cy': cy_val if cy_val is not None else 0,
                'py': py_val if py_val is not None else 0,
                'is_bold': is_bold_row,
                'row_num': data_row
            })
            row = data_row

        notes.append({
            'note_num': note_num,
            'note_heading': note_heading,
            'data': data_rows,
            'is_asset_note': is_asset_note,
            'cy_year': cy_yr,
            'py_year': py_yr
        })

        row += 1

    return notes


def _col_letter(col_idx):
    """Convert 1-based column index to Excel letter (1='A', 2='B', etc.)."""
    result = ''
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _find_first_nonzero(data, keywords, match_type='includes'):
    """
    Search rows for the first row matching any keyword that has a non-zero CY value.
    Falls back to first matching row even if value is zero.
    Logs a warning for each zero-value match.

    Returns (cy_val, py_val).
    """
    first_match_cy = None
    first_match_py = None

    for kw in keywords:
        kw_lower = kw.lower().strip()
        for row in data:
            p = row.get('particulars', '').lower().strip()
            matched = False
            if match_type == 'exact':
                matched = (p == kw_lower)
            else:
                matched = (kw_lower in p)

            if matched:
                cy = float(row.get('cy', 0) or 0)
                py = float(row.get('py', 0) or 0)

                if cy != 0 or py != 0:
                    logger.info("Found non-zero value for '%s' on row '%s': CY=%.2f, PY=%.2f",
                                kw, row.get('particulars', ''), cy, py)
                    return cy, py

                # Record first zero match as fallback
                if first_match_cy is None:
                    first_match_cy = cy
                    first_match_py = py
                    logger.warning(
                        "Keyword '%s' matched row '%s' but value is ZERO — "
                        "may be a heading row; scanning further",
                        kw, row.get('particulars', ''))

    if first_match_cy is not None:
        logger.warning(
            "All matches for keywords %s returned zero — using zero as fallback", keywords)
        return first_match_cy, first_match_py

    logger.warning("No row found matching any of keywords: %s", keywords)
    return 0.0, 0.0


def extract_bs_summary(bs_data):
    """
    Extract summary values from parsed BS data for ratio calculations.
    Returns dict with keys like total_equity_cy, total_ca_cy, etc.
    Uses robust multi-keyword, non-zero scanning to avoid returning 0 incorrectly.
    """
    summary = {}

    def find_value(keyword, match_type='includes'):
        """Find CY and PY values by keyword match in particulars."""
        kw = keyword.lower().strip()
        for row in bs_data:
            p = row.get('particulars', '').lower().strip()
            if match_type == 'exact':
                if p == kw:
                    return row.get('cy', 0), row.get('py', 0)
            else:
                if kw in p:
                    return row.get('cy', 0), row.get('py', 0)
        return 0, 0

    def find_liability_by_note(note_num, is_current):
        """Find liability values by note number and current/non-current section."""
        cy_total = 0
        py_total = 0
        in_current_section = False

        for row in bs_data:
            p = row.get('particulars', '').lower().strip()
            if 'current liabilities' in p and 'non' not in p:
                in_current_section = True
            elif 'non current liabilities' in p or 'non-current liabilities' in p:
                in_current_section = False

            note = str(row.get('note', '')).strip()
            if note == note_num and in_current_section == is_current:
                cy_total += float(row.get('cy', 0) or 0)
                py_total += float(row.get('py', 0) or 0)

        return cy_total, py_total

    # Total Equity
    cy, py = find_value('total equity', 'exact')
    if cy == 0 and py == 0:
        cy, py = find_value('total equity')
    summary['total_equity_cy'] = cy
    summary['total_equity_py'] = py

    # Total Current Assets
    cy, py = find_value('total current assets')
    summary['total_ca_cy'] = cy
    summary['total_ca_py'] = py

    # Total Current Liabilities
    cy, py = find_value('total current liabilities')
    summary['total_cl_cy'] = cy
    summary['total_cl_py'] = py

    # Total Non-Current Liabilities
    cy, py = find_value('total non-current liabilities')
    if cy == 0 and py == 0:
        cy, py = find_value('total non current liabilities')
    summary['total_ncl_cy'] = cy
    summary['total_ncl_py'] = py

    # Total Assets
    cy, py = find_value('total assets')
    summary['total_assets_cy'] = cy
    summary['total_assets_py'] = py

    # LTB (Long-Term Borrowings): Non-current Preference shares + Lease liabilities
    ltb_pref_cy, ltb_pref_py = find_liability_by_note('11', False)
    ltb_lease_cy, ltb_lease_py = find_liability_by_note('12', False)
    summary['ltb_cy'] = ltb_pref_cy + ltb_lease_cy
    summary['ltb_py'] = ltb_pref_py + ltb_lease_py

    # STB (Short-Term Borrowings): Current Preference shares + Lease liabilities
    stb_pref_cy, stb_pref_py = find_liability_by_note('11', True)
    stb_lease_cy, stb_lease_py = find_liability_by_note('12', True)
    summary['stb_cy'] = stb_pref_cy + stb_lease_cy
    summary['stb_py'] = stb_pref_py + stb_lease_py

    # Trade Receivables (Debtors) — scan multiple keywords, pick first non-zero row
    debtors_keywords = [
        'trade receivables',
        'sundry debtors',
        'trade and other receivables',
        'receivables',
        'debtors',
    ]
    cy, py = _find_first_nonzero(bs_data, debtors_keywords)
    summary['debtors_cy'] = cy
    summary['debtors_py'] = py

    if cy == 0:
        logger.warning(
            "DEBTORS: Could not find non-zero trade receivables in BS — "
            "Debtor Turnover Ratio will be 0. "
            "Check that the Notes sheet contains the actual receivables figures.")

    logger.info("BS Summary: total_ca=%.0f, total_cl=%.0f, debtors=%.0f, equity=%.0f",
                summary['total_ca_cy'], summary['total_cl_cy'],
                summary['debtors_cy'], summary['total_equity_cy'])

    return summary


def extract_pl_summary(pl_data):
    """
    Extract summary values from parsed PL data for ratio calculations.
    Uses robust multi-keyword, non-zero scanning to avoid returning 0 incorrectly.
    Extracts both PBT and PAT (Profit After Tax).
    """
    summary = {}

    # ── Revenue from Operations ──
    # Try specific sub-items first, then heading rows with non-zero values
    revenue_keywords = [
        'revenue from operations',
        'net revenue from operations',
        'income from operations',
        'software services',
        'sale of products',
        'sale of services',
        'revenue',
    ]
    cy, py = _find_first_nonzero(pl_data, revenue_keywords)
    summary['revenue_ops_cy'] = cy
    summary['revenue_ops_py'] = py

    if cy == 0:
        logger.warning(
            "REVENUE: Could not find non-zero revenue from operations in P&L — "
            "Net Profit Ratio and Debtor Turnover Ratio will be 0. "
            "Check P&L sheet structure.")

    # ── Total Income ──
    total_income_cy, total_income_py = 0.0, 0.0
    for kw in ['total income', 'total revenue']:
        cy_t, py_t = _find_first_nonzero(pl_data, [kw])
        if cy_t != 0:
            total_income_cy, total_income_py = cy_t, py_t
            break
    summary['total_revenue_cy'] = total_income_cy
    summary['total_revenue_py'] = total_income_py

    # ── Profit Before Tax (PBT) ──
    # Must match "profit before tax" specifically — not just any "profit" row
    pbt_cy, pbt_py = 0.0, 0.0
    pbt_keywords = [
        'profit before tax',
        'profit/(loss) before tax',
        'profit before income tax',
        'earnings before tax',
        'pbt',
    ]
    # Exact-priority search: look for rows where ALL of (profit, before, tax) appear
    for row in pl_data:
        p = row.get('particulars', '').lower().strip()
        if 'profit' in p and 'before' in p and 'tax' in p:
            cy_v = float(row.get('cy', 0) or 0)
            py_v = float(row.get('py', 0) or 0)
            pbt_cy = cy_v
            pbt_py = py_v
            logger.info("PBT found: row='%s', CY=%.2f, PY=%.2f",
                        row.get('particulars', ''), cy_v, py_v)
            break
    else:
        pbt_cy, pbt_py = _find_first_nonzero(pl_data, pbt_keywords)
        logger.warning("PBT: exact match failed, using keyword fallback: CY=%.2f", pbt_cy)

    summary['pbt_cy'] = pbt_cy
    summary['pbt_py'] = pbt_py

    # ── Profit After Tax / Net Profit (PAT) ──
    # This is the correct numerator for Net Profit Ratio
    pat_cy, pat_py = 0.0, 0.0
    # Priority ordered: most specific first
    pat_row_candidates = [
        ('profit after tax', False),
        ('profit/(loss) after tax', False),
        ('net profit after tax', False),
        ('profit for the year', False),
        ('profit/(loss) for the year', False),
        ('net profit for the year', False),
        ('net profit', False),
        ('profit for the period', False),
        ('total comprehensive income', False),
    ]
    for kw, exact in pat_row_candidates:
        kw_lower = kw.lower()
        for row in pl_data:
            p = row.get('particulars', '').lower().strip()
            if (exact and p == kw_lower) or (not exact and kw_lower in p):
                cy_v = float(row.get('cy', 0) or 0)
                py_v = float(row.get('py', 0) or 0)
                if cy_v != 0 or py_v != 0:
                    pat_cy = cy_v
                    pat_py = py_v
                    logger.info("PAT found via keyword '%s': row='%s', CY=%.2f, PY=%.2f",
                                kw, row.get('particulars', ''), cy_v, py_v)
                    break
        if pat_cy != 0 or pat_py != 0:
            break

    if pat_cy == 0 and pat_py == 0:
        # Last resort: use PBT if PAT not found
        pat_cy = pbt_cy
        pat_py = pbt_py
        logger.warning(
            "PAT: Could not find Profit After Tax row — falling back to PBT (%.2f). "
            "Net Profit Ratio may be slightly overstated (ignores tax).", pat_cy)

    summary['pat_cy'] = pat_cy
    summary['pat_py'] = pat_py

    # ── Finance Charges / Finance Cost ──
    finance_keywords = ['finance charge', 'finance cost', 'interest expense',
                        'interest and finance cost', 'borrowing cost']
    cy, py = _find_first_nonzero(pl_data, finance_keywords)
    summary['finance_cost_cy'] = cy
    summary['finance_cost_py'] = py

    logger.info("PL Summary: revenue_ops=%.0f, pbt=%.0f, pat=%.0f, finance_cost=%.0f",
                summary['revenue_ops_cy'], summary['pbt_cy'],
                summary['pat_cy'], summary['finance_cost_cy'])

    return summary
