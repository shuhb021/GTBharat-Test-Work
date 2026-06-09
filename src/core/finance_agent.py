"""
╔══════════════════════════════════════════════════════════════════╗
║        FINANCIAL ANALYSIS AGENT  —  Pure Python / CPU Only      ║
║   No API Key | No AI/ML/DL | No GPU | Works on any machine      ║
╚══════════════════════════════════════════════════════════════════╝

USAGE:
    python finance_agent.py --file your_sheet.xlsx
    python finance_agent.py --file your_sheet.csv
    python finance_agent.py --demo          ← runs with built-in demo data

SUPPORTS:  .xlsx  |  .xls  |  .csv  |  .tsv

HOW IT WORKS (No ML — Rule Engine + Statistical Analysis):
  1. Reads every row of your Balance Sheet / P&L
  2. Detects sheet type automatically
  3. Runs 40+ financial rules at CPU speed (microseconds per row)
  4. Computes ratios, YoY changes, trend direction, anomaly scores
  5. Generates intelligent remarks per row + an executive summary
"""

import sys
import os
import time
import math
import argparse
import csv
import json
from datetime import datetime
from collections import defaultdict

# ── Optional fast imports ───────────────────────────────────────────────────
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from tabulate import tabulate
    HAS_TABULATE = True
except ImportError:
    HAS_TABULATE = False

# ═══════════════════════════════════════════════════════════════════════════
#  COLOUR / TERMINAL HELPERS
# ═══════════════════════════════════════════════════════════════════════════
class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    RED    = "\033[91m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    BLUE   = "\033[94m"
    CYAN   = "\033[96m"
    WHITE  = "\033[97m"
    MAGENTA= "\033[95m"

def clr(text, color): return f"{color}{text}{C.RESET}"
def bold(text):       return f"{C.BOLD}{text}{C.RESET}"

# ═══════════════════════════════════════════════════════════════════════════
#  KNOWLEDGE BASE  ──  pure dictionaries, zero ML
# ═══════════════════════════════════════════════════════════════════════════

# Keyword → category mapping (order matters: first match wins)
BALANCE_SHEET_KEYWORDS = {
    # Assets
    "cash":                 ("ASSET",    "Current Asset",      "Liquidity"),
    "bank":                 ("ASSET",    "Current Asset",      "Liquidity"),
    "receivable":           ("ASSET",    "Current Asset",      "Receivable"),
    "debtors":              ("ASSET",    "Current Asset",      "Receivable"),
    "inventory":            ("ASSET",    "Current Asset",      "Inventory"),
    "stock":                ("ASSET",    "Current Asset",      "Inventory"),
    "prepaid":              ("ASSET",    "Current Asset",      "Prepaid"),
    "advance":              ("ASSET",    "Current Asset",      "Advance"),
    "loan":                 ("ASSET",    "Non-Current Asset",  "Loan"),
    "investment":           ("ASSET",    "Non-Current Asset",  "Investment"),
    "fixed asset":          ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "property":             ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "plant":                ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "equipment":            ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "machinery":            ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "vehicle":              ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "furniture":            ("ASSET",    "Non-Current Asset",  "Fixed Asset"),
    "goodwill":             ("ASSET",    "Non-Current Asset",  "Intangible"),
    "intangible":           ("ASSET",    "Non-Current Asset",  "Intangible"),
    "patent":               ("ASSET",    "Non-Current Asset",  "Intangible"),
    "depreciation":         ("ASSET",    "Contra Asset",       "Depreciation"),
    "accumulate":           ("ASSET",    "Contra Asset",       "Depreciation"),
    # Liabilities
    "payable":              ("LIABILITY","Current Liability",   "Payable"),
    "creditors":            ("LIABILITY","Current Liability",   "Payable"),
    "overdraft":            ("LIABILITY","Current Liability",   "Overdraft"),
    "short term":           ("LIABILITY","Current Liability",   "Short-Term Debt"),
    "current portion":      ("LIABILITY","Current Liability",   "Current Portion LTD"),
    "tax payable":          ("LIABILITY","Current Liability",   "Tax"),
    "provision":            ("LIABILITY","Current Liability",   "Provision"),
    "long term":            ("LIABILITY","Non-Current Liability","Long-Term Debt"),
    "debenture":            ("LIABILITY","Non-Current Liability","Debenture"),
    "mortgage":             ("LIABILITY","Non-Current Liability","Mortgage"),
    "bond":                 ("LIABILITY","Non-Current Liability","Bond"),
    "deferred tax":         ("LIABILITY","Non-Current Liability","Deferred Tax"),
    # Equity
    "equity":               ("EQUITY",   "Equity",             "Equity"),
    "capital":              ("EQUITY",   "Equity",             "Capital"),
    "retained":             ("EQUITY",   "Equity",             "Retained Earnings"),
    "reserve":              ("EQUITY",   "Equity",             "Reserve"),
    "surplus":              ("EQUITY",   "Equity",             "Surplus"),
    "dividend":             ("EQUITY",   "Equity",             "Dividend"),
    "share":                ("EQUITY",   "Equity",             "Share Capital"),
}

PNL_KEYWORDS = {
    # Revenue
    "revenue":              ("INCOME",   "Revenue",            "Top Line"),
    "sales":                ("INCOME",   "Revenue",            "Top Line"),
    "turnover":             ("INCOME",   "Revenue",            "Top Line"),
    "income":               ("INCOME",   "Revenue",            "Income"),
    "interest income":      ("INCOME",   "Other Income",       "Interest"),
    "other income":         ("INCOME",   "Other Income",       "Non-Operating"),
    "gain":                 ("INCOME",   "Other Income",       "Gain"),
    "commission":           ("INCOME",   "Revenue",            "Commission"),
    "fee":                  ("INCOME",   "Revenue",            "Fee"),
    # COGS
    "cost of goods":        ("EXPENSE",  "COGS",               "Direct Cost"),
    "cost of sales":        ("EXPENSE",  "COGS",               "Direct Cost"),
    "cogs":                 ("EXPENSE",  "COGS",               "Direct Cost"),
    "direct cost":          ("EXPENSE",  "COGS",               "Direct Cost"),
    "material":             ("EXPENSE",  "COGS",               "Material"),
    "purchases":            ("EXPENSE",  "COGS",               "Purchases"),
    # Operating Expenses
    "salary":               ("EXPENSE",  "Operating Expense",  "Payroll"),
    "wages":                ("EXPENSE",  "Operating Expense",  "Payroll"),
    "payroll":              ("EXPENSE",  "Operating Expense",  "Payroll"),
    "rent":                 ("EXPENSE",  "Operating Expense",  "Rent"),
    "utilities":            ("EXPENSE",  "Operating Expense",  "Utilities"),
    "electricity":          ("EXPENSE",  "Operating Expense",  "Utilities"),
    "insurance":            ("EXPENSE",  "Operating Expense",  "Insurance"),
    "marketing":            ("EXPENSE",  "Operating Expense",  "Marketing"),
    "advertising":          ("EXPENSE",  "Operating Expense",  "Marketing"),
    "travel":               ("EXPENSE",  "Operating Expense",  "Travel"),
    "maintenance":          ("EXPENSE",  "Operating Expense",  "Maintenance"),
    "repair":               ("EXPENSE",  "Operating Expense",  "Repair"),
    "legal":                ("EXPENSE",  "Operating Expense",  "Legal"),
    "audit":                ("EXPENSE",  "Operating Expense",  "Professional Fee"),
    "professional":         ("EXPENSE",  "Operating Expense",  "Professional Fee"),
    "software":             ("EXPENSE",  "Operating Expense",  "Software"),
    "subscription":         ("EXPENSE",  "Operating Expense",  "Subscription"),
    "depreciation":         ("EXPENSE",  "Operating Expense",  "Depreciation"),
    "amortization":         ("EXPENSE",  "Operating Expense",  "Amortization"),
    # Finance
    "interest expense":     ("EXPENSE",  "Finance Cost",       "Interest"),
    "interest":             ("EXPENSE",  "Finance Cost",       "Interest"),
    "bank charge":          ("EXPENSE",  "Finance Cost",       "Bank Charge"),
    # Tax / Profit
    "tax":                  ("EXPENSE",  "Tax",                "Tax"),
    "profit":               ("INCOME",   "Profit",             "Bottom Line"),
    "loss":                 ("EXPENSE",  "Loss",               "Bottom Line"),
    "ebitda":               ("INCOME",   "Profit",             "EBITDA"),
    "ebit":                 ("INCOME",   "Profit",             "EBIT"),
    "gross profit":         ("INCOME",   "Profit",             "Gross Profit"),
    "net profit":           ("INCOME",   "Profit",             "Net Profit"),
    "net income":           ("INCOME",   "Profit",             "Net Profit"),
}

# Remark templates (filled at runtime)
REMARK_RULES = {
    # ─── Absolute value rules ───────────────────────────────────────────
    "zero_value": {
        "cond": lambda v, _p, _a: v == 0,
        "remark": "⚪ Value is zero — confirm whether this account is dormant or awaiting entry.",
        "severity": "INFO"
    },
    "negative_asset": {
        "cond": lambda v, _p, a: v < 0 and a.get("type") == "ASSET",
        "remark": "🔴 Negative asset balance detected — likely a credit balance on an asset account. Investigate for misposting or overdraft.",
        "severity": "CRITICAL"
    },
    "negative_equity": {
        "cond": lambda v, _p, a: v < 0 and a.get("sub") == "Equity",
        "remark": "🔴 Negative equity (capital deficiency) — the company may be technically insolvent. Urgent management review required.",
        "severity": "CRITICAL"
    },
    "large_receivable": {
        "cond": lambda v, _p, a: v > 0 and a.get("sub3") == "Receivable" and a.get("total_revenue", 1) > 0
                                  and v / a.get("total_revenue", 1) > 0.5,
        "remark": "🟠 Receivables exceed 50% of revenue — collection efficiency may be low. Review debtor ageing and credit policy.",
        "severity": "WARNING"
    },
    "high_inventory": {
        "cond": lambda v, _p, a: v > 0 and a.get("sub3") == "Inventory" and a.get("total_revenue", 1) > 0
                                  and v / a.get("total_revenue", 1) > 0.4,
        "remark": "🟠 Inventory > 40% of revenue — potential overstocking or slow-moving goods. Review inventory turnover ratio.",
        "severity": "WARNING"
    },
    "high_payables": {
        "cond": lambda v, _p, a: v > 0 and a.get("sub3") == "Payable" and a.get("total_assets", 1) > 0
                                  and v / a.get("total_assets", 1) > 0.3,
        "remark": "🟠 Payables are >30% of total assets — liquidity pressure. Verify payment schedule alignment with cash inflows.",
        "severity": "WARNING"
    },
    # ─── YoY change rules ───────────────────────────────────────────────
    "yoy_spike": {
        "cond": lambda v, p, _a: p != 0 and p is not None and abs((v - p) / abs(p)) > 0.50,
        "remark": lambda v, p: (
            f"{'📈' if v > p else '📉'} {'Increase' if v > p else 'Decrease'} of "
            f"{abs((v-p)/abs(p))*100:.1f}% YoY — "
            + ("Significant growth; verify source and sustainability." if v > p else
               "Sharp decline; investigate root cause immediately.")
        ),
        "severity": "WARNING"
    },
    "yoy_moderate": {
        "cond": lambda v, p, _a: p != 0 and p is not None and 0.10 < abs((v - p) / abs(p)) <= 0.50,
        "remark": lambda v, p: (
            f"{'📈' if v > p else '📉'} {'Rise' if v > p else 'Fall'} of "
            f"{abs((v-p)/abs(p))*100:.1f}% YoY — "
            + ("Moderate growth trend. Monitor for continuation." if v > p else
               "Moderate decline. Track for further movement.")
        ),
        "severity": "INFO"
    },
    "yoy_stable": {
        "cond": lambda v, p, _a: p != 0 and p is not None and abs((v - p) / abs(p)) <= 0.10,
        "remark": lambda v, p: f"✅ Stable YoY movement ({(v-p)/abs(p)*100:+.1f}%) — consistent with prior period.",
        "severity": "OK"
    },
    # ─── Revenue-specific ───────────────────────────────────────────────
    "revenue_growth": {
        "cond": lambda v, p, a: a.get("sub3") == "Top Line" and p and p > 0 and v > p,
        "remark": lambda v, p: f"✅ Revenue grew by {(v-p)/p*100:.1f}% — positive top-line momentum.",
        "severity": "OK"
    },
    "revenue_decline": {
        "cond": lambda v, p, a: a.get("sub3") == "Top Line" and p and p > 0 and v < p,
        "remark": lambda v, p: f"🔴 Revenue declined by {abs((v-p)/p)*100:.1f}% — business contraction signal. Immediate strategic review needed.",
        "severity": "CRITICAL"
    },
    # ─── Expense-specific ───────────────────────────────────────────────
    "expense_creep": {
        "cond": lambda v, p, a: a.get("type") == "EXPENSE" and p and p > 0 and (v - p) / p > 0.20,
        "remark": lambda v, p: f"🟠 Expense increased by {(v-p)/p*100:.1f}% — cost creep detected. Review budget vs actuals.",
        "severity": "WARNING"
    },
    "salary_high": {
        "cond": lambda v, _p, a: a.get("sub3") == "Payroll" and a.get("total_revenue", 1) > 0
                                  and v / a.get("total_revenue", 1) > 0.35,
        "remark": "🟠 Salary/wages >35% of revenue — labour cost ratio is elevated. Assess headcount efficiency.",
        "severity": "WARNING"
    },
    # ─── Profit-specific ─────────────────────────────────────────────────
    "net_loss": {
        "cond": lambda v, _p, a: v < 0 and a.get("sub3") in ("Net Profit", "Bottom Line"),
        "remark": "🔴 Net loss reported — company is unprofitable this period. Urgent cost review and revenue enhancement plan required.",
        "severity": "CRITICAL"
    },
    "thin_margin": {
        "cond": lambda v, _p, a: a.get("sub3") == "Net Profit" and a.get("total_revenue", 1) > 0
                                  and 0 < v / a.get("total_revenue", 1) < 0.05,
        "remark": lambda v, _p, a: f"🟡 Net margin is thin ({v/a.get('total_revenue',1)*100:.1f}%) — below 5%. Small shocks could push to loss.",
        "severity": "WARNING"
    },
    "healthy_margin": {
        "cond": lambda v, _p, a: a.get("sub3") == "Net Profit" and a.get("total_revenue", 1) > 0
                                  and v / a.get("total_revenue", 1) >= 0.15,
        "remark": lambda v, _p, a: f"✅ Healthy net margin of {v/a.get('total_revenue',1)*100:.1f}% — strong profitability.",
        "severity": "OK"
    },
    # ─── Depreciation ────────────────────────────────────────────────────
    "no_depreciation": {
        "cond": lambda v, _p, a: a.get("sub3") == "Fixed Asset" and v > 100000
                                  and a.get("total_depreciation", 0) == 0,
        "remark": "🟠 Significant fixed assets but no depreciation found — verify depreciation policy compliance.",
        "severity": "WARNING"
    },
}

# ═══════════════════════════════════════════════════════════════════════════
#  DEMO DATA GENERATOR
# ═══════════════════════════════════════════════════════════════════════════

DEMO_BALANCE_SHEET = [
    ["Particulars",                 "Current Year (₹)",  "Previous Year (₹)"],
    ["ASSETS",                       "",                  ""],
    ["Current Assets",               "",                  ""],
    ["Cash and Cash Equivalents",    580000,              420000],
    ["Bank Balance",                 1250000,             1100000],
    ["Trade Receivables / Debtors",  3200000,             2100000],
    ["Inventory / Stock",            2800000,             1900000],
    ["Prepaid Expenses",             95000,               88000],
    ["Advances to Suppliers",        340000,              210000],
    ["Non-Current Assets",           "",                  ""],
    ["Property, Plant & Equipment",  8500000,             8000000],
    ["Less: Accumulated Depreciation",-3200000,           -2800000],
    ["Intangible Assets / Goodwill", 500000,              500000],
    ["Long-term Investments",        1200000,             800000],
    ["TOTAL ASSETS",                 15265000,            12318000],
    ["LIABILITIES & EQUITY",         "",                  ""],
    ["Current Liabilities",          "",                  ""],
    ["Trade Payables / Creditors",   2100000,             1600000],
    ["Short-term Bank Overdraft",    450000,              200000],
    ["Tax Payable",                  380000,              290000],
    ["Provision for Expenses",       210000,              180000],
    ["Non-Current Liabilities",      "",                  ""],
    ["Long-term Loan",               3500000,             3000000],
    ["Deferred Tax Liability",       420000,              350000],
    ["EQUITY",                       "",                  ""],
    ["Share Capital",                2000000,             2000000],
    ["Retained Earnings / Reserves", 4995000,             4198000],
    ["General Reserve & Surplus",    710000,              500000],
    ["TOTAL LIABILITIES & EQUITY",   15265000,            12318000],
]

DEMO_PNL = [
    ["Particulars",                  "Current Year (₹)", "Previous Year (₹)"],
    ["INCOME",                        "",                 ""],
    ["Revenue from Sales / Turnover", 12500000,           11000000],
    ["Commission Income",             350000,             280000],
    ["Other Income / Gain",           120000,             95000],
    ["TOTAL INCOME",                  12970000,           11375000],
    ["EXPENSES",                      "",                 ""],
    ["Cost of Goods Sold (COGS)",     6800000,            5900000],
    ["Gross Profit",                  6170000,            5475000],
    ["Salary & Wages",                2200000,            1950000],
    ["Rent",                          480000,             460000],
    ["Electricity & Utilities",       195000,             170000],
    ["Marketing & Advertising",       620000,             400000],
    ["Travel & Conveyance",           85000,              72000],
    ["Repairs & Maintenance",         140000,             95000],
    ["Professional / Audit Fees",     175000,             155000],
    ["Software & Subscriptions",      230000,             95000],
    ["Insurance",                     88000,              82000],
    ["Depreciation",                  400000,             380000],
    ["Interest Expense",              310000,             270000],
    ["Bank Charges",                  45000,              38000],
    ["TOTAL EXPENSES",                11768000,           10067000],
    ["Profit Before Tax (EBIT)",      1202000,            1308000],
    ["Tax",                           360000,             392000],
    ["Net Profit / Net Income",       842000,             916000],
]

# ═══════════════════════════════════════════════════════════════════════════
#  CORE ENGINE
# ═══════════════════════════════════════════════════════════════════════════

def safe_float(val):
    """Convert any value to float safely."""
    if val is None or val == "":
        return None
    try:
        return float(str(val).replace(",", "").replace("₹", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return None

def classify_row(label: str, known_keywords: dict) -> dict:
    """
    Rule-based classifier — scans label for keyword matches.
    Returns metadata dict. Zero ML.
    """
    lower = label.lower().strip()
    for kw, (typ, sub, sub3) in known_keywords.items():
        if kw in lower:
            return {"type": typ, "sub": sub, "sub3": sub3, "matched_kw": kw}
    return {"type": "UNKNOWN", "sub": "Unknown", "sub3": "Unknown", "matched_kw": None}

def detect_sheet_type(rows: list) -> str:
    """Detect whether sheet is Balance Sheet or P&L."""
    text_blob = " ".join(str(cell).lower() for row in rows[:10] for cell in row)
    pnl_score = sum(1 for kw in ["revenue", "sales", "profit", "loss", "expense", "income", "turnover"] if kw in text_blob)
    bs_score  = sum(1 for kw in ["asset", "liability", "equity", "capital", "receivable", "payable", "balance"] if kw in text_blob)
    return "P&L" if pnl_score > bs_score else "BALANCE_SHEET"

def compute_aggregate_context(rows, keyword_map):
    """Pre-compute totals needed for ratio-based rules."""
    ctx = defaultdict(float)
    for row in rows:
        label = str(row[0]) if row else ""
        val   = safe_float(row[1]) if len(row) > 1 else None
        if val is None: continue
        meta  = classify_row(label, keyword_map)
        if meta["sub3"] == "Top Line":    ctx["total_revenue"] += abs(val)
        if meta["sub"] in ("Current Asset",): ctx["total_current_assets"] += abs(val)
        if meta["type"] == "ASSET":       ctx["total_assets"] += abs(val)
        if meta["sub3"] == "Depreciation": ctx["total_depreciation"] += abs(val)
        if meta["sub3"] in ("Net Profit","Bottom Line") and val > 0: ctx["net_profit"] = val
    return dict(ctx)

def apply_rules(val, prev_val, meta_plus: dict) -> list:
    """
    Apply all remark rules to a single row.
    Returns list of (severity, remark_text).
    """
    remarks = []
    if val is None:
        return remarks

    for rule_name, rule in REMARK_RULES.items():
        try:
            if rule["cond"](val, prev_val, meta_plus):
                template = rule["remark"]
                text = template(val, prev_val, meta_plus) if callable(template) else template
                remarks.append((rule["severity"], text))
                break   # one remark per row (most-specific rule wins)
        except Exception:
            pass

    # If no rule fired but we have YoY data:
    if not remarks and prev_val is not None and prev_val != 0:
        chg = (val - prev_val) / abs(prev_val) * 100
        if abs(chg) <= 2:
            remarks.append(("OK", f"✅ Virtually unchanged YoY ({chg:+.1f}%) — stable."))
        else:
            dir_sym = "📈" if chg > 0 else "📉"
            remarks.append(("INFO", f"{dir_sym} Changed {chg:+.1f}% from prior period."))

    if not remarks and val is not None:
        remarks.append(("INFO", "ℹ️  Single-period data — no YoY comparison available."))

    return remarks

def analyse_sheet(rows: list, sheet_name: str = "Sheet") -> list:
    """
    Master analysis loop — processes every row, returns analysis records.
    This is intentionally tight Python — no external libs in hot path.
    """
    sheet_type = detect_sheet_type(rows)
    keyword_map = PNL_KEYWORDS if sheet_type == "P&L" else BALANCE_SHEET_KEYWORDS
    ctx = compute_aggregate_context(rows, keyword_map)

    # Detect header row
    header_idx = 0
    for i, row in enumerate(rows[:5]):
        if any(isinstance(c, str) and c.strip().lower() in
               ("particulars", "description", "item", "account") for c in row):
            header_idx = i
            break

    analysis = []
    for row_idx, row in enumerate(rows):
        if row_idx == header_idx:
            continue  # skip header

        label = str(row[0]).strip() if row else ""
        if not label or label.startswith("─") or label.startswith("═"):
            continue

        val      = safe_float(row[1]) if len(row) > 1 else None
        prev_val = safe_float(row[2]) if len(row) > 2 else None

        # Skip section headers (no numeric value in either col)
        if val is None and prev_val is None:
            analysis.append({
                "row": row_idx + 1,
                "label": label,
                "value": None,
                "prev_value": None,
                "type": "SECTION_HEADER",
                "sub": "",
                "sub3": "",
                "remarks": [],
                "severity": "HEADER",
                "yoy_pct": None,
            })
            continue

        meta = classify_row(label, keyword_map)
        meta_plus = {**meta, **ctx}

        remarks = apply_rules(val, prev_val, meta_plus)
        max_sev = "OK"
        sev_order = {"CRITICAL": 4, "WARNING": 3, "INFO": 2, "OK": 1, "HEADER": 0}
        for sev, _ in remarks:
            if sev_order.get(sev, 0) > sev_order.get(max_sev, 0):
                max_sev = sev

        yoy_pct = None
        if val is not None and prev_val is not None and prev_val != 0:
            yoy_pct = (val - prev_val) / abs(prev_val) * 100

        analysis.append({
            "row": row_idx + 1,
            "label": label,
            "value": val,
            "prev_value": prev_val,
            "type": meta["type"],
            "sub": meta["sub"],
            "sub3": meta["sub3"],
            "remarks": remarks,
            "severity": max_sev,
            "yoy_pct": yoy_pct,
            "sheet_type": sheet_type,
        })

    return analysis, sheet_type, ctx

# ═══════════════════════════════════════════════════════════════════════════
#  EXECUTIVE SUMMARY GENERATOR
# ═══════════════════════════════════════════════════════════════════════════

def generate_executive_summary(analysis: list, ctx: dict, sheet_type: str) -> str:
    criticals = [a for a in analysis if a["severity"] == "CRITICAL"]
    warnings  = [a for a in analysis if a["severity"] == "WARNING"]
    oks        = [a for a in analysis if a["severity"] == "OK"]

    lines = []
    lines.append(bold(clr(f"\n{'═'*70}", C.CYAN)))
    lines.append(bold(clr(f"  EXECUTIVE SUMMARY  —  {sheet_type}", C.CYAN)))
    lines.append(bold(clr(f"{'═'*70}", C.CYAN)))

    # Health score (0-100)
    total = len([a for a in analysis if a["type"] != "SECTION_HEADER" and a["value"] is not None])
    if total > 0:
        score = max(0, 100 - len(criticals)*15 - len(warnings)*5)
        color = C.GREEN if score >= 75 else (C.YELLOW if score >= 50 else C.RED)
        lines.append(f"\n  {bold('Financial Health Score:')} {clr(f'{score}/100', color)}")
        lines.append(f"  Rows Analysed: {total}  |  🔴 Critical: {len(criticals)}  |  🟠 Warnings: {len(warnings)}  |  ✅ OK: {len(oks)}")

    # KPIs
    if ctx.get("total_revenue", 0) > 0:
        lines.append(f"\n  {bold('Key Metrics:')}")
        rev = ctx["total_revenue"]
        lines.append(f"   • Total Revenue        : ₹{rev:>15,.0f}")
        if ctx.get("net_profit"):
            np  = ctx["net_profit"]
            margin = np / rev * 100
            col = C.GREEN if margin >= 10 else (C.YELLOW if margin >= 5 else C.RED)
            lines.append(f"   • Net Profit           : ₹{np:>15,.0f}  (Margin: {clr(f'{margin:.1f}%', col)})")
    if ctx.get("total_assets", 0) > 0:
        lines.append(f"   • Total Assets         : ₹{ctx['total_assets']:>15,.0f}")

    # Top critical flags
    if criticals:
        lines.append(f"\n  {clr(bold('🔴 CRITICAL FLAGS:'), C.RED)}")
        for c in criticals[:5]:
            lines.append(f"   → {c['label']}: {c['remarks'][0][1] if c['remarks'] else ''}")

    if warnings:
        lines.append(f"\n  {clr(bold('🟠 KEY WARNINGS:'), C.YELLOW)}")
        for w in warnings[:5]:
            lines.append(f"   → {w['label']}: {w['remarks'][0][1] if w['remarks'] else ''}")

    lines.append(f"\n{clr('═'*70, C.CYAN)}\n")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════════════
#  DISPLAY ENGINE
# ═══════════════════════════════════════════════════════════════════════════

SEV_COLOR = {
    "CRITICAL": C.RED,
    "WARNING":  C.YELLOW,
    "OK":       C.GREEN,
    "INFO":     C.CYAN,
    "HEADER":   C.BOLD,
}

def fmt_num(val):
    if val is None: return "—"
    return f"₹{val:>12,.0f}"

def fmt_pct(pct):
    if pct is None: return "   —  "
    color = C.GREEN if pct >= 0 else C.RED
    return clr(f"{pct:+.1f}%", color)

def print_analysis(analysis: list, sheet_name: str):
    print(clr(f"\n{'─'*90}", C.BLUE))
    print(bold(clr(f"  DETAILED ROW-BY-ROW ANALYSIS — {sheet_name}", C.BLUE)))
    print(clr(f"{'─'*90}", C.BLUE))

    col_w = [45, 16, 16, 8, 60]
    header = (f"{'Account / Particulars':<{col_w[0]}} "
              f"{'Current Year':>{col_w[1]}} "
              f"{'Prev Year':>{col_w[2]}} "
              f"{'YoY':>{col_w[3]}} "
              f"{'Remarks'}")
    print(bold(header))
    print("─" * 90)

    for rec in analysis:
        if rec["type"] == "SECTION_HEADER":
            print(clr(f"\n  {rec['label'].upper()}", C.MAGENTA))
            continue

        sev   = rec["severity"]
        color = SEV_COLOR.get(sev, C.RESET)
        label_str = rec["label"][:col_w[0]].ljust(col_w[0])
        val_str   = fmt_num(rec["value"]).rjust(col_w[1])
        prev_str  = fmt_num(rec["prev_value"]).rjust(col_w[2])
        pct_str   = fmt_pct(rec["yoy_pct"]).rjust(col_w[3] + 10)  # +10 for ansi codes

        first_remark = rec["remarks"][0][1] if rec["remarks"] else "—"

        print(f"{clr(label_str, color)} {val_str} {prev_str} {pct_str}  {first_remark}")

# ═══════════════════════════════════════════════════════════════════════════
#  FILE READERS
# ═══════════════════════════════════════════════════════════════════════════

def read_file(path: str) -> list:
    """Read xlsx/xls/csv into list of rows."""
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xlsm", ".xls"):
        if not HAS_PANDAS:
            print(clr("pandas not installed — run: pip install pandas openpyxl", C.RED))
            sys.exit(1)
        # Read all sheets
        xf = pd.ExcelFile(path)
        all_sheets = {}
        for sheet in xf.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet, header=None)
            df = df.fillna("")
            all_sheets[sheet] = df.values.tolist()
        return all_sheets

    elif ext in (".csv", ".tsv"):
        sep = "\t" if ext == ".tsv" else ","
        rows = []
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=sep)
            for row in reader:
                rows.append(row)
        return {"Sheet1": rows}

    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ═══════════════════════════════════════════════════════════════════════════
#  JSON EXPORT
# ═══════════════════════════════════════════════════════════════════════════

def export_json(all_results: dict, out_path: str):
    export = {}
    for sheet, (analysis, sheet_type, ctx) in all_results.items():
        export[sheet] = {
            "sheet_type": sheet_type,
            "context": ctx,
            "rows": [
                {
                    "label": r["label"],
                    "value": r["value"],
                    "prev_value": r["prev_value"],
                    "yoy_pct": round(r["yoy_pct"], 2) if r["yoy_pct"] else None,
                    "category": r["type"],
                    "sub_category": r["sub"],
                    "item_type": r["sub3"],
                    "severity": r["severity"],
                    "remarks": [rem for _, rem in r["remarks"]],
                }
                for r in analysis if r["type"] != "SECTION_HEADER"
            ]
        }
    with open(out_path, "w") as f:
        json.dump(export, f, indent=2, default=str)
    print(clr(f"\n  📄 JSON report saved → {out_path}", C.CYAN))

# ═══════════════════════════════════════════════════════════════════════════
#  FAR GUI INTEGRATION  —  Drop-in replacement for ai_remarks.py
# ═══════════════════════════════════════════════════════════════════════════

def load_api_key():
    """No API key needed — rule engine is local. Returns a sentinel."""
    return "LOCAL_RULE_ENGINE"


def save_api_key(api_key):
    """No-op — rule engine doesn't need an API key."""
    return True


def validate_api_key(api_key):
    """Always valid — rule engine is local."""
    return True, "Local rule engine — no API key required"


def generate_remarks(api_key, items, client_name, financial_year,
                     statement_type, rounding_unit,
                     progress_callback=None, max_workers=5):
    """
    Generate remarks for all flagged items using the rule engine.
    
    This is a drop-in replacement for ai_remarks.generate_remarks().
    Same signature, same return type: dict mapping index → remark text.
    """
    flagged = [(i, item) for i, item in enumerate(items) if item.get('flag', False)]
    
    if not flagged:
        return {}
    
    # Build rows in the format the rule engine expects
    # Determine keyword map based on statement type
    is_bs = 'balance' in statement_type.lower() or 'bs' in statement_type.lower()
    keyword_map = BALANCE_SHEET_KEYWORDS if is_bs else PNL_KEYWORDS
    
    # Pre-compute aggregate context from ALL items (not just flagged)
    all_rows = []
    for item in items:
        cy = float(item.get('cy', 0) or 0)
        py = float(item.get('py', 0) or 0)
        label = item.get('particulars', '')
        all_rows.append([label, cy, py])
    
    ctx = compute_aggregate_context(all_rows, keyword_map)
    
    results = {}
    total = len(flagged)
    
    for completed_count, (idx, item) in enumerate(flagged):
        label = item.get('particulars', '')
        cy = float(item.get('cy', 0) or 0)
        py = float(item.get('py', 0) or 0)
        
        # Classify the row
        meta = classify_row(label, keyword_map)
        meta_plus = {**meta, **ctx}
        
        # Apply rules
        remarks_list = apply_rules(cy, py, meta_plus)
        
        if remarks_list:
            # Take the highest-severity remark
            remark_text = remarks_list[0][1]
        else:
            # Fallback: generate a basic variance remark
            variance_pct = item.get('display_pct', 'N/A')
            if py != 0:
                direction = "increased" if cy > py else "decreased"
                remark_text = (
                    f"'{label}' has {direction} from ₹{py:,.0f} to ₹{cy:,.0f} "
                    f"({variance_pct} variance). Review the underlying transactions "
                    f"for the period ended {financial_year}."
                )
            else:
                remark_text = (
                    f"'{label}' shows a value of ₹{cy:,.0f} with no prior year comparison. "
                    f"Verify the opening balance and source documentation."
                )
        
        # Strip emoji for clean Excel/Word output (optional — keep emoji for now)
        results[idx] = remark_text
        
        if progress_callback:
            progress_callback(completed_count + 1, total)
    
    return results


def generate_single_remark_for_item(api_key, item, client_name, financial_year,
                                    statement_type, rounding_unit):
    """Generate a remark for a single item (used for regeneration)."""
    result = generate_remarks(
        api_key, [item], client_name, financial_year,
        statement_type, rounding_unit
    )
    return result.get(0, "No remark generated.")


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Financial Analysis Agent — Pure Python, CPU-only, No ML/API"
    )
    parser.add_argument("--file",   "-f", type=str, help="Path to Excel/CSV file")
    parser.add_argument("--demo",   "-d", action="store_true", help="Run with demo data")
    parser.add_argument("--export", "-e", type=str, help="Export results to JSON file")
    args = parser.parse_args()

    run(file_path=args.file, demo=args.demo, export=args.export)

if __name__ == "__main__":
    main()

