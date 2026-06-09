"""
FAR Automation Tool — Chart Generator for Reports
Generates light-themed, professional versions of the dashboard visuals
for embedding in Excel and Word documents.
"""

import os
import numpy as np
import matplotlib
# Use a non-interactive backend to avoid GUI threads issues
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ── Color Palette for Light Theme ──
BLUE = '#007ACC'
GREEN = '#22C55E'
YELLOW = '#EAB308'
RED = '#EF4444'
TEXT_COLOR = '#333333'
BORDER_COLOR = '#E5E7EB'
GRID_COLOR = '#F3F4F6'


def _find_val(data, keyword):
    """Find CY and PY values from data by keyword."""
    kw = keyword.lower()
    for row in data:
        if kw in row.get('particulars', '').lower():
            return float(row.get('cy', 0) or 0), float(row.get('py', 0) or 0)
    return 0.0, 0.0


def _style_axes(ax):
    """Apply professional, light-themed styling to axes."""
    ax.set_facecolor('#FFFFFF')
    ax.tick_params(colors=TEXT_COLOR, labelsize=9)
    for spine in ax.spines.values():
        spine.set_color(BORDER_COLOR)
        spine.set_linewidth(1)
    ax.xaxis.label.set_color(TEXT_COLOR)
    ax.yaxis.label.set_color(TEXT_COLOR)
    ax.grid(True, linestyle='--', alpha=0.5, color=GRID_COLOR)


def generate_all_charts(bs_data, pl_data, ratios, cy_year, py_year, rounding_unit, temp_dir):
    """
    Generates the 6 visuals using light themes and saves them as PNGs.
    Returns a list of absolute file paths to the generated charts.
    """
    os.makedirs(temp_dir, exist_ok=True)
    paths = []

    # Format helpers
    fmt_suffix = ''
    divisor = 1.0
    if rounding_unit == 'Thousands':
        fmt_suffix = 'K'
        divisor = 1e3
    elif rounding_unit == 'Lakhs':
        fmt_suffix = 'L'
        divisor = 1e5
    elif rounding_unit == 'Millions':
        fmt_suffix = 'M'
        divisor = 1e6

    # 1. Revenue vs Expenses
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        rev_cy, rev_py = _find_val(pl_data, 'total income')
        exp_cy, exp_py = _find_val(pl_data, 'total expense')
        pbt_cy, pbt_py = 0.0, 0.0
        for row in pl_data:
            p = row.get('particulars', '').lower()
            if 'profit' in p and 'before' in p and 'tax' in p:
                pbt_cy = float(row.get('cy', 0) or 0)
                pbt_py = float(row.get('py', 0) or 0)
                break

        x = np.arange(3)
        width = 0.35
        
        cy_vals = [rev_cy, exp_cy, pbt_cy]
        py_vals = [rev_py, exp_py, pbt_py]
        
        ax.bar(x - width/2, cy_vals, width, label=f'CY ({cy_year})', color=BLUE, edgecolor='none')
        ax.bar(x + width/2, py_vals, width, label=f'PY ({py_year})', color=GREEN, alpha=0.8, edgecolor='none')
        
        ax.set_title('Revenue vs Expenses (CY/PY)', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        ax.set_xticks(x)
        ax.set_xticklabels(['Revenue', 'Expenses', 'PBT'], color=TEXT_COLOR, fontsize=9)
        ax.legend(facecolor='#FFFFFF', edgecolor=BORDER_COLOR, labelcolor=TEXT_COLOR, fontsize=8, loc='upper right')
        
        # Formatter for y-axis
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, _: f'{val/divisor:.1f}{fmt_suffix}' if val != 0 else '0'))
        
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_revenue_expenses.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 1: {e}")

    # 2. Profit Before Tax Trend
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        pbt_cy, pbt_py = 0.0, 0.0
        for row in pl_data:
            p = row.get('particulars', '').lower()
            if 'profit' in p and 'before' in p and 'tax' in p:
                pbt_cy = float(row.get('cy', 0) or 0)
                pbt_py = float(row.get('py', 0) or 0)
                break
        
        years = [str(py_year), str(cy_year)]
        values = [pbt_py, pbt_cy]
        
        ax.plot(years, values, color=BLUE, linewidth=2.5, marker='o',
                markersize=8, markerfacecolor=BLUE, markeredgecolor='#FFFFFF',
                markeredgewidth=2)
        ax.fill_between(years, values, alpha=0.15, color=BLUE)
        
        for i, (year, val) in enumerate(zip(years, values)):
            ax.annotate(f'{val:,.2f}' if rounding_unit in ('Lakhs', 'Millions', 'Actuals') else f'{val:,.0f}',
                        (year, val), textcoords="offset points",
                        xytext=(0, 10), ha='center', fontsize=9, color=TEXT_COLOR, fontweight='bold')
        
        ax.set_title('Profit Before Tax Trend', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        ax.set_ylabel(f'Amount (Rs. in {rounding_unit})', color=TEXT_COLOR, fontsize=9)
        
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_pbt_trend.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 2: {e}")

    # 3. Asset Composition (CY)
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        nca_cy, _ = _find_val(bs_data, 'total non-current assets')
        if nca_cy == 0:
            nca_cy, _ = _find_val(bs_data, 'total non current assets')
        ca_cy, _ = _find_val(bs_data, 'total current assets')
        
        if nca_cy == 0 and ca_cy == 0:
            ax.text(0.5, 0.5, 'No Asset data available', ha='center', va='center',
                    fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
        else:
            values = [abs(nca_cy), abs(ca_cy)]
            labels = ['Non-Current Assets', 'Current Assets']
            colors = [BLUE, GREEN]
            
            wedges, texts, autotexts = ax.pie(
                values, labels=labels, colors=colors, autopct='%1.1f%%',
                startangle=90, textprops={'color': TEXT_COLOR, 'fontsize': 9},
                wedgeprops={'edgecolor': '#FFFFFF', 'linewidth': 1.5}
            )
            for t in autotexts:
                t.set_color('#FFFFFF')
                t.set_fontweight('bold')
        
        ax.set_title('Asset Composition (CY)', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_asset_composition.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 3: {e}")

    # 4. Liability & Equity Composition (CY)
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        eq_cy, _ = _find_val(bs_data, 'total equity')
        ncl_cy, _ = _find_val(bs_data, 'total non-current liabilities')
        if ncl_cy == 0:
            ncl_cy, _ = _find_val(bs_data, 'total non current liabilities')
        cl_cy, _ = _find_val(bs_data, 'total current liabilities')
        
        if eq_cy == 0 and ncl_cy == 0 and cl_cy == 0:
            ax.text(0.5, 0.5, 'No Liability & Equity data available', ha='center', va='center',
                    fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
        else:
            values = [abs(eq_cy), abs(ncl_cy), abs(cl_cy)]
            labels = ['Equity', 'Non-Current\nLiabilities', 'Current\nLiabilities']
            colors = [GREEN, YELLOW, RED]
            
            wedges, texts, autotexts = ax.pie(
                values, labels=labels, colors=colors, autopct='%1.1f%%',
                startangle=90, textprops={'color': TEXT_COLOR, 'fontsize': 9},
                wedgeprops={'edgecolor': '#FFFFFF', 'linewidth': 1.5}
            )
            for t in autotexts:
                t.set_color('#FFFFFF')
                t.set_fontweight('bold')
        
        ax.set_title('Liability & Equity Composition (CY)', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_liability_composition.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 4: {e}")

    # 5. Top 5 BS Variances
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        items = []
        for row in bs_data:
            pct = row.get('variance_pct', 0)
            if pct is not None and isinstance(pct, (int, float)) and row.get('particulars'):
                if not row.get('is_bold', False):
                    items.append((row['particulars'][:30], pct))
        
        items.sort(key=lambda x: abs(x[1]), reverse=True)
        top5 = items[:5]
        
        if not top5:
            ax.text(0.5, 0.5, 'No variance data available', ha='center', va='center',
                    fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
        else:
            names = [t[0] for t in reversed(top5)]
            values = [t[1] for t in reversed(top5)]
            colors = [GREEN if v >= 0 else RED for v in values]
            
            ax.barh(names, values, color=colors, edgecolor='none', height=0.6)
            ax.set_xlabel('Variance %', color=TEXT_COLOR, fontsize=9)
            ax.tick_params(axis='y', labelsize=8)
            ax.grid(True, axis='x', linestyle='--', alpha=0.5, color=GRID_COLOR)
        
        ax.set_title('Top 5 BS Variances', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_top_bs_variances.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 5: {e}")

    # 6. Top 5 PL Variances
    try:
        fig, ax = plt.subplots(figsize=(6, 4.2), dpi=150)
        _style_axes(ax)
        
        items = []
        for row in pl_data:
            pct = row.get('variance_pct', 0)
            if pct is not None and isinstance(pct, (int, float)) and row.get('particulars'):
                if not row.get('is_bold', False):
                    items.append((row['particulars'][:30], pct))
        
        items.sort(key=lambda x: abs(x[1]), reverse=True)
        top5 = items[:5]
        
        if not top5:
            ax.text(0.5, 0.5, 'No variance data available', ha='center', va='center',
                    fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
        else:
            names = [t[0] for t in reversed(top5)]
            values = [t[1] for t in reversed(top5)]
            colors = [GREEN if v >= 0 else RED for v in values]
            
            ax.barh(names, values, color=colors, edgecolor='none', height=0.6)
            ax.set_xlabel('Variance %', color=TEXT_COLOR, fontsize=9)
            ax.tick_params(axis='y', labelsize=8)
            ax.grid(True, axis='x', linestyle='--', alpha=0.5, color=GRID_COLOR)
        
        ax.set_title('Top 5 PL Variances', color=TEXT_COLOR, fontsize=11, fontweight='bold', pad=10)
        fig.tight_layout()
        path = os.path.join(temp_dir, 'chart_top_pl_variances.png')
        fig.savefig(path, facecolor='#FFFFFF', edgecolor='none', bbox_inches='tight')
        plt.close(fig)
        paths.append(path)
    except Exception as e:
        print(f"Error generating chart 6: {e}")

    return paths
