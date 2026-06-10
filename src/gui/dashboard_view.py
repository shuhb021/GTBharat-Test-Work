"""
FAR Automation Tool — Visual Analytics Dashboard (Module 6)
Matplotlib charts embedded in PyQt6 showing financial analysis visualizations.
"""

import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QPushButton, QFrame, QFileDialog, QScrollArea
)
from PyQt6.QtCore import Qt

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt


# Chart colors
BLUE = '#4A1A6B'
GREEN = '#22C55E'
YELLOW = '#EAB308'
RED = '#EF4444'
BG_MAIN = '#FFFFFF'
BG_CARD = '#FFFFFF'
TEXT_COLOR = '#555555'
BORDER = '#E0E0E0'


class ChartCard(QFrame):
    """A card container for a single chart."""
    
    def __init__(self, title='', parent=None):
        super().__init__(parent)
        self.setObjectName('cardFrame')
        self.setObjectName('chartCardFrame')
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)
        
        if title:
            label = QLabel(title)
            label.setObjectName('chartTitle')
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(label)
        
        # Chart canvas
        self.figure = Figure(figsize=(5, 3.5), dpi=100)
        self.figure.patch.set_facecolor(BG_CARD)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
    
    def get_ax(self):
        """Get or create the axes."""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor(BG_CARD)
        ax.tick_params(colors=TEXT_COLOR, labelsize=9)
        for spine in ax.spines.values():
            spine.set_color(BORDER)
        ax.title.set_color('#4A1A6B')
        ax.xaxis.label.set_color(TEXT_COLOR)
        ax.yaxis.label.set_color(TEXT_COLOR)
        return ax
    
    def refresh(self):
        self.figure.tight_layout(pad=1.5)
        self.canvas.draw()


class DashboardView(QWidget):
    """Module 6 — Visual Analytics Dashboard with 6 charts."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.charts = {}
        self._setup_ui()
    
    def _setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        content = QWidget()
        main_layout = QVBoxLayout(content)
        main_layout.setContentsMargins(24, 16, 24, 16)
        main_layout.setSpacing(12)
        
        # Title bar
        title_row = QHBoxLayout()
        title = QLabel('📉 Visual Analytics Dashboard')
        title.setObjectName('headingLabel')
        title_row.addWidget(title)
        title_row.addStretch()
        
        self.export_btn = QPushButton('📷 Export Charts as PNG')
        self.export_btn.setObjectName('secondaryButton')
        self.export_btn.clicked.connect(self._export_charts)
        title_row.addWidget(self.export_btn)
        
        main_layout.addLayout(title_row)
        
        self.client_label = QLabel('')
        self.client_label.setObjectName('mutedLabel')
        main_layout.addWidget(self.client_label)
        
        # 2x3 grid of charts
        grid = QGridLayout()
        grid.setSpacing(12)
        
        chart_titles = [
            'Revenue vs Expenses (CY/PY)',
            'Profit Before Tax Trend',
            'Asset Composition (CY)',
            'Liability & Equity Composition (CY)',
            'Top 5 BS Variances',
            'Top 5 PL Variances'
        ]
        
        for i, t in enumerate(chart_titles):
            card = ChartCard(t)
            row, col = divmod(i, 2)
            grid.addWidget(card, row, col)
            self.charts[i] = card
        
        main_layout.addLayout(grid)
        main_layout.addStretch()
        
        scroll.setWidget(content)
        
        wrapper = QVBoxLayout(self)
        wrapper.setContentsMargins(0, 0, 0, 0)
        wrapper.addWidget(scroll)
        
        # Placeholder
        self.placeholder = QLabel('Generate FAR from the Input Form to view dashboard.')
        self.placeholder.setObjectName('mutedLabel')
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
    
    def load_data(self, bs_data, pl_data, ratios, client_name='',
                  cy_year=None, py_year=None, rounding_unit='Lakhs'):
        """Generate all 6 charts from data."""
        self.client_label.setText(
            f'Client: {client_name}  |  FY {cy_year}  |  (Rs. in {rounding_unit})')
        
        self._chart_revenue_expenses(bs_data, pl_data, cy_year, py_year)
        self._chart_pbt_trend(pl_data, cy_year, py_year)
        self._chart_asset_composition(bs_data)
        self._chart_liability_composition(bs_data)
        self._chart_top_variances(bs_data, 4, 'BS')
        self._chart_top_variances(pl_data, 5, 'PL')
    
    def _find_val(self, data, keyword):
        """Find CY and PY values from data by keyword."""
        kw = keyword.lower()
        for row in data:
            if kw in row.get('particulars', '').lower():
                return float(row.get('cy', 0) or 0), float(row.get('py', 0) or 0)
        return 0, 0
    
    def _chart_revenue_expenses(self, bs_data, pl_data, cy_year, py_year):
        """Chart 1: Grouped bar — Revenue, Expenses, PBT."""
        card = self.charts[0]
        ax = card.get_ax()
        
        rev_cy, rev_py = self._find_val(pl_data, 'total income')
        exp_cy, exp_py = self._find_val(pl_data, 'total expense')
        pbt_cy, pbt_py = 0, 0
        for row in pl_data:
            p = row.get('particulars', '').lower()
            if 'profit' in p and 'before' in p and 'tax' in p:
                pbt_cy = float(row.get('cy', 0) or 0)
                pbt_py = float(row.get('py', 0) or 0)
                break
        
        import numpy as np
        x = np.arange(3)
        width = 0.35
        
        cy_vals = [rev_cy, exp_cy, pbt_cy]
        py_vals = [rev_py, exp_py, pbt_py]
        
        bars1 = ax.bar(x - width/2, cy_vals, width, label=f'CY ({cy_year})',
                       color=BLUE, edgecolor='none')
        bars2 = ax.bar(x + width/2, py_vals, width, label=f'PY ({py_year})',
                       color=GREEN, alpha=0.7, edgecolor='none')
        
        ax.set_xticks(x)
        ax.set_xticklabels(['Revenue', 'Expenses', 'PBT'], color=TEXT_COLOR)
        ax.legend(facecolor=BG_CARD, edgecolor=BORDER, labelcolor=TEXT_COLOR, fontsize=9)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/1e6:.1f}M' if abs(x) >= 1e6 else f'{x/1e3:.0f}K'))
        
        card.refresh()
    
    def _chart_pbt_trend(self, pl_data, cy_year, py_year):
        """Chart 2: Line — PBT trend."""
        card = self.charts[1]
        ax = card.get_ax()
        
        pbt_cy, pbt_py = 0, 0
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
        ax.fill_between(years, values, alpha=0.1, color=BLUE)
        
        for i, (year, val) in enumerate(zip(years, values)):
            ax.annotate(f'{val:,.0f}', (year, val), textcoords="offset points",
                       xytext=(0, 12), ha='center', fontsize=9, color=TEXT_COLOR)
        
        ax.set_ylabel('Amount', color=TEXT_COLOR)
        card.refresh()
    
    def _chart_asset_composition(self, bs_data):
        """Chart 3: Pie — Asset composition (CY)."""
        card = self.charts[2]
        ax = card.get_ax()
        
        nca_cy, _ = self._find_val(bs_data, 'total non-current assets')
        if nca_cy == 0:
            nca_cy, _ = self._find_val(bs_data, 'total non current assets')
        ca_cy, _ = self._find_val(bs_data, 'total current assets')
        
        if nca_cy == 0 and ca_cy == 0:
            ax.text(0.5, 0.5, 'No data', ha='center', va='center',
                   fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
            card.refresh()
            return
        
        values = [abs(nca_cy), abs(ca_cy)]
        labels = ['Non-Current\nAssets', 'Current\nAssets']
        colors = [BLUE, GREEN]
        
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, colors=colors, autopct='%1.1f%%',
            startangle=90, textprops={'color': TEXT_COLOR, 'fontsize': 9})
        for t in autotexts:
            t.set_color('#FFFFFF')
            t.set_fontweight('bold')
        
        card.refresh()
    
    def _chart_liability_composition(self, bs_data):
        """Chart 4: Pie — Liability + Equity composition (CY)."""
        card = self.charts[3]
        ax = card.get_ax()
        
        eq_cy, _ = self._find_val(bs_data, 'total equity')
        ncl_cy, _ = self._find_val(bs_data, 'total non-current liabilities')
        if ncl_cy == 0:
            ncl_cy, _ = self._find_val(bs_data, 'total non current liabilities')
        cl_cy, _ = self._find_val(bs_data, 'total current liabilities')
        
        if eq_cy == 0 and ncl_cy == 0 and cl_cy == 0:
            ax.text(0.5, 0.5, 'No data', ha='center', va='center',
                   fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
            card.refresh()
            return
        
        values = [abs(eq_cy), abs(ncl_cy), abs(cl_cy)]
        labels = ['Equity', 'Non-Current\nLiabilities', 'Current\nLiabilities']
        colors = [GREEN, YELLOW, RED]
        
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, colors=colors, autopct='%1.1f%%',
            startangle=90, textprops={'color': TEXT_COLOR, 'fontsize': 9})
        for t in autotexts:
            t.set_color('#FFFFFF')
            t.set_fontweight('bold')
        
        card.refresh()
    
    def _chart_top_variances(self, data, chart_idx, label):
        """Chart 5/6: Horizontal bar — Top 5 variances."""
        card = self.charts[chart_idx]
        ax = card.get_ax()
        
        # Get top 5 by absolute variance %
        items = []
        for row in data:
            pct = row.get('variance_pct', 0)
            if pct is not None and isinstance(pct, (int, float)) and row.get('particulars'):
                if not row.get('is_bold', False):
                    items.append((row['particulars'][:30], pct))
        
        items.sort(key=lambda x: abs(x[1]), reverse=True)
        top5 = items[:5]
        
        if not top5:
            ax.text(0.5, 0.5, 'No variance data', ha='center', va='center',
                   fontsize=12, color=TEXT_COLOR, transform=ax.transAxes)
            card.refresh()
            return
        
        names = [t[0] for t in reversed(top5)]
        values = [t[1] for t in reversed(top5)]
        colors = [GREEN if v >= 0 else RED for v in values]
        
        ax.barh(names, values, color=colors, edgecolor='none', height=0.6)
        ax.set_xlabel('Variance %', color=TEXT_COLOR, fontsize=9)
        ax.tick_params(axis='y', labelsize=8)
        
        card.refresh()
    
    def _export_charts(self):
        """Export all charts as PNG files."""
        folder = QFileDialog.getExistingDirectory(self, 'Select Export Folder')
        if not folder:
            return
        
        chart_names = [
            'revenue_expenses', 'pbt_trend', 'asset_composition',
            'liability_composition', 'top_bs_variances', 'top_pl_variances'
        ]
        
        for i, name in enumerate(chart_names):
            filepath = os.path.join(folder, f'chart_{name}.png')
            self.charts[i].figure.savefig(filepath, dpi=150,
                                          facecolor=BG_MAIN, edgecolor='none',
                                          bbox_inches='tight')
