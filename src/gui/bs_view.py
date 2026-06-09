"""
FAR Automation Tool — Balance Sheet View (Module 2)
Displays parsed BS data with variance analysis in a styled table.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView, QFrame, QLineEdit, QScrollArea
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush


class AnalysisTableView(QWidget):
    """
    Reusable analysis table view for BS and PL modules.
    Shows header info, legends, and a styled variance table.
    """
    
    def __init__(self, title, date_prefix='As at', parent=None):
        super().__init__(parent)
        self.title = title
        self.date_prefix = date_prefix
        self.data = []
        self.remarks = {}
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # ── Header Section ──
        header_frame = QFrame()
        header_frame.setObjectName('cardFrame')
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(16, 12, 16, 12)
        
        # Left: Title info
        left_layout = QVBoxLayout()
        
        self.client_label = QLabel('Client: —')
        self.client_label.setStyleSheet("color: #FFFFFF; font-size: 14px; font-weight: 600; background: transparent;")
        left_layout.addWidget(self.client_label)
        
        self.title_label = QLabel(self.title)
        self.title_label.setObjectName('subheadingLabel')
        left_layout.addWidget(self.title_label)
        
        self.period_label = QLabel('')
        self.period_label.setObjectName('mutedLabel')
        left_layout.addWidget(self.period_label)
        
        self.scope_label = QLabel('Scope: All variances above TE and (+/-) 10% selected')
        self.scope_label.setObjectName('mutedLabel')
        left_layout.addWidget(self.scope_label)
        
        header_layout.addLayout(left_layout)
        header_layout.addStretch()
        
        # Right: Legends
        right_layout = QVBoxLayout()
        legends_title = QLabel('Legends')
        legends_title.setStyleSheet("color: #007ACC; font-weight: 700; font-size: 12px; background: transparent;")
        right_layout.addWidget(legends_title)
        
        legends = [
            ('a', 'Traced to Financials CY'),
            ('b', 'Traced to Financials PY'),
            ('e', 'Recomputed'),
        ]
        for key, desc in legends:
            row = QHBoxLayout()
            k = QLabel(key)
            k.setStyleSheet("color: #007ACC; font-weight: 700; font-size: 11px; min-width: 14px; background: transparent;")
            d = QLabel(f'= {desc}')
            d.setStyleSheet("color: #888888; font-size: 11px; background: transparent;")
            row.addWidget(k)
            row.addWidget(d)
            row.addStretch()
            right_layout.addLayout(row)
        
        header_layout.addLayout(right_layout)
        layout.addWidget(header_frame)
        
        # ── Summary Stats ──
        stats_frame = QFrame()
        stats_layout = QHBoxLayout(stats_frame)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        
        self.total_label = QLabel('Total items: 0')
        self.total_label.setObjectName('mutedLabel')
        stats_layout.addWidget(self.total_label)
        
        self.flagged_label = QLabel('Flagged (≥10%): 0')
        self.flagged_label.setStyleSheet("color: #EAB308; font-size: 12px; background: transparent;")
        stats_layout.addWidget(self.flagged_label)
        
        stats_layout.addStretch()
        
        self.unit_label = QLabel('')
        self.unit_label.setObjectName('mutedLabel')
        stats_layout.addWidget(self.unit_label)
        
        layout.addWidget(stats_frame)
        
        # ── Data Table ──
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            'Particulars', 'Note No.', 'CY (b)', 'PY (a)',
            'Var. Abs (e)', 'Var. % (e)', 'Remarks'
        ])
        
        # Column widths
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(1, 80)
        self.table.setColumnWidth(2, 130)
        self.table.setColumnWidth(3, 130)
        self.table.setColumnWidth(4, 130)
        self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 280)
        
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(True)
        
        layout.addWidget(self.table)
    
    def load_data(self, data, client_name='', cy_year=None, py_year=None,
                  rounding_unit='Lakhs', remarks=None):
        """Populate the table with parsed and computed data."""
        self.data = data
        self.remarks = remarks or {}
        
        self.client_label.setText(f'Client: {client_name}')
        self.rounding_unit = rounding_unit
        if cy_year and py_year:
            self.period_label.setText(
                f'{self.date_prefix} 31 March {cy_year} vs 31 March {py_year}')
        self.unit_label.setText(f'(Rs. in {rounding_unit})')
        
        # Update table
        self.table.setRowCount(len(data))
        
        flagged_count = 0
        for i, row in enumerate(data):
            flag = row.get('flag', False)
            is_bold = row.get('is_bold', False)
            if flag:
                flagged_count += 1
            
            # Set items
            items = [
                row.get('particulars', ''),
                str(row.get('note', '')),
                self._format_number(row.get('cy', 0)),
                self._format_number(row.get('py', 0)),
                self._format_number(row.get('variance_abs', 0)),
                row.get('display_pct', ''),
            ]
            
            for j, val in enumerate(items):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                
                # Styling
                if is_bold:
                    font = QFont('Segoe UI', 11, QFont.Weight.Bold)
                    item.setFont(font)
                    item.setBackground(QBrush(QColor('#2A2D2E')))
                
                if flag:
                    item.setBackground(QBrush(QColor('#EAB308')))
                    item.setForeground(QBrush(QColor('#000000')))
                
                # Color variance %
                if j == 5:
                    pct = row.get('variance_pct', 0)
                    if isinstance(pct, (int, float)):
                        if pct > 0:
                            item.setForeground(QBrush(QColor('#22C55E')))
                        elif pct < 0:
                            item.setForeground(QBrush(QColor('#EF4444')))
                        if flag:
                            item.setForeground(QBrush(QColor('#000000')))
                
                # Right-align numbers
                if j >= 2:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                
                self.table.setItem(i, j, item)
            
            # Remarks column — editable
            remark_text = self.remarks.get(i, row.get('remark', ''))
            remark_item = QTableWidgetItem(str(remark_text))
            self.table.setItem(i, 6, remark_item)
        
        # Update stats
        self.total_label.setText(f'Total items: {len(data)}')
        self.flagged_label.setText(f'Flagged (≥10%): {flagged_count}')
    
    def _format_number(self, value):
        """Format a number with commas."""
        if value is None:
            return '—'
        try:
            v = float(value)
            if v == 0:
                return '0'
            unit = getattr(self, 'rounding_unit', 'Lakhs')
            if unit in ('Lakhs', 'Millions', 'Actuals'):
                return f'{v:,.2f}'
            return f'{v:,.0f}'
        except (ValueError, TypeError):
            return str(value)
    
    def get_remarks(self):
        """Get current remarks from the table."""
        remarks = {}
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 6)
            if item and item.text().strip():
                remarks[i] = item.text().strip()
        return remarks
    
    def update_remark(self, row_idx, text):
        """Update a single remark cell."""
        if 0 <= row_idx < self.table.rowCount():
            item = self.table.item(row_idx, 6)
            if item:
                item.setText(text)
            else:
                self.table.setItem(row_idx, 6, QTableWidgetItem(text))


class BSView(AnalysisTableView):
    """Module 2 — Balance Sheet Analysis View."""
    
    def __init__(self, parent=None):
        super().__init__(
            title='Analysis of Balance Sheet',
            date_prefix='As at',
            parent=parent
        )


class PLView(AnalysisTableView):
    """Module 3 — Profit & Loss Analysis View."""
    
    def __init__(self, parent=None):
        super().__init__(
            title='Analysis of Statement of Profit and Loss',
            date_prefix='For the Year ended',
            parent=parent
        )
