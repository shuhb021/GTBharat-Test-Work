"""
FAR Automation Tool — YoY Financial Comparison View
Displays side-by-side comparison of actual vs annualized values, with metrics and normalization.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView, QFrame, QTabBar
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush


class ComparisonView(QWidget):
    """
    Dedicated view for Year-over-Year Financial Comparison.
    Provides coverage metrics and normalized values comparison.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.bs_data = []
        self.pl_data = []
        self.meta = {}
        self.client_name = ''
        self.rounding_unit = 'Lakhs'
        self.current_tab = 0  # 0 for BS, 1 for PL
        self._setup_ui()
        
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # ── Header Card ──
        self.header_frame = QFrame()
        self.header_frame.setObjectName('cardFrame')
        header_layout = QVBoxLayout(self.header_frame)
        header_layout.setContentsMargins(16, 12, 16, 12)
        header_layout.setSpacing(8)
        
        # Row 1: Client and Title
        row1 = QHBoxLayout()
        self.client_label = QLabel('Client: —')
        self.client_label.setObjectName('clientInfoLabel')
        row1.addWidget(self.client_label)
        
        title_label = QLabel('⚖️ Year-over-Year Financial Comparison')
        title_label.setObjectName('subheadingLabel')
        row1.addWidget(title_label)
        row1.addStretch()
        header_layout.addLayout(row1)
        
        # Row 2: Periods and Metrics (Side-by-side details)
        periods_layout = QHBoxLayout()
        
        self.py_details = QLabel('Previous Year: —')
        self.py_details.setObjectName('mutedLabel')
        periods_layout.addWidget(self.py_details)
        
        periods_layout.addSpacing(20)
        
        self.cy_details = QLabel('Current Year: —')
        self.cy_details.setObjectName('mutedLabel')
        periods_layout.addWidget(self.cy_details)
        periods_layout.addStretch()
        header_layout.addLayout(periods_layout)
        
        # Row 3: Warning Banner
        self.warning_banner = QLabel("⚠️ Selected periods are different. Results are normalized to a full-year equivalent basis.")
        self.warning_banner.setStyleSheet("""
            background-color: #FFFBEB;
            color: #B45309;
            border: 1px solid #FDE68A;
            border-radius: 4px;
            padding: 8px 12px;
            font-weight: 600;
            font-size: 12px;
        """)
        self.warning_banner.setVisible(False)
        header_layout.addWidget(self.warning_banner)
        
        layout.addWidget(self.header_frame)
        
        # ── Tab Selector (Beautiful QTabBar styled via QSS) ──
        tab_layout = QHBoxLayout()
        self.tab_bar = QTabBar()
        self.tab_bar.addTab('Balance Sheet')
        self.tab_bar.addTab('Profit & Loss')
        self.tab_bar.setShape(QTabBar.Shape.RoundedNorth)
        self.tab_bar.currentChanged.connect(self._on_tab_changed)
        tab_layout.addWidget(self.tab_bar)
        
        self.unit_label = QLabel('')
        self.unit_label.setObjectName('mutedLabel')
        self.unit_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        tab_layout.addWidget(self.unit_label)
        
        layout.addLayout(tab_layout)
        
        # ── Comparison Table ──
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            'Particulars', 
            'PY Actual', 'PY Annualized', 
            'CY Actual', 'CY Annualized', 
            'Growth % (Norm)', 'Variance % (Act)'
        ])
        
        # Column resizing modes and widths
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(1, 120)
        self.table.setColumnWidth(2, 125)
        self.table.setColumnWidth(3, 120)
        self.table.setColumnWidth(4, 125)
        self.table.setColumnWidth(5, 125)
        self.table.setColumnWidth(6, 125)
        
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(True)
        
        layout.addWidget(self.table)
        
    def _on_tab_changed(self, index):
        self.current_tab = index
        self._refresh_table()
        
    def load_data(self, app_data):
        """Load Balance Sheet and Profit & Loss dataset with metrics."""
        self.bs_data = app_data.get('bs_result', [])
        self.pl_data = app_data.get('pl_result', [])
        self.meta = app_data.get('meta', {})
        self.client_name = app_data.get('client_name', '')
        self.rounding_unit = app_data.get('rounding_unit', 'Lakhs')
        
        # Update Header Details
        self.client_label.setText(f'Client: {self.client_name}')
        self.unit_label.setText(f'(Rs. in {self.rounding_unit})')
        
        # Parse comparison periods metadata
        cy_days = self.meta.get('cy_days', 365)
        py_days = self.meta.get('py_days', 365)
        cy_months = self.meta.get('cy_months', 12.0)
        py_months = self.meta.get('py_months', 12.0)
        cy_cov = self.meta.get('cy_coverage', 100.0)
        py_cov = self.meta.get('py_coverage', 100.0)
        
        cy_yr = self.meta.get('cy_year', 2025)
        py_yr = self.meta.get('py_year', 2024)
        
        self.py_details.setText(
            f'<b>Previous Year ({py_yr}):</b> Days: {py_days} | Months: {py_months} | Coverage: {py_cov}%'
        )
        self.cy_details.setText(
            f'<b>Current Year ({cy_yr}):</b> Days: {cy_days} | Months: {cy_months} | Coverage: {cy_cov}%'
        )
        
        # Show normalization warning if auto-annualization was triggered
        is_annualized = self.meta.get('is_annualized', False)
        self.warning_banner.setVisible(is_annualized)
        
        # Refresh table contents
        self._refresh_table()
        
    def _refresh_table(self):
        data = self.pl_data if self.current_tab == 1 else self.bs_data
        self.table.setRowCount(len(data))
        
        cy_factor = self.meta.get('cy_factor', 1.0)
        py_factor = self.meta.get('py_factor', 1.0)
        is_pl = (self.current_tab == 1)
        
        for i, row in enumerate(data):
            particulars = row.get('particulars', '')
            is_bold = row.get('is_bold', False)
            flag = row.get('flag', False)
            
            # Values: Actual and Annualized
            # Note: Balance sheet items are not annualized even if the toggle is checked.
            # In BS: Annualized Value = Actual Value.
            # In PL: Annualized Value = Actual Value * factor (if factor is set).
            py_act = row.get('py_actual', row.get('py'))
            cy_act = row.get('cy_actual', row.get('cy'))
            
            # Normalized/Annualized values are stored in 'py' and 'cy' inside DataEngine
            py_ann = row.get('py')
            cy_ann = row.get('cy')
            
            # Growth % (on annualized values)
            if py_ann in (0, 0.0, None, '-') or cy_ann in (None, '-'):
                growth_pct = 0.0
                growth_str = 'N/A'
            else:
                try:
                    growth_pct = ((float(cy_ann) - float(py_ann)) / abs(float(py_ann))) * 100.0
                    growth_str = f'{growth_pct:+.1f}%'
                except Exception:
                    growth_pct = 0.0
                    growth_str = 'N/A'
                
            # Variance % (on raw actual values)
            if py_act in (0, 0.0, None, '-') or cy_act in (None, '-'):
                var_pct = 0.0
                var_str = 'N/A'
            else:
                try:
                    var_pct = ((float(cy_act) - float(py_act)) / abs(float(py_act))) * 100.0
                    var_str = f'{var_pct:+.1f}%'
                except Exception:
                    var_pct = 0.0
                    var_str = 'N/A'
                
            items = [
                particulars,
                self._format_number(py_act),
                self._format_number(py_ann),
                self._format_number(cy_act),
                self._format_number(cy_ann),
                growth_str,
                var_str
            ]
            
            for j, val in enumerate(items):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                
                # Apply styling
                if is_bold:
                    item.setFont(QFont('Segoe UI', 11, QFont.Weight.Bold))
                    item.setBackground(QBrush(QColor('#F5F0FA')))
                    
                if flag:
                    item.setBackground(QBrush(QColor('#EAB308')))
                    item.setForeground(QBrush(QColor('#000000')))
                    
                # Format Growth % and Variance % text colors
                if j in (5, 6) and val != 'N/A':
                    val_float = growth_pct if j == 5 else var_pct
                    if val_float > 0:
                        item.setForeground(QBrush(QColor('#22C55E')))
                    elif val_float < 0:
                        item.setForeground(QBrush(QColor('#EF4444')))
                    if flag:
                        item.setForeground(QBrush(QColor('#000000')))
                        
                # Right-align number columns
                if j >= 1:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    
                self.table.setItem(i, j, item)
                
    def _format_number(self, value):
        if value is None or value == '-':
            return '-'
        try:
            v = float(value)
            if v == 0:
                return '0'
            if self.rounding_unit in ('Lakhs', 'Millions', 'Actuals'):
                return f'{v:,.2f}'
            return f'{v:,.0f}'
        except (ValueError, TypeError):
            return str(value)
