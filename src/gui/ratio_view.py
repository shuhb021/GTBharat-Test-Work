"""
FAR Automation Tool — Ratio Analysis View (Module 5)
Displays computed financial ratios with variance highlighting.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush


class RatioView(QWidget):
    """Module 5 — Financial Ratio Analysis."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ratios = []
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # Title
        title = QLabel('📐 Ratio Analysis')
        title.setObjectName('headingLabel')
        layout.addWidget(title)
        
        # Header
        header_frame = QFrame()
        header_frame.setObjectName('cardFrame')
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(16, 12, 16, 12)
        
        self.client_label = QLabel('Client: —')
        self.client_label.setObjectName('clientInfoLabel')
        header_layout.addWidget(self.client_label)
        
        self.period_label = QLabel('')
        self.period_label.setObjectName('mutedLabel')
        header_layout.addWidget(self.period_label)
        
        scope = QLabel('Scope: Percentage variance of +/- 10% selected for analysis')
        scope.setObjectName('mutedLabel')
        header_layout.addWidget(scope)
        
        layout.addWidget(header_frame)
        
        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            'Key Metric', 'Formula', 'CY Value', 'PY Value',
            'Change %', 'Remarks'
        ])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(1, 280)
        self.table.setColumnWidth(2, 120)
        self.table.setColumnWidth(3, 120)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 250)
        
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.table)
        
        # Placeholder
        self.placeholder = QLabel('Generate FAR from the Input Form to view ratios.')
        self.placeholder.setObjectName('mutedLabel')
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.placeholder)
    
    def load_data(self, ratios, client_name='', cy_year=None, py_year=None):
        """Populate the ratio table."""
        self.ratios = ratios
        self.placeholder.hide()
        
        self.client_label.setText(f'Client: {client_name}')
        if cy_year and py_year:
            self.period_label.setText(
                f'As at 31 March {cy_year} vs 31 March {py_year}')
        
        self.table.setRowCount(len(ratios))
        
        for i, ratio in enumerate(ratios):
            flag = ratio.get('flag', False)
            
            items = [
                ratio.get('key', ''),
                ratio.get('formula', ''),
                f"{ratio.get('cy', 0):.2f}",
                f"{ratio.get('py', 0):.2f}",
                ratio.get('display_change', ''),
            ]
            
            for j, val in enumerate(items):
                item = QTableWidgetItem(str(val))
                
                if j == 0:
                    item.setFont(QFont('Segoe UI', 11, QFont.Weight.Bold))
                
                if j >= 2:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                
                if flag:
                    item.setBackground(QBrush(QColor('#EAB308')))
                    item.setForeground(QBrush(QColor('#000000')))
                
                # Color change %
                if j == 4 and not flag:
                    change = ratio.get('change', 0)
                    if change > 0:
                        item.setForeground(QBrush(QColor('#22C55E')))
                    elif change < 0:
                        item.setForeground(QBrush(QColor('#EF4444')))
                
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(i, j, item)
            
            # Editable remarks
            self.table.setItem(i, 5, QTableWidgetItem(''))
