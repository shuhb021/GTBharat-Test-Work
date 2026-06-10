"""
FAR Automation Tool — Notes to Accounts View (Module 4)
Tab widget showing parsed notes data with variance analysis.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QTableWidget,
    QTableWidgetItem, QLabel, QHeaderView, QFrame, QScrollArea
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor, QBrush


class NotesView(QWidget):
    """Module 4 — Notes to Accounts with tabbed display."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # Title
        title = QLabel('📋 Analysis of Notes to Accounts')
        title.setObjectName('headingLabel')
        layout.addWidget(title)
        
        # Header with legends
        header_frame = QFrame()
        header_frame.setObjectName('cardFrame')
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(16, 12, 16, 12)
        
        left = QVBoxLayout()
        self.client_label = QLabel('Client: —')
        self.client_label.setObjectName('clientInfoLabel')
        left.addWidget(self.client_label)
        self.period_label = QLabel('')
        self.period_label.setObjectName('mutedLabel')
        left.addWidget(self.period_label)
        header_layout.addLayout(left)
        header_layout.addStretch()
        
        # Legends
        right = QVBoxLayout()
        legends_title = QLabel('Legend:')
        legends_title.setObjectName('legendTitle')
        right.addWidget(legends_title)
        for key, desc in [('a', 'Traced to Financials CY'), ('b', 'Traced to Financials PY'),
                          ('j', 'Linked'), ('e', 'Recomputed'), ('h', 'Immaterial')]:
            row = QHBoxLayout()
            k = QLabel(key)
            k.setObjectName('legendKeySmall')
            d = QLabel(f'= {desc}')
            d.setObjectName('legendDescSmall')
            row.addWidget(k)
            row.addWidget(d)
            row.addStretch()
            right.addLayout(row)
        header_layout.addLayout(right)
        
        layout.addWidget(header_frame)
        
        # Tab widget for note sheets
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Placeholder
        self.placeholder = QLabel('Upload notes files in the Input Form to view analysis here.')
        self.placeholder.setObjectName('mutedLabel')
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.placeholder)
    
    def load_data(self, notes_data, client_name='', cy_year=None, py_year=None,
                  rounding_unit='Lakhs'):
        """
        Populate tabs with notes data.
        notes_data: dict mapping sheet_name → list of note groups
        """
        self.client_label.setText(f'Client: {client_name}')
        if cy_year and py_year:
            self.period_label.setText(f'As at 31 March {cy_year} vs 31 March {py_year}')
        
        self.tab_widget.clear()
        self.placeholder.hide()
        
        if not notes_data:
            self.placeholder.show()
            return
        
        for sheet_name, note_groups in notes_data.items():
            tab = self._create_notes_tab(note_groups, cy_year, py_year, rounding_unit)
            self.tab_widget.addTab(tab, f'Notes {sheet_name}')
    
    def _create_notes_tab(self, note_groups, cy_year, py_year, rounding_unit):
        """Create a tab with tables for each note group."""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(16)
        
        for note_group in note_groups:
            # Note heading
            heading = QLabel(f"Note {note_group['note_num']}: {note_group['note_heading']}")
            heading.setObjectName('noteHeading')
            layout.addWidget(heading)
            
            data = note_group.get('data', [])
            if not data:
                empty = QLabel('No data rows found for this note.')
                empty.setObjectName('mutedLabel')
                layout.addWidget(empty)
                continue
            
            # Create table
            table = QTableWidget()
            table.setColumnCount(6)
            
            cy_label = f'As at 31 March {cy_year}' if cy_year else 'CY'
            py_label = f'As at 31 March {py_year}' if py_year else 'PY'
            table.setHorizontalHeaderLabels([
                'Particulars', cy_label, py_label,
                'Variance', 'Variance %', 'Remarks'
            ])
            
            header = table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            table.setColumnWidth(1, 120)
            table.setColumnWidth(2, 120)
            table.setColumnWidth(3, 120)
            table.setColumnWidth(4, 100)
            table.setColumnWidth(5, 250)
            
            table.setAlternatingRowColors(True)
            table.verticalHeader().setVisible(False)
            table.setRowCount(len(data))
            
            # Compute variances for notes data
            from src.core.cpp_bridge import compute_variances
            computed = compute_variances(data)
            
            for i, row in enumerate(computed):
                flag = row.get('flag', False)
                is_bold = row.get('is_bold', False)
                
                items = [
                    row.get('particulars', ''),
                    self._fmt(row.get('cy', 0), rounding_unit),
                    self._fmt(row.get('py', 0), rounding_unit),
                    self._fmt(row.get('variance_abs', 0), rounding_unit),
                    row.get('display_pct', ''),
                ]
                
                for j, val in enumerate(items):
                    item = QTableWidgetItem(str(val))
                    if j > 0:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    if is_bold:
                        item.setFont(QFont('Segoe UI', 11, QFont.Weight.Bold))
                        item.setBackground(QBrush(QColor('#F5F0FA')))
                    if flag:
                        item.setBackground(QBrush(QColor('#EAB308')))
                        item.setForeground(QBrush(QColor('#000000')))
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(i, j, item)
                
                # Editable remarks
                table.setItem(i, 5, QTableWidgetItem(''))
            
            # Set fixed height based on rows
            table.setMinimumHeight(min(len(data) * 30 + 40, 400))
            layout.addWidget(table)
        
        layout.addStretch()
        scroll.setWidget(content)
        return scroll
    
    def _fmt(self, value, rounding_unit='Lakhs'):
        try:
            v = float(value)
            if v == 0:
                return '0'
            if rounding_unit in ('Lakhs', 'Millions', 'Actuals'):
                return f'{v:,.2f}'
            return f'{v:,.0f}'
        except (ValueError, TypeError):
            return str(value)
