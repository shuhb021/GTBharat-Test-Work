"""
FAR Automation Tool — AI Remarks View (Module 7)
Manages AI remark generation for flagged variance items.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QLabel, QPushButton, QProgressBar, QFrame, QHeaderView, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor, QBrush


class RemarkWorker(QThread):
    """Background worker for AI remark generation."""
    progress = pyqtSignal(int, int)   # current, total
    remarkReady = pyqtSignal(int, str, str)  # index, remark, statement_type
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, api_key, bs_items, pl_items, form_data):
        super().__init__()
        self.api_key = api_key
        self.bs_items = bs_items
        self.pl_items = pl_items
        self.form_data = form_data
    
    def run(self):
        try:
            from src.core.finance_agent import generate_remarks
            
            total_flagged = sum(1 for x in self.bs_items if x.get('flag')) + \
                           sum(1 for x in self.pl_items if x.get('flag'))
            completed = [0]
            
            def on_progress(current, total):
                completed[0] += 1
                self.progress.emit(completed[0], total_flagged)
            
            # Generate BS remarks
            bs_remarks = generate_remarks(
                self.api_key, self.bs_items,
                self.form_data.get('client_name', ''),
                self.form_data.get('financial_year', ''),
                'Balance Sheet',
                self.form_data.get('rounding_unit', 'Lakhs'),
                progress_callback=on_progress
            )
            for idx, remark in bs_remarks.items():
                self.remarkReady.emit(idx, remark, 'BS')
            
            # Generate PL remarks
            pl_remarks = generate_remarks(
                self.api_key, self.pl_items,
                self.form_data.get('client_name', ''),
                self.form_data.get('financial_year', ''),
                'Profit & Loss',
                self.form_data.get('rounding_unit', 'Lakhs'),
                progress_callback=on_progress
            )
            for idx, remark in pl_remarks.items():
                self.remarkReady.emit(idx, remark, 'PL')
            
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class RemarksView(QWidget):
    """Module 7 — AI Remarks generation and management."""
    
    remarksGenerated = pyqtSignal(dict, dict)  # bs_remarks, pl_remarks
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.bs_data = []
        self.pl_data = []
        self.bs_remarks = {}
        self.pl_remarks = {}
        self.form_data = {}
        self.worker = None
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # Title
        title = QLabel('🤖 AI-Powered Audit Remarks')
        title.setObjectName('headingLabel')
        layout.addWidget(title)
        
        subtitle = QLabel('Generate professional audit remarks for items with significant variances (≥10%)')
        subtitle.setObjectName('mutedLabel')
        layout.addWidget(subtitle)
        
        # Controls
        controls = QHBoxLayout()
        
        self.generate_btn = QPushButton('🚀 Generate AI Remarks')
        self.generate_btn.setMinimumWidth(200)
        self.generate_btn.setFixedHeight(40)
        self.generate_btn.setFont(QFont('Segoe UI', 12, QFont.Weight.Bold))
        self.generate_btn.clicked.connect(self._start_generation)
        controls.addWidget(self.generate_btn)
        
        self.regen_all_btn = QPushButton('🔄 Regenerate All')
        self.regen_all_btn.setObjectName('secondaryButton')
        self.regen_all_btn.clicked.connect(self._start_generation)
        self.regen_all_btn.setEnabled(False)
        controls.addWidget(self.regen_all_btn)
        
        controls.addStretch()
        
        self.status_label = QLabel('')
        self.status_label.setObjectName('mutedLabel')
        controls.addWidget(self.status_label)
        
        layout.addLayout(controls)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(20)
        layout.addWidget(self.progress_bar)
        
        # Stats
        stats = QHBoxLayout()
        self.bs_count = QLabel('BS flagged: 0')
        self.bs_count.setObjectName('mutedLabel')
        stats.addWidget(self.bs_count)
        self.pl_count = QLabel('PL flagged: 0')
        self.pl_count.setObjectName('mutedLabel')
        stats.addWidget(self.pl_count)
        self.total_generated = QLabel('Generated: 0')
        self.total_generated.setObjectName('generatedCount')
        stats.addWidget(self.total_generated)
        stats.addStretch()
        layout.addLayout(stats)
        
        # Results table
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            'Type', 'Line Item', 'Variance %', 'Remark', 'Status', 'Action'
        ])
        
        header = self.table.horizontalHeader()
        self.table.setColumnWidth(0, 50)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 350)
        self.table.setColumnWidth(4, 80)
        self.table.setColumnWidth(5, 80)
        
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)
        
        # Placeholder
        self.placeholder = QLabel('Generate FAR from Input Form first, then click "Generate AI Remarks".')
        self.placeholder.setObjectName('mutedLabel')
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.placeholder)
    
    def set_data(self, bs_data, pl_data, form_data):
        """Set the data to generate remarks for."""
        self.bs_data = bs_data
        self.pl_data = pl_data
        self.form_data = form_data
        
        bs_flagged = sum(1 for x in bs_data if x.get('flag'))
        pl_flagged = sum(1 for x in pl_data if x.get('flag'))
        
        self.bs_count.setText(f'BS flagged: {bs_flagged}')
        self.pl_count.setText(f'PL flagged: {pl_flagged}')
        
        self.generate_btn.setEnabled(True)
        self.placeholder.hide()
        
        # Populate table with flagged items
        self._populate_table()
    
    def _populate_table(self):
        """Fill table with all flagged items (before remarks are generated)."""
        flagged_items = []
        
        for i, row in enumerate(self.bs_data):
            if row.get('flag'):
                flagged_items.append(('BS', i, row))
        for i, row in enumerate(self.pl_data):
            if row.get('flag'):
                flagged_items.append(('PL', i, row))
        
        self.table.setRowCount(len(flagged_items))
        self._flagged_items = flagged_items
        
        for table_row, (stmt_type, idx, row) in enumerate(flagged_items):
            # Type
            type_item = QTableWidgetItem(stmt_type)
            type_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            type_item.setBackground(QBrush(QColor(BLUE if stmt_type == 'BS' else GREEN)))
            type_item.setForeground(QBrush(QColor('#FFFFFF')))
            type_item.setFont(QFont('Segoe UI', 10, QFont.Weight.Bold))
            type_item.setFlags(type_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(table_row, 0, type_item)
            
            # Line item
            name_item = QTableWidgetItem(row.get('particulars', ''))
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(table_row, 1, name_item)
            
            # Variance %
            pct_item = QTableWidgetItem(row.get('display_pct', ''))
            pct_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            pct_item.setBackground(QBrush(QColor('#EAB308')))
            pct_item.setForeground(QBrush(QColor('#000000')))
            pct_item.setFlags(pct_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(table_row, 2, pct_item)
            
            # Remark (empty initially, editable)
            remark = self.bs_remarks.get(idx, '') if stmt_type == 'BS' else self.pl_remarks.get(idx, '')
            self.table.setItem(table_row, 3, QTableWidgetItem(remark))
            
            # Status
            status = '✅' if remark else '⏳'
            status_item = QTableWidgetItem(status)
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(table_row, 4, status_item)
            
            # Action (placeholder for regen button)
            action_item = QTableWidgetItem('↺')
            action_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            action_item.setFlags(action_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(table_row, 5, action_item)
    
    def _start_generation(self):
        """Start AI remark generation in background."""
        from src.core.finance_agent import load_api_key
        api_key = load_api_key()
        if not api_key:
            QMessageBox.warning(self, 'API Key Required',
                              'API key is missing in the code.')
            return
        
        self.generate_btn.setEnabled(False)
        self.regen_all_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText('⏳ Generating remarks...')
        
        self.worker = RemarkWorker(api_key, self.bs_data, self.pl_data, self.form_data)
        self.worker.progress.connect(self._on_progress)
        self.worker.remarkReady.connect(self._on_remark_ready)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()
    
    def _on_progress(self, current, total):
        if total > 0:
            pct = int(current / total * 100)
            self.progress_bar.setValue(pct)
            self.status_label.setText(f'⏳ Generating remarks: {current} of {total}')
    
    def _on_remark_ready(self, idx, remark, stmt_type):
        """Handle a single remark being ready."""
        if stmt_type == 'BS':
            self.bs_remarks[idx] = remark
        else:
            self.pl_remarks[idx] = remark
        
        # Update table
        for table_row, (st, orig_idx, _) in enumerate(self._flagged_items):
            if st == stmt_type and orig_idx == idx:
                self.table.item(table_row, 3).setText(remark)
                self.table.item(table_row, 4).setText('✅')
                break
        
        total = len(self.bs_remarks) + len(self.pl_remarks)
        self.total_generated.setText(f'Generated: {total}')
    
    def _on_finished(self):
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.regen_all_btn.setEnabled(True)
        self.status_label.setText('✅ All remarks generated!')
        self.status_label.setObjectName('successLabel')
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        
        self.remarksGenerated.emit(self.bs_remarks, self.pl_remarks)
    
    def _on_error(self, error_msg):
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.status_label.setText(f'❌ Error: {error_msg}')
        self.status_label.setObjectName('errorLabel')
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)


BLUE = '#4A1A6B'
GREEN = '#22C55E'
