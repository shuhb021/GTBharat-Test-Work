"""
FAR Automation Tool — Input Form (Module 1)
Client information form with file upload panels and Generate button.
"""

import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLineEdit, QComboBox, QRadioButton, QCheckBox, QPushButton,
    QLabel, QFileDialog, QProgressDialog, QButtonGroup, QFrame,
    QScrollArea, QSizePolicy, QMessageBox
)
from PyQt6.QtCore import pyqtSignal, Qt, QThread, pyqtSlot
from PyQt6.QtGui import QFont


class GenerateWorker(QThread):
    """Background thread for FAR generation."""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
    
    def run(self):
        try:
            self.callback(self.progress.emit)
            self.finished.emit(True, "FAR generated successfully!")
        except Exception as e:
            self.finished.emit(False, str(e))


class FileUploadPanel(QGroupBox):
    """A file upload panel with Browse button and file info display."""
    
    fileSelected = pyqtSignal(str)
    
    def __init__(self, title, multi_select=False, parent=None):
        super().__init__(title, parent)
        self.multi_select = multi_select
        self.file_paths = []
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        
        # Browse button row
        btn_row = QHBoxLayout()
        self.browse_btn = QPushButton('📂 Browse...')
        self.browse_btn.setFixedWidth(140)
        self.browse_btn.clicked.connect(self._browse)
        btn_row.addWidget(self.browse_btn)
        
        self.status_icon = QLabel('⬜')
        self.status_icon.setObjectName('statusIcon')
        btn_row.addWidget(self.status_icon)
        
        btn_row.addStretch()
        layout.addLayout(btn_row)
        
        # File path display
        self.path_label = QLabel('No file selected')
        self.path_label.setObjectName('mutedLabel')
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)
        
        # Sheet info
        self.sheet_label = QLabel('')
        self.sheet_label.setObjectName('mutedLabel')
        self.sheet_label.setWordWrap(True)
        layout.addWidget(self.sheet_label)
    
    def _browse(self):
        if self.multi_select:
            files, _ = QFileDialog.getOpenFileNames(
                self, 'Select Excel Files', '',
                'Excel Files (*.xlsx *.xls)')
            if files:
                self.file_paths = files
                self._update_display()
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self, 'Select Excel File', '',
                'Excel Files (*.xlsx *.xls)')
            if file_path:
                self.file_paths = [file_path]
                self._update_display()
    
    def _update_display(self):
        if self.file_paths:
            names = [os.path.basename(f) for f in self.file_paths]
            self.path_label.setText(', '.join(names))
            self.path_label.setObjectName('successLabel')
            self.status_icon.setText('✅')
            
            # Try to show sheet names
            try:
                from openpyxl import load_workbook
                sheets_info = []
                for fp in self.file_paths:
                    wb = load_workbook(fp, read_only=True)
                    sheets_info.append(f"{os.path.basename(fp)}: {', '.join(wb.sheetnames)}")
                    wb.close()
                self.sheet_label.setText('Sheets: ' + ' | '.join(sheets_info))
            except Exception:
                self.sheet_label.setText('')
            
            self.fileSelected.emit(self.file_paths[0])
    
    def get_paths(self):
        return self.file_paths
    
    def is_loaded(self):
        return len(self.file_paths) > 0


class InputForm(QWidget):
    """Module 1 — Input Form for client information and file uploads."""
    
    generateRequested = pyqtSignal(dict)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
    
    def _setup_ui(self):
        # Scroll area wrapper
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        content = QWidget()
        main_layout = QVBoxLayout(content)
        main_layout.setContentsMargins(32, 24, 32, 24)
        main_layout.setSpacing(20)
        
        # Title
        title = QLabel('📝 Input Form')
        title.setObjectName('headingLabel')
        main_layout.addWidget(title)
        
        subtitle = QLabel('Enter client details and upload financial statements')
        subtitle.setObjectName('mutedLabel')
        main_layout.addWidget(subtitle)
        
        # ── Client Information ──
        info_group = QGroupBox('Client Information')
        info_layout = QFormLayout(info_group)
        info_layout.setSpacing(12)
        info_layout.setContentsMargins(16, 20, 16, 16)
        
        self.firm_name = QLineEdit()
        self.firm_name.setPlaceholderText('Enter firm name...')
        info_layout.addRow('Firm Name:', self.firm_name)
        
        self.client_name = QLineEdit()
        self.client_name.setPlaceholderText('Enter client name...')
        info_layout.addRow('Client Name:', self.client_name)
        
        self.financial_year = QComboBox()
        years = [f'{y}-{str(y+1)[-2:]}' for y in range(2025, 2019, -1)]
        self.financial_year.addItems(years)
        info_layout.addRow('Financial Year:', self.financial_year)
        
        # Fetch details button
        self.fetch_details_btn = QPushButton("🔍 Auto-Fetch Details from BS File")
        self.fetch_details_btn.setObjectName('fetchDetailsBtn')
        self.fetch_details_btn.clicked.connect(self._fetch_details_from_file)
        info_layout.addRow('', self.fetch_details_btn)
        
        # Round Off
        self.round_off = QCheckBox("Enable Rounding")
        self.round_off.setChecked(True)
        info_layout.addRow('Round Off:', self.round_off)
        
        # Rounding unit
        self.unit_widget = QWidget()
        unit_layout = QHBoxLayout(self.unit_widget)
        unit_layout.setContentsMargins(0, 0, 0, 0)
        self.unit_group = QButtonGroup()
        
        units = ['Thousands', 'Lakhs', 'Millions']
        for i, unit in enumerate(units):
            rb = QRadioButton(unit)
            self.unit_group.addButton(rb, i)
            unit_layout.addWidget(rb)
            if unit == 'Lakhs':
                rb.setChecked(True)
        
        unit_layout.addStretch()
        info_layout.addRow('Rounding Unit:', self.unit_widget)
        
        # Connect checkbox to enable/disable units
        self.round_off.toggled.connect(self.unit_widget.setEnabled)
        
        # Notes required
        self.notes_required = QCheckBox('Include Notes to Accounts analysis')
        self.notes_required.setChecked(True)
        info_layout.addRow('Notes Required:', self.notes_required)
        
        main_layout.addWidget(info_group)
        

        # ── File Upload ──
        files_group = QGroupBox('Financial Statement')
        files_layout = QVBoxLayout(files_group)
        files_layout.setSpacing(12)
        files_layout.setContentsMargins(16, 20, 16, 16)
        
        self.fs_upload = FileUploadPanel('Financial Statement File (BS, P&L, Notes)')
        self.fs_upload.fileSelected.connect(self._auto_fill_client_name)
        files_layout.addWidget(self.fs_upload)
        
        main_layout.addWidget(files_group)
        
        # ── Generate Button ──
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        
        self.generate_btn = QPushButton('🚀  Generate FAR Workpaper')
        self.generate_btn.setObjectName('generateFarBtn')
        self.generate_btn.setFixedHeight(48)
        self.generate_btn.setMinimumWidth(280)
        self.generate_btn.setFont(QFont('Segoe UI', 14, QFont.Weight.Bold))
        self.generate_btn.clicked.connect(self._on_generate)
        btn_row.addWidget(self.generate_btn)
        
        btn_row.addStretch()
        main_layout.addLayout(btn_row)
        
        # Validation error label
        self.error_label = QLabel('')
        self.error_label.setObjectName('errorLabel')
        self.error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.error_label)
        
        main_layout.addStretch()
        
        scroll.setWidget(content)
        
        # Wrap scroll in main layout
        wrapper = QVBoxLayout(self)
        wrapper.setContentsMargins(0, 0, 0, 0)
        wrapper.addWidget(scroll)
        


    def _fetch_details_from_file(self):
        """Fetch client name and financial year from the selected Balance Sheet file."""
        if not self.fs_upload.is_loaded():
            QMessageBox.warning(self, 'No File Selected', 'Please select a Financial Statement file first.')
            return
        
        filepath = self.fs_upload.get_paths()[0]
        try:
            from openpyxl import load_workbook
            from src.core.excel_parser import _extract_client_name, _find_date_columns
            
            wb = load_workbook(filepath, data_only=True, read_only=True)
            
            # Find BS sheet
            ws = None
            for name in ['BS', 'Balance Sheet', 'BalanceSheet']:
                if name in wb.sheetnames:
                    ws = wb[name]
                    break
            if ws is None:
                ws = wb.worksheets[0]
                
            client_name = _extract_client_name(ws)
            _, _, cy_year, _, _ = _find_date_columns(ws)
            wb.close()
            
            extracted_info = []
            if client_name:
                self.client_name.setText(client_name)
                extracted_info.append(f"Client Name: {client_name}")
            
            if cy_year:
                fy_text = f"{cy_year - 1}-{str(cy_year)[-2:]}"
                index = self.financial_year.findText(fy_text)
                if index >= 0:
                    self.financial_year.setCurrentIndex(index)
                else:
                    self.financial_year.addItem(fy_text)
                    self.financial_year.setCurrentText(fy_text)
                extracted_info.append(f"Financial Year: {fy_text}")
                
            if extracted_info:
                QMessageBox.information(
                    self, 'Details Fetched',
                    "Successfully fetched details from Balance Sheet:\n\n" + "\n".join(extracted_info)
                )
            else:
                QMessageBox.warning(
                    self, 'Details Not Found',
                    "Could not extract Client Name or Financial Year from the selected Balance Sheet file."
                )
        except Exception as e:
            QMessageBox.critical(
                self, 'Error Fetching Details',
                f"Failed to parse details from the Balance Sheet file:\n\n{str(e)}"
            )

    def _validate(self):
        """Validate all form fields."""
        errors = []
        
        if not self.client_name.text().strip():
            errors.append('Client Name is required')
        if not self.fs_upload.is_loaded():
            errors.append('Financial Statement file is required')
        
        if errors:
            self.error_label.setText('⚠️ ' + ' | '.join(errors))
            return False
        
        self.error_label.setText('')
        return True
    
    def _auto_fill_client_name(self, filepath):
        """
        Automatically populate the Client Name field from the selected BS workbook.
        Scans the first 5 rows of columns A-C to find the company name.
        Only fills if the field is currently empty (does not overwrite user input).
        """
        if self.client_name.text().strip():
            return  # Don't overwrite if user already typed something
        try:
            from openpyxl import load_workbook as _load_wb
            wb = _load_wb(filepath, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            # Try to find 'BS' sheet
            for name in ['BS', 'Balance Sheet', 'BalanceSheet']:
                if name in wb.sheetnames:
                    ws = wb[name]
                    break
            extracted = ''
            for row_idx in range(1, 6):
                for col_idx in range(1, 4):
                    cell_val = ws.cell(row=row_idx, column=col_idx).value
                    if cell_val and isinstance(cell_val, str) and len(cell_val.strip()) > 3:
                        extracted = cell_val.strip()
                        break
                if extracted:
                    break
            wb.close()
            if extracted:
                self.client_name.setText(extracted)
                import logging
                logging.getLogger(__name__).info(
                    "Auto-filled client name from workbook: %s", extracted)
        except Exception as e:
            import logging
            logging.getLogger(__name__).warning(
                "Could not auto-fill client name from workbook: %s", e)

    def _on_generate(self):
        """Handle Generate button click."""
        if not self._validate():
            return
        
        unit_btn = self.unit_group.checkedButton()
        unit_text = unit_btn.text() if unit_btn else 'Lakhs'
        
        fs_path = self.fs_upload.get_paths()[0] if self.fs_upload.is_loaded() else ''
        form_data = {
            'firm_name': self.firm_name.text().strip() or 'Audit Firm',
            'client_name': self.client_name.text().strip(),
            'financial_year': self.financial_year.currentText(),
            'round_off': self.round_off.isChecked(),
            'rounding_unit': unit_text,
            'notes_required': self.notes_required.isChecked(),
            'bs_file': fs_path,
            'pl_file': fs_path,
            'notes_files': [fs_path] if fs_path else [],
        }
        
        self.generateRequested.emit(form_data)
    
    def get_form_data(self):
        """Get current form data without triggering generation."""
        unit_btn = self.unit_group.checkedButton()
        unit_text = unit_btn.text() if unit_btn else 'Lakhs'
        
        fs_path = self.fs_upload.get_paths()[0] if self.fs_upload.is_loaded() else ''
        return {
            'firm_name': self.firm_name.text().strip() or 'Audit Firm',
            'client_name': self.client_name.text().strip(),
            'financial_year': self.financial_year.currentText(),
            'round_off': self.round_off.isChecked(),
            'rounding_unit': unit_text,
            'notes_required': self.notes_required.isChecked(),
            'bs_file': fs_path,
            'pl_file': fs_path,
            'notes_files': [fs_path] if fs_path else [],
        }
