"""
FAR Automation Tool — Export View (Module 8)
Export panel for generating Excel and Word output files.
"""

import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QFrame, QProgressBar, QCheckBox, QMessageBox,
    QGroupBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QFont, QDesktopServices


class ExportWorker(QThread):
    """Background thread for file export."""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str, str)  # success, message, filepath
    
    def __init__(self, export_func, *args, **kwargs):
        super().__init__()
        self.export_func = export_func
        self.args = args
        self.kwargs = kwargs
    
    def run(self):
        try:
            filepath = self.export_func(*self.args, **self.kwargs)
            self.finished.emit(True, "Export successful!", filepath)
        except Exception as e:
            self.finished.emit(False, str(e), "")


class ExportView(QWidget):
    """Module 8 — Export panel for Excel and Word documents."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.export_data = None
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 24)
        layout.setSpacing(20)
        
        # Title
        title = QLabel('💾 Export FAR Workpaper')
        title.setObjectName('headingLabel')
        layout.addWidget(title)
        
        subtitle = QLabel('Generate the final audit workpaper files')
        subtitle.setObjectName('mutedLabel')
        layout.addWidget(subtitle)
        
        # ── Excel Export Card ──
        excel_group = QGroupBox('Excel Workbook Export')
        excel_layout = QVBoxLayout(excel_group)
        excel_layout.setSpacing(12)
        excel_layout.setContentsMargins(16, 20, 16, 16)
        
        excel_desc = QLabel(
            'Generates a fully-linked Excel workbook with:\n'
            '• Balance Sheet analysis with variance formulas\n'
            '• Profit & Loss analysis with variance formulas\n'
            '• Notes to Accounts analysis\n'
            '• Ratio Analysis with cross-sheet references\n'
            '• Dashboard summary with key metrics\n'
            '• Cover page with client information'
        )
        excel_desc.setObjectName('mutedLabel')
        excel_desc.setWordWrap(True)
        excel_layout.addWidget(excel_desc)
        
        # Sheets checklist
        self.include_notes = QCheckBox('Include Notes to Accounts sheets')
        self.include_notes.setChecked(True)
        excel_layout.addWidget(self.include_notes)
        
        self.include_dashboard = QCheckBox('Include Dashboard sheet')
        self.include_dashboard.setChecked(True)
        excel_layout.addWidget(self.include_dashboard)
        
        self.include_remarks = QCheckBox('Include AI Remarks in output')
        self.include_remarks.setChecked(True)
        excel_layout.addWidget(self.include_remarks)
        
        excel_btn_row = QHBoxLayout()
        self.excel_btn = QPushButton('📊 Export as Excel (.xlsx)')
        self.excel_btn.setFixedHeight(44)
        self.excel_btn.setMinimumWidth(250)
        self.excel_btn.setFont(QFont('Segoe UI', 13, QFont.Weight.Bold))
        self.excel_btn.clicked.connect(self._export_excel)
        excel_btn_row.addWidget(self.excel_btn)
        excel_btn_row.addStretch()
        excel_layout.addLayout(excel_btn_row)
        
        layout.addWidget(excel_group)
        
        # ── Word Export Card ──
        docx_group = QGroupBox('Word Document Export')
        docx_layout = QVBoxLayout(docx_group)
        docx_layout.setSpacing(12)
        docx_layout.setContentsMargins(16, 20, 16, 16)
        
        docx_desc = QLabel(
            'Generates a professional Word document with:\n'
            '• Cover page with firm branding\n'
            '• BS and PL analysis tables\n'
            '• Ratio analysis summary\n'
            '• AI-generated remarks compilation'
        )
        docx_desc.setObjectName('mutedLabel')
        docx_desc.setWordWrap(True)
        docx_layout.addWidget(docx_desc)
        
        docx_btn_row = QHBoxLayout()
        self.docx_btn = QPushButton('📄 Export as Word (.docx)')
        self.docx_btn.setFixedHeight(44)
        self.docx_btn.setMinimumWidth(250)
        self.docx_btn.setFont(QFont('Segoe UI', 13, QFont.Weight.Bold))
        self.docx_btn.setObjectName('secondaryButton')
        self.docx_btn.clicked.connect(self._export_docx)
        docx_btn_row.addWidget(self.docx_btn)
        docx_btn_row.addStretch()
        docx_layout.addLayout(docx_btn_row)
        
        layout.addWidget(docx_group)
        
        # ── Progress & Status ──
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(20)
        layout.addWidget(self.progress_bar)
        
        self.status_frame = QFrame()
        self.status_frame.setObjectName('cardFrame')
        self.status_frame.setVisible(False)
        status_layout = QHBoxLayout(self.status_frame)
        status_layout.setContentsMargins(16, 12, 16, 12)
        
        self.status_icon = QLabel('✅')
        self.status_icon.setObjectName('exportStatusIcon')
        status_layout.addWidget(self.status_icon)
        
        self.status_text = QLabel('')
        self.status_text.setObjectName('exportStatusText')
        status_layout.addWidget(self.status_text)
        
        status_layout.addStretch()
        
        self.open_file_btn = QPushButton('📂 Open File')
        self.open_file_btn.setObjectName('successButton')
        self.open_file_btn.setVisible(False)
        self.open_file_btn.clicked.connect(self._open_last_file)
        status_layout.addWidget(self.open_file_btn)
        
        self.open_folder_btn = QPushButton('📁 Open Folder')
        self.open_folder_btn.setObjectName('secondaryButton')
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_last_folder)
        status_layout.addWidget(self.open_folder_btn)
        
        layout.addWidget(self.status_frame)
        
        layout.addStretch()
        
        self.last_file = ''
    
    def set_export_data(self, data):
        """Set the data needed for export."""
        self.export_data = data
        self.excel_btn.setEnabled(True)
        self.docx_btn.setEnabled(True)
    
    def _get_default_filename(self, ext):
        if self.export_data:
            client = self.export_data.get('client_name', 'Client').replace(' ', '_')
            fy = self.export_data.get('financial_year', 'FY')
            prefix = self.export_data.get('report_type', 'FAR')
            return f"{client}_{prefix}_{fy}.{ext}"
        prefix = self.export_data.get('report_type', 'FAR') if self.export_data else 'FAR'
        return f"{prefix}_Output.{ext}"
    
    def _generate_temp_charts(self):
        import tempfile
        import shutil
        from src.core.chart_generator import generate_all_charts
        
        temp_dir = tempfile.mkdtemp(prefix='far_charts_')
        d = self.export_data
        try:
            chart_paths = generate_all_charts(
                d.get('bs_result', []),
                d.get('pl_result', []),
                d.get('ratios', []),
                d.get('cy_year', 2025),
                d.get('py_year', 2024),
                d.get('rounding_unit', 'Lakhs'),
                temp_dir
            )
            return temp_dir, chart_paths
        except Exception as e:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise e

    def _export_excel(self):
        if not self.export_data:
            QMessageBox.warning(self, 'No Data', 'Generate FAR from Input Form first.')
            return
        
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        default_name = self._get_default_filename('xlsx')
        filepath, _ = QFileDialog.getSaveFileName(
            self, 'Save Excel File',
            os.path.join(desktop, default_name),
            'Excel Files (*.xlsx)')
        
        if not filepath:
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(0)  # Indeterminate
        self.excel_btn.setEnabled(False)
        
        temp_dir = None
        chart_paths = None
        try:
            if self.include_dashboard.isChecked():
                temp_dir, chart_paths = self._generate_temp_charts()

            from src.core.excel_export import export_far_workbook
            
            d = self.export_data
            export_far_workbook(
                filepath,
                d.get('bs_result', []),
                d.get('pl_result', []),
                d.get('notes_result', {}),
                d.get('ratios', []),
                d.get('remarks_bs', {}),
                d.get('remarks_pl', {}),
                d.get('client_name', ''),
                d.get('firm_name', ''),
                d.get('financial_year', ''),
                d.get('rounding_unit', 'Lakhs'),
                d.get('cy_year', 2025),
                d.get('py_year', 2024),
                chart_paths=chart_paths,
                bs_notes=d.get('bs_notes', []),
                bs_signatures=d.get('bs_signatures', []),
                bs_footers=d.get('bs_footers', []),
                pl_notes=d.get('pl_notes', []),
                pl_signatures=d.get('pl_signatures', []),
                pl_footers=d.get('pl_footers', []),
                notes_sheet_notes=d.get('notes_sheet_notes', {}),
                notes_signatures=d.get('notes_signatures', {}),
                notes_footers=d.get('notes_footers', {}),
                report_type=d.get('report_type', 'FAR'),
                meta=d.get('meta', {}),
            )
            
            self.last_file = filepath
            self._show_success(f'Excel exported: {os.path.basename(filepath)}')
        except Exception as e:
            self._show_error(str(e))
        finally:
            if temp_dir:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.progress_bar.setVisible(False)
            self.excel_btn.setEnabled(True)
    
    def _export_docx(self):
        if not self.export_data:
            QMessageBox.warning(self, 'No Data', 'Generate FAR from Input Form first.')
            return
        
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        default_name = self._get_default_filename('docx')
        filepath, _ = QFileDialog.getSaveFileName(
            self, 'Save Word File',
            os.path.join(desktop, default_name),
            'Word Documents (*.docx)')
        
        if not filepath:
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(0)
        self.docx_btn.setEnabled(False)
        
        temp_dir = None
        chart_paths = None
        try:
            temp_dir, chart_paths = self._generate_temp_charts()

            from src.core.docx_export import export_docx
            
            d = self.export_data
            export_docx(
                filepath,
                d.get('bs_result', []),
                d.get('pl_result', []),
                d.get('ratios', []),
                d.get('remarks_bs', {}),
                d.get('remarks_pl', {}),
                d.get('client_name', ''),
                d.get('firm_name', ''),
                d.get('financial_year', ''),
                d.get('rounding_unit', 'Lakhs'),
                d.get('cy_year', 2025),
                d.get('py_year', 2024),
                chart_paths=chart_paths,
                bs_notes=d.get('bs_notes', []),
                bs_signatures=d.get('bs_signatures', []),
                bs_footers=d.get('bs_footers', []),
                pl_notes=d.get('pl_notes', []),
                pl_signatures=d.get('pl_signatures', []),
                pl_footers=d.get('pl_footers', []),
                notes_sheet_notes=d.get('notes_sheet_notes', {}),
                notes_signatures=d.get('notes_signatures', {}),
                notes_footers=d.get('notes_footers', {}),
                report_type=d.get('report_type', 'FAR'),
                meta=d.get('meta', {}),
            )
            
            self.last_file = filepath
            self._show_success(f'Word document exported: {os.path.basename(filepath)}')
        except Exception as e:
            self._show_error(str(e))
        finally:
            if temp_dir:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.progress_bar.setVisible(False)
            self.docx_btn.setEnabled(True)
    
    def _show_success(self, message):
        self.status_frame.setVisible(True)
        self.status_icon.setText('✅')
        self.status_text.setText(message)
        self.status_text.setObjectName('successLabel')
        self.status_text.style().unpolish(self.status_text)
        self.status_text.style().polish(self.status_text)
        self.open_file_btn.setVisible(True)
        self.open_folder_btn.setVisible(True)
    
    def _show_error(self, message):
        self.status_frame.setVisible(True)
        self.status_icon.setText('❌')
        self.status_text.setText(f'Error: {message}')
        self.status_text.setObjectName('errorLabel')
        self.status_text.style().unpolish(self.status_text)
        self.status_text.style().polish(self.status_text)
        self.open_file_btn.setVisible(False)
        self.open_folder_btn.setVisible(False)
    
    def _open_last_file(self):
        if self.last_file and os.path.exists(self.last_file):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.last_file))
    
    def _open_last_folder(self):
        if self.last_file:
            folder = os.path.dirname(self.last_file)
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
