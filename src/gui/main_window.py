"""
FAR Automation Tool — Main Window
Central application window with sidebar navigation and stacked content area.
"""

import logging
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QStackedWidget,
    QLabel, QFrame, QStatusBar, QProgressBar, QMessageBox,
    QProgressDialog, QApplication, QPushButton
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QAction, QKeySequence, QFont

from src.gui.sidebar import Sidebar
from src.gui.input_form import InputForm
from src.gui.bs_view import BSView, PLView
from src.gui.notes_view import NotesView
from src.gui.ratio_view import RatioView
from src.gui.dashboard_view import DashboardView
from src.gui.remarks_view import RemarksView
from src.gui.export_view import ExportView

logger = logging.getLogger(__name__)


class TopBar(QFrame):
    """Top information bar showing client context."""
    exportRequested = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('topBarFrame')
        self.setFixedHeight(44)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)
        
        self.client_label = QLabel('Client: —')
        self.client_label.setStyleSheet(
            "color: #FFFFFF; font-size: 13px; font-weight: 600; background: transparent;")
        layout.addWidget(self.client_label)
        
        self.fy_label = QLabel('')
        self.fy_label.setStyleSheet(
            "color: #CCCCCC; font-size: 12px; background: transparent;")
        layout.addWidget(self.fy_label)
        
        self.unit_label = QLabel('')
        self.unit_label.setStyleSheet(
            "color: #888888; font-size: 12px; background: transparent;")
        layout.addWidget(self.unit_label)
        
        layout.addStretch()
        
        self.export_btn = QPushButton('📊 Generate Excel Sheet')
        self.export_btn.setObjectName('successButton')
        self.export_btn.setFixedHeight(28)
        self.export_btn.setStyleSheet("font-weight: bold; background-color: #22C55E; color: white; border-radius: 4px; padding: 4px 12px;")
        self.export_btn.clicked.connect(self.exportRequested.emit)
        self.export_btn.hide()
        layout.addWidget(self.export_btn)
        
        self.status_text = QLabel('Ready')
        self.status_text.setStyleSheet(
            "color: #888888; font-size: 11px; background: transparent; padding-left: 10px;")
        layout.addWidget(self.status_text)
    
    def update_info(self, client='', fy='', unit=''):
        self.client_label.setText(f'Client: {client}' if client else 'Client: —')
        self.fy_label.setText(f'  |  FY: {fy}' if fy else '')
        self.unit_label.setText(f'  |  {unit}' if unit else '')
        if client:
            self.export_btn.show()


class ToastNotification(QLabel):
    """Floating toast notification widget."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedHeight(40)
        self.setMinimumWidth(300)
        self.hide()
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)
    
    def show_toast(self, message, toast_type='info', duration=3000):
        colors = {
            'success': ('#22C55E', '#000000'),
            'error': ('#EF4444', '#FFFFFF'),
            'info': ('#007ACC', '#FFFFFF'),
            'warning': ('#EAB308', '#000000'),
        }
        bg, fg = colors.get(toast_type, colors['info'])
        
        self.setText(f'  {message}  ')
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg};
                color: {fg};
                font-size: 13px;
                font-weight: 600;
                border-radius: 6px;
                padding: 8px 20px;
            }}
        """)
        
        # Position at bottom-right of parent
        if self.parent():
            parent_rect = self.parent().rect()
            self.move(parent_rect.width() - self.width() - 20,
                     parent_rect.height() - 60)
        
        self.show()
        self.raise_()
        self._timer.start(duration)


class MainWindow(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle('FAR Automation Tool')
        self.setMinimumSize(1024, 768)
        self.resize(1280, 800)
        
        # Application state
        self.app_data = {}
        
        self._setup_ui()
        self._setup_menu()
        self._setup_shortcuts()
        self._setup_statusbar()
    
    def _setup_ui(self):
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Sidebar
        self.sidebar = Sidebar()
        self.sidebar.moduleChanged.connect(self._on_module_changed)
        main_layout.addWidget(self.sidebar)
        
        # Right panel
        right_panel = QVBoxLayout()
        right_panel.setContentsMargins(0, 0, 0, 0)
        right_panel.setSpacing(0)
        
        # Top bar
        self.top_bar = TopBar()
        self.top_bar.exportRequested.connect(self._direct_excel_export)
        right_panel.addWidget(self.top_bar)
        
        # Stacked content area
        self.stack = QStackedWidget()
        
        # Create all module views
        self.input_form = InputForm()
        self.input_form.generateRequested.connect(self._on_generate)
        
        self.bs_view = BSView()
        self.pl_view = PLView()
        self.notes_view = NotesView()
        self.ratio_view = RatioView()
        self.dashboard_view = DashboardView()
        self.remarks_view = RemarksView()
        self.remarks_view.remarksGenerated.connect(self._on_remarks_generated)
        self.export_view = ExportView()
        
        # Settings placeholder
        self.settings_view = QWidget()
        settings_layout = QVBoxLayout(self.settings_view)
        settings_layout.setContentsMargins(32, 24, 32, 24)
        settings_title = QLabel('⚙️ Settings')
        settings_title.setObjectName('headingLabel')
        settings_layout.addWidget(settings_title)
        settings_layout.addWidget(QLabel('Settings panel — coming soon.'))
        settings_layout.addStretch()
        
        # Add to stack
        self.stack.addWidget(self.input_form)       # 0
        self.stack.addWidget(self.bs_view)           # 1
        self.stack.addWidget(self.pl_view)           # 2
        self.stack.addWidget(self.notes_view)        # 3
        self.stack.addWidget(self.ratio_view)        # 4
        self.stack.addWidget(self.dashboard_view)    # 5
        self.stack.addWidget(self.remarks_view)      # 6
        self.stack.addWidget(self.export_view)       # 7
        self.stack.addWidget(self.settings_view)     # 8
        
        right_panel.addWidget(self.stack)
        main_layout.addLayout(right_panel)
        
        # Toast notification
        self.toast = ToastNotification(self)
    
    def _setup_menu(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('&File')
        
        generate_action = QAction('&Generate FAR', self)
        generate_action.setShortcut('Ctrl+G')
        generate_action.triggered.connect(lambda: self.sidebar.set_active_module(0))
        file_menu.addAction(generate_action)
        
        export_action = QAction('&Export', self)
        export_action.setShortcut('Ctrl+E')
        export_action.triggered.connect(lambda: self.sidebar.set_active_module(7))
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        quit_action = QAction('&Quit', self)
        quit_action.setShortcut('Ctrl+Q')
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)
        
        # View menu
        view_menu = menubar.addMenu('&View')
        
        for i, (_, name) in enumerate(Sidebar.MODULES):
            action = QAction(name, self)
            action.triggered.connect(lambda checked, idx=i: self.sidebar.set_active_module(idx))
            view_menu.addAction(action)
        
        # Help menu
        help_menu = menubar.addMenu('&Help')
        
        about_action = QAction('&About', self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)
    
    def _setup_shortcuts(self):
        pass  # Shortcuts set via menu actions
    
    def _setup_statusbar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        self.status_progress = QProgressBar()
        self.status_progress.setFixedWidth(200)
        self.status_progress.setFixedHeight(16)
        self.status_progress.setVisible(False)
        self.status_bar.addPermanentWidget(self.status_progress)
        
        self.status_items = QLabel('')
        self.status_bar.addPermanentWidget(self.status_items)
        
        version_label = QLabel('v1.0.0')
        version_label.setStyleSheet("color: rgba(255,255,255,0.7); font-size: 11px;")
        self.status_bar.addPermanentWidget(version_label)
        
        self.status_bar.showMessage('Ready')
    
    def _on_module_changed(self, index):
        """Handle sidebar module navigation."""
        if index == -1:
            # Settings
            self.stack.setCurrentIndex(8)
        elif 0 <= index <= 7:
            self.stack.setCurrentIndex(index)
    
    def _on_generate(self, form_data):
        """Handle Generate FAR request from input form."""
        self.app_data['form'] = form_data
        
        round_off = form_data.get('round_off', True)
        rounding_unit = form_data.get('rounding_unit', 'Lakhs') if round_off else 'Actuals'
        
        # Update top bar
        self.top_bar.update_info(
            form_data.get('client_name', ''),
            form_data.get('financial_year', ''),
            rounding_unit
        )
        
        # Show progress dialog
        progress = QProgressDialog('Generating FAR...', None, 0, 100, self)
        progress.setWindowTitle('Processing')
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumWidth(400)
        progress.setAutoClose(True)
        progress.setAutoReset(True)
        progress.show()
        QApplication.processEvents()
        
        try:
            # Step 1: Parse BS
            progress.setLabelText('📄 Parsing Balance Sheet...')
            progress.setValue(10)
            QApplication.processEvents()
            
            from src.core.excel_parser import (
                parse_balance_sheet, parse_profit_loss, parse_notes
            )
            from src.core.cpp_bridge import compute_variances, compute_ratios_from_raw
            from src.core.validation import validate_far_data
            
            bs_parsed = parse_balance_sheet(form_data['bs_file'])
            
            # Step 2: Parse PL
            progress.setLabelText('📄 Parsing Profit & Loss...')
            progress.setValue(25)
            QApplication.processEvents()
            
            pl_parsed = parse_profit_loss(form_data['pl_file'])
            
            if bs_parsed['cy_year'] != pl_parsed['cy_year']:
                progress.close()
                QMessageBox.warning(self, 'Year Mismatch',
                    f"BS year ({bs_parsed['cy_year']}) and P&L year ({pl_parsed['cy_year']}) don't match!\n"
                    "Please check your files.")
                return

            # Scale parsed values by the rounding unit factor
            round_off = form_data.get('round_off', True)
            rounding_unit = form_data.get('rounding_unit', 'Lakhs') if round_off else 'Actuals'
            factor = 1.0
            decimals = 2
            if round_off:
                if rounding_unit == 'Thousands':
                    factor = 1000.0
                    decimals = 0
                elif rounding_unit == 'Lakhs':
                    factor = 100000.0
                    decimals = 2
                elif rounding_unit == 'Millions':
                    factor = 1000000.0
                    decimals = 2
            else:
                factor = 1.0
                decimals = 2

            # Scale BS data
            for row in bs_parsed['data']:
                if row.get('cy') is not None:
                    row['cy'] = round(row['cy'] / factor, decimals)
                if row.get('py') is not None:
                    row['py'] = round(row['py'] / factor, decimals)

            # Scale PL data
            for row in pl_parsed['data']:
                if row.get('cy') is not None:
                    row['cy'] = round(row['cy'] / factor, decimals)
                if row.get('py') is not None:
                    row['py'] = round(row['py'] / factor, decimals)
            
            # Step 3: Parse Notes
            notes_result = {}
            if form_data.get('notes_required') and form_data.get('notes_files'):
                progress.setLabelText('📄 Parsing Notes to Accounts...')
                progress.setValue(35)
                QApplication.processEvents()
                
                for notes_file in form_data['notes_files']:
                    notes_data = parse_notes(
                        notes_file,
                        cy_year=bs_parsed.get('cy_year', 2025),
                        py_year=bs_parsed.get('py_year', 2024)
                    )
                    # Scale Notes data
                    for sheet_name, note_groups in notes_data.items():
                        for group in note_groups:
                            for row in group.get('data', []):
                                if row.get('cy') is not None:
                                    row['cy'] = round(row['cy'] / factor, decimals)
                                if row.get('py') is not None:
                                    row['py'] = round(row['py'] / factor, decimals)
                    notes_result.update(notes_data)
            
            # Step 4: Compute variances via C++ engine
            progress.setLabelText('⚡ Computing variances...')
            progress.setValue(50)
            QApplication.processEvents()
            
            bs_result = compute_variances(bs_parsed['data'])
            pl_result = compute_variances(pl_parsed['data'])
            
            # Step 5: Compute ratios via C++ (which also extracts summaries internally)
            progress.setLabelText('📐 Calculating financial ratios...')
            progress.setValue(65)
            QApplication.processEvents()
            
            raw_ratios_result = compute_ratios_from_raw(bs_parsed['data'], pl_parsed['data'])
            ratios = raw_ratios_result.get('ratios', [])
            bs_summary = raw_ratios_result.get('bs_summary', {})
            pl_summary = raw_ratios_result.get('pl_summary', {})
            
            # Step 5b: Run validation layer
            progress.setLabelText('🔍 Validating data...')
            progress.setValue(70)
            QApplication.processEvents()
            
            validation_report = validate_far_data(
                bs_parsed['data'], pl_parsed['data'],
                bs_summary, pl_summary, ratios
            )
            
            # Step 6: Store results
            progress.setLabelText('📊 Building views...')
            progress.setValue(80)
            QApplication.processEvents()
            
            cy_year = bs_parsed.get('cy_year', 2025)
            py_year = bs_parsed.get('py_year', 2024)
            
            # Use form input client name; fall back to auto-extracted name from workbook
            form_client_name = form_data.get('client_name', '').strip()
            parsed_client_name = bs_parsed.get('client_name', '').strip()
            client_name = form_client_name if form_client_name else parsed_client_name
            if not form_client_name and parsed_client_name:
                logger.info("Using auto-extracted client name: %s", parsed_client_name)
            
            rounding_unit = form_data.get('rounding_unit', 'Lakhs') if round_off else 'Actuals'
            
            self.app_data.update({
                'bs_result': bs_result,
                'pl_result': pl_result,
                'notes_result': notes_result,
                'ratios': ratios,
                'cy_year': cy_year,
                'py_year': py_year,
                'client_name': client_name,
                'firm_name': form_data.get('firm_name', ''),
                'financial_year': form_data.get('financial_year', ''),
                'rounding_unit': rounding_unit,
                'remarks_bs': {},
                'remarks_pl': {},
                'validation_report': validation_report,
            })
            
            # Load data into views
            self.bs_view.load_data(bs_result, client_name, cy_year, py_year, rounding_unit)
            self.pl_view.load_data(pl_result, client_name, cy_year, py_year, rounding_unit)
            self.notes_view.load_data(notes_result, client_name, cy_year, py_year, rounding_unit)
            self.ratio_view.load_data(ratios, client_name, cy_year, py_year)
            self.dashboard_view.load_data(bs_result, pl_result, ratios,
                                         client_name, cy_year, py_year, rounding_unit)
            self.remarks_view.set_data(bs_result, pl_result, form_data)
            self.export_view.set_export_data(self.app_data)
            
            progress.setLabelText('✅ Done!')
            progress.setValue(100)
            QApplication.processEvents()
            
            # Update status bar with validation summary
            val_summary = validation_report.summary_text()
            self.status_items.setText(
                f'BS: {len(bs_result)} items  |  PL: {len(pl_result)} items  |  {val_summary}')
            self.status_bar.showMessage('FAR generated successfully', 5000)
            
            # Switch to BS view
            self.sidebar.set_active_module(1)
            
            # Toast — show validation result
            if validation_report.has_failures:
                self.toast.show_toast(
                    f'⚠ FAR generated — {len(validation_report.failed)} validation issue(s) detected',
                    'warning', 6000)
            else:
                self.toast.show_toast('✓ FAR generated successfully!', 'success')
            
            logger.info("FAR generated: BS=%d rows, PL=%d rows, Ratios=%d",
                        len(bs_result), len(pl_result), len(ratios))
        
        except Exception as e:
            progress.close()
            logger.error("FAR generation failed: %s", e, exc_info=True)
            QMessageBox.critical(self, 'Generation Error',
                               f'Failed to generate FAR:\n\n{str(e)}')
            self.toast.show_toast(f'✗ Error: {str(e)[:60]}', 'error')
    
    def _on_remarks_generated(self, bs_remarks, pl_remarks):
        """Handle AI remarks being generated."""
        self.app_data['remarks_bs'] = bs_remarks
        self.app_data['remarks_pl'] = pl_remarks
        
        # Update BS/PL views with remarks
        for idx, remark in bs_remarks.items():
            self.bs_view.update_remark(idx, remark)
        for idx, remark in pl_remarks.items():
            self.pl_view.update_remark(idx, remark)
        
        # Update export data
        self.export_view.set_export_data(self.app_data)
        
        self.toast.show_toast(
            f'✓ AI remarks generated: {len(bs_remarks) + len(pl_remarks)} items',
            'success')
            
    def _direct_excel_export(self):
        """Directly trigger Excel export from top bar."""
        self.export_view._export_excel()
    
    def _show_about(self):
        QMessageBox.about(
            self, 'About FAR Automation Tool',
            'FAR Automation Tool v1.0.0\n\n'
            'Final Analytical Review Workpaper\n'
            'Automation Tool for Audit Firms.\n\n'
            'Powered by Python, PyQt6, C++,\n'
            'and Claude Sonnet AI.\n\n'
            '© 2025 All rights reserved.')
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Reposition toast
        if hasattr(self, 'toast'):
            self.toast.move(self.width() - self.toast.width() - 20,
                          self.height() - 80)
