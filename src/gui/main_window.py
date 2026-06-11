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
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtGui import QAction, QKeySequence, QFont

from src.gui.sidebar import Sidebar
from src.gui.input_form import InputForm
from src.gui.bs_view import BSView, PLView
from src.gui.notes_view import NotesView
from src.gui.ratio_view import RatioView
from src.gui.dashboard_view import DashboardView
from src.gui.remarks_view import RemarksView
from src.gui.export_view import ExportView
from src.gui.settings_view import SettingsView
from src.gui.comparison_view import ComparisonView

logger = logging.getLogger(__name__)

class GenerateWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, form_data):
        super().__init__()
        self.form_data = form_data
        
    def run(self):
        try:
            from src.core.data_engine import DataEngine
            from src.core.validation import validate_far_data
            import traceback
            
            engine = DataEngine()
            
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)
                
            # Offload to DataEngine which handles Caching, Polars, and Multiprocessing
            try:
                payload = engine.process_far_data(self.form_data, progress_callback=on_progress)
            except Exception as e:
                logger.error("DataEngine.process_far_data failed with full traceback:")
                traceback.print_exc()
                logger.error("FULL TRACEBACK:\n%s", traceback.format_exc())
                raise
            
            if payload is None:
                raise ValueError("DataEngine.process_far_data returned None")
            
            try:
                bs_result = payload['bs_data']
                pl_result = payload['pl_data']
                notes_result = payload['notes_data']
                ratios = payload['ratios']
                bs_summary = payload['bs_summary']
                pl_summary = payload['pl_summary']
                meta = payload['meta']
            except Exception as e:
                logger.error("Failed to unpack payload keys: %s", e, exc_info=True)
                logger.error("Payload type: %s", type(payload))
                logger.error("Payload keys: %s", list(payload.keys()) if isinstance(payload, dict) else "Not a dict")
                raise
            
            self.progress.emit(85, '🔍 Validating data...')
            validation_report = validate_far_data(
                bs_result, pl_result,
                bs_summary, pl_summary, ratios
            )
            
            self.progress.emit(95, '📊 Building views...')
            
            result_dict = {
                'bs_result': bs_result,
                'pl_result': pl_result,
                'notes_result': notes_result,
                'bs_notes': payload.get('bs_notes', []),
                'bs_signatures': payload.get('bs_signatures', []),
                'bs_footers': payload.get('bs_footers', []),
                'pl_notes': payload.get('pl_notes', []),
                'pl_signatures': payload.get('pl_signatures', []),
                'pl_footers': payload.get('pl_footers', []),
                'notes_sheet_notes': payload.get('notes_sheet_notes', {}),
                'notes_signatures': payload.get('notes_signatures', {}),
                'notes_footers': payload.get('notes_footers', {}),
                'ratios': ratios,
                'cy_year': meta.get('cy_year', 2025),
                'py_year': meta.get('py_year', 2024),
                'parsed_client_name': meta.get('client_name', '').strip(),
                'validation_report': validation_report
            }
            self.finished.emit(result_dict)
        except Exception as e:
            import traceback
            logger.error("GenerateWorker exception: %s", e, exc_info=True)
            logger.error("FULL TRACEBACK:\n%s", traceback.format_exc())
            traceback.print_exc()
            self.error.emit(str(e))



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
        self.client_label.setObjectName('topBarClient')
        layout.addWidget(self.client_label)
        
        self.fy_label = QLabel('')
        self.fy_label.setObjectName('topBarFY')
        layout.addWidget(self.fy_label)
        
        self.unit_label = QLabel('')
        self.unit_label.setObjectName('topBarUnit')
        layout.addWidget(self.unit_label)
        
        layout.addStretch()
        
        self.export_btn = QPushButton('📊 Generate Excel Sheet')
        self.export_btn.setObjectName('topBarExportBtn')
        self.export_btn.setFixedHeight(28)
        self.export_btn.clicked.connect(self.exportRequested.emit)
        self.export_btn.hide()
        layout.addWidget(self.export_btn)
        
        # Purple status dot + "Ready" text
        self.status_dot = QLabel('●')
        self.status_dot.setObjectName('statusDot')
        layout.addWidget(self.status_dot)
        
        self.status_text = QLabel('Ready')
        self.status_text.setObjectName('statusText')
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
            'info': ('#4A1A6B', '#FFFFFF'),
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
        self.app_config = self._load_config()
        
        self._setup_ui()
        self._setup_menu()
        self._setup_shortcuts()
        self._setup_statusbar()
        
        # Load config into settings view
        self.settings_view.load_config(self.app_config)
        
    def _load_config(self):
        import json
        import os
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        config_path = os.path.join(project_root, 'config.json')
        defaults = {
            'font_family': 'Segoe UI',
            'font_size': 11,
            'ui_zoom': 100
        }
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    return {**defaults, **json.load(f)}
            except Exception:
                pass
        return defaults

    def _save_config(self):
        import json
        import os
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        config_path = os.path.join(project_root, 'config.json')
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.app_config, f, indent=4)
        except Exception as e:
            logger.warning("Failed to save config.json: %s", e)

    def apply_global_settings(self, family, size, zoom):
        """Dynamically scales and applies stylesheet and base font settings."""
        self.app_config = {
            'font_family': family,
            'font_size': size,
            'ui_zoom': zoom
        }
        self._save_config()
        
        # Load and scale stylesheet
        import os
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        qss_path = os.path.join(project_root, 'src', 'assets', 'styles.qss')
        if os.path.exists(qss_path):
            try:
                with open(qss_path, 'r', encoding='utf-8') as f:
                    stylesheet = f.read()
                
                import re
                zoom_factor = zoom / 100.0
                size_factor = size / 11.0
                total_factor = zoom_factor * size_factor
                
                # Replace font family
                stylesheet = re.sub(r"font-family:\s*[^;]+;", f"font-family: '{family}', sans-serif;", stylesheet)
                
                # Scale font size
                def scale_size(match):
                    val = int(match.group(1))
                    new_val = max(8, int(round(val * total_factor)))
                    return f"font-size: {new_val}px"
                stylesheet = re.sub(r"font-size:\s*(\d+)\s*px", scale_size, stylesheet)
                
                QApplication.instance().setStyleSheet(stylesheet)
            except Exception as e:
                logger.warning("Failed to scale QSS: %s", e)
            
        # Set application base font size
        scaled_size = int(round(size * (zoom / 100.0)))
        font = QFont(family, scaled_size)
        QApplication.instance().setFont(font)
    
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
        self.comparison_view = ComparisonView()
        self.remarks_view = RemarksView()
        self.remarks_view.remarksGenerated.connect(self._on_remarks_generated)
        self.export_view = ExportView()
        
        # Create Settings view
        self.settings_view = SettingsView()
        self.settings_view.settingsChanged.connect(self.apply_global_settings)
        
        # Add to stack
        self.stack.addWidget(self.input_form)       # 0
        self.stack.addWidget(self.bs_view)           # 1
        self.stack.addWidget(self.pl_view)           # 2
        self.stack.addWidget(self.notes_view)        # 3
        self.stack.addWidget(self.ratio_view)        # 4
        self.stack.addWidget(self.dashboard_view)    # 5
        self.stack.addWidget(self.comparison_view)   # 6
        self.stack.addWidget(self.remarks_view)      # 7
        self.stack.addWidget(self.export_view)       # 8
        self.stack.addWidget(self.settings_view)     # 9
        
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
        export_action.triggered.connect(lambda: self.sidebar.set_active_module(8))
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
        self.status_bar.addPermanentWidget(version_label)
        
        self.status_bar.showMessage('Ready')
    
    def _on_module_changed(self, index):
        """Handle sidebar module navigation."""
        if index == -1:
            # Settings
            self.stack.setCurrentIndex(9)
        elif 0 <= index <= 8:
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
        report_type = form_data.get('report_type', 'FAR')
        self.progress_dialog = QProgressDialog(f'Generating {report_type}...', None, 0, 100, self)
        self.progress_dialog.setWindowTitle('Processing')
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumWidth(400)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.show()
        
        self.generate_worker = GenerateWorker(form_data)
        self.generate_worker.progress.connect(self._on_generate_progress)
        self.generate_worker.finished.connect(self._on_generate_finished)
        self.generate_worker.error.connect(self._on_generate_error)
        self.generate_worker.start()
        
    def _on_generate_progress(self, val, text):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.setValue(val)
            self.progress_dialog.setLabelText(text)
            
    def _on_generate_error(self, err_msg):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.close()
        logger.error("FAR generation failed: %s", err_msg)
        QMessageBox.critical(self, 'Generation Error', f'Failed to generate FAR:\n\n{err_msg}')
        self.toast.show_toast(f'✗ Error: {err_msg[:60]}', 'error')
        
    def _on_generate_finished(self, result):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.setLabelText('✅ Done!')
            self.progress_dialog.setValue(100)
            self.progress_dialog.close()
            
        bs_result = result['bs_result']
        pl_result = result['pl_result']
        notes_result = result['notes_result']
        ratios = result['ratios']
        cy_year = result['cy_year']
        py_year = result['py_year']
        parsed_client_name = result['parsed_client_name']
        validation_report = result['validation_report']
        
        form_data = self.app_data['form']
        form_client_name = form_data.get('client_name', '').strip()
        client_name = form_client_name if form_client_name else parsed_client_name
        if not form_client_name and parsed_client_name:
            logger.info("Using auto-extracted client name: %s", parsed_client_name)
            
        round_off = form_data.get('round_off', True)
        rounding_unit = form_data.get('rounding_unit', 'Lakhs') if round_off else 'Actuals'
        
        self.app_data.update({
            'bs_result': bs_result,
            'pl_result': pl_result,
            'notes_result': notes_result,
            'bs_notes': result.get('bs_notes', []),
            'bs_signatures': result.get('bs_signatures', []),
            'bs_footers': result.get('bs_footers', []),
            'pl_notes': result.get('pl_notes', []),
            'pl_signatures': result.get('pl_signatures', []),
            'pl_footers': result.get('pl_footers', []),
            'notes_sheet_notes': result.get('notes_sheet_notes', {}),
            'notes_signatures': result.get('notes_signatures', {}),
            'notes_footers': result.get('notes_footers', {}),
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
            'report_type': form_data.get('report_type', 'FAR'),
        })
        
        # Load data into views with fault isolation
        try:
            self.bs_view.load_data(
                bs_result, client_name, cy_year, py_year, rounding_unit,
                notes=result.get('bs_notes', []),
                signatures=result.get('bs_signatures', []),
                footers=result.get('bs_footers', [])
            )
            self.pl_view.load_data(
                pl_result, client_name, cy_year, py_year, rounding_unit,
                notes=result.get('pl_notes', []),
                signatures=result.get('pl_signatures', []),
                footers=result.get('pl_footers', [])
            )
        except Exception as e:
            logger.error(f"Failed to load BS/PL view: {e}")
            
        try:
            self.notes_view.load_data(
                notes_result, client_name, cy_year, py_year, rounding_unit,
                sheet_notes=result.get('notes_sheet_notes', {}),
                signatures=result.get('notes_signatures', {}),
                footers=result.get('notes_footers', {})
            )
        except Exception as e:
            logger.error(f"Failed to load Notes view: {e}")
            
        try:
            self.ratio_view.load_data(ratios, client_name, cy_year, py_year)
            self.dashboard_view.load_data(bs_result, pl_result, ratios, client_name, cy_year, py_year, rounding_unit)
        except Exception as e:
            logger.error(f"Failed to load Ratio/Dashboard views: {e}")
            
        try:
            self.remarks_view.set_data(bs_result, pl_result, form_data)
            self.comparison_view.load_data(self.app_data)
            self.export_view.set_export_data(self.app_data)
        except Exception as e:
            logger.error(f"Failed to setup Remarks/Comparison/Export views: {e}")
        
        val_summary = validation_report.summary_text()
        self.status_items.setText(
            f'BS: {len(bs_result)} items  |  PL: {len(pl_result)} items  |  {val_summary}')
        self.status_bar.showMessage('FAR generated successfully', 5000)
        
        # Switch to BS view
        self.sidebar.set_active_module(1)
        
        # Toast
        report_type = self.app_data.get('report_type', 'FAR')
        if validation_report.has_failures:
            self.toast.show_toast(
                f'⚠ {report_type} generated — {len(validation_report.failed)} validation issue(s) detected',
                'warning', 6000)
        else:
            self.toast.show_toast(f'✓ {report_type} generated successfully!', 'success')
        
        logger.info("FAR generated: BS=%d rows, PL=%d rows, Ratios=%d",
                    len(bs_result), len(pl_result), len(ratios))
    
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
