"""
FAR Automation Tool — Settings View
Provides controls for adjusting font style, font size, and UI zoom levels.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QComboBox, QLabel, QFrame, QPushButton
)
from PyQt6.QtCore import pyqtSignal, Qt


class SettingsView(QWidget):
    """
    SettingsView allows configuration of typography and display scaling.
    Changes are applied instantly across the entire application.
    """
    
    settingsChanged = pyqtSignal(str, int, int)  # family, size, zoom
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = {
            'font_family': 'Segoe UI',
            'font_size': 11,
            'ui_zoom': 100
        }
        self._setup_ui()
        
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 24)
        layout.setSpacing(20)
        
        # Title
        title = QLabel('⚙️ Settings')
        title.setObjectName('headingLabel')
        layout.addWidget(title)
        
        subtitle = QLabel('Customize typography and application scaling preferences.')
        subtitle.setObjectName('mutedLabel')
        layout.addWidget(subtitle)
        
        # Settings Card Frame
        settings_group = QGroupBox('Appearance Settings')
        form_layout = QFormLayout(settings_group)
        form_layout.setSpacing(16)
        form_layout.setContentsMargins(20, 24, 20, 20)
        
        # 1. Font Family Dropdown
        self.font_family_cb = QComboBox()
        font_families = ['Segoe UI', 'Arial', 'Times New Roman', 'Helvetica', 'Georgia', 'Courier New']
        self.font_family_cb.addItems(font_families)
        self.font_family_cb.currentTextChanged.connect(self._on_setting_changed)
        form_layout.addRow('Font Family / Style:', self.font_family_cb)
        
        # 2. Font Size Dropdown
        self.font_size_cb = QComboBox()
        font_sizes = ['9', '10', '11', '12', '13', '14', '16', '18']
        self.font_size_cb.addItems(font_sizes)
        self.font_size_cb.currentTextChanged.connect(self._on_setting_changed)
        form_layout.addRow('Font Size (Base):', self.font_size_cb)
        
        # 3. UI Zoom Scale Dropdown
        self.ui_zoom_cb = QComboBox()
        zoom_levels = ['75%', '100%', '125%', '150%', '175%', '200%']
        self.ui_zoom_cb.addItems(zoom_levels)
        self.ui_zoom_cb.currentTextChanged.connect(self._on_setting_changed)
        form_layout.addRow('UI Zoom level:', self.ui_zoom_cb)
        
        layout.addWidget(settings_group)
        
        # Informational description card at the bottom
        info_frame = QFrame()
        info_frame.setObjectName('cardFrame')
        info_layout = QVBoxLayout(info_frame)
        
        info_title = QLabel('💡 Layout Auto-Scaling')
        info_title.setObjectName('subheadingLabel')
        info_layout.addWidget(info_title)
        
        info_desc = QLabel(
            'Adjusting the base font size or UI zoom level scales all headings, '
            'tables, text inputs, sidebars, and control components in proportion. '
            'Settings are persisted across application restarts.'
        )
        info_desc.setObjectName('mutedLabel')
        info_desc.setWordWrap(True)
        info_layout.addWidget(info_desc)
        
        layout.addWidget(info_frame)
        layout.addStretch()
        
    def load_config(self, config):
        """Pre-populate selectors from loaded configuration."""
        self.config = config
        
        # Font Family
        idx_fam = self.font_family_cb.findText(config.get('font_family', 'Segoe UI'))
        if idx_fam >= 0:
            self.font_family_cb.setCurrentIndex(idx_fam)
            
        # Font Size
        idx_size = self.font_size_cb.findText(str(config.get('font_size', 11)))
        if idx_size >= 0:
            self.font_size_cb.setCurrentIndex(idx_size)
            
        # UI Zoom
        zoom_str = f"{config.get('ui_zoom', 100)}%"
        idx_zoom = self.ui_zoom_cb.findText(zoom_str)
        if idx_zoom >= 0:
            self.ui_zoom_cb.setCurrentIndex(idx_zoom)
            
    def _on_setting_changed(self, text):
        """Collect current values and emit change signal."""
        family = self.font_family_cb.currentText()
        try:
            size = int(self.font_size_cb.currentText())
        except ValueError:
            size = 11
            
        zoom_str = self.ui_zoom_cb.currentText().replace('%', '')
        try:
            zoom = int(zoom_str)
        except ValueError:
            zoom = 100
            
        self.settingsChanged.emit(family, size, zoom)
