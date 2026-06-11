"""
FAR Automation Tool — Sidebar Navigation
Left panel with module navigation buttons — Light Purple Professional Theme.
Uses QSS property selectors instead of per-widget setStyleSheet for performance.
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, QFrame, QSizePolicy
)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QFont


class SidebarButton(QPushButton):
    """A sidebar navigation button with icon and label."""
    
    def __init__(self, icon_text, label_text, index, parent=None):
        super().__init__(parent)
        self.index = index
        self.setText(f"  {icon_text}  {label_text}")
        self.setCheckable(True)
        self.setFixedHeight(44)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFont(QFont('Segoe UI', 12))
        # Use property for QSS-driven active/inactive styling
        self.setProperty("sidebarActive", False)
    
    def set_active(self, active):
        self.setChecked(active)
        self.setProperty("sidebarActive", active)
        # Tell Qt to re-evaluate QSS for this widget only
        self.style().unpolish(self)
        self.style().polish(self)


class Sidebar(QWidget):
    """Left navigation panel with module buttons."""
    
    moduleChanged = pyqtSignal(int)
    
    MODULES = [
        ('📝', 'Input Form'),
        ('📊', 'Balance Sheet'),
        ('📈', 'Profit & Loss'),
        ('📋', 'Notes'),
        ('📐', 'Ratios'),
        ('📉', 'Dashboard'),
        ('⚖️', 'YoY Comparison'),
        ('🤖', 'AI Remarks'),
        ('💾', 'Export'),
    ]
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.buttons = []
        self.current_index = 0
        self._setup_ui()
    
    def _setup_ui(self):
        self.setFixedWidth(220)
        self.setObjectName('sidebarFrame')
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Firm name header
        title_frame = QFrame()
        title_frame.setObjectName('sidebarTitleFrame')
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(16, 14, 16, 6)
        
        firm_label = QLabel('Walker Chandiok & Co LLP')
        firm_label.setObjectName('sidebarFirmLabel')
        firm_label.setWordWrap(True)
        title_layout.addWidget(firm_label)
        
        title_label = QLabel('FAR Tool')
        title_label.setObjectName('sidebarTitleLabel')
        title_layout.addWidget(title_label)
        
        version_label = QLabel('v1.0.0')
        version_label.setObjectName('sidebarVersionLabel')
        title_layout.addWidget(version_label)
        
        layout.addWidget(title_frame)
        
        # Module buttons
        for i, (icon, label) in enumerate(self.MODULES):
            btn = SidebarButton(icon, label, i)
            btn.clicked.connect(lambda checked, idx=i: self._on_button_clicked(idx))
            self.buttons.append(btn)
            layout.addWidget(btn)
        
        # Spacer
        layout.addStretch()
        
        # Separator
        sep = QFrame()
        sep.setObjectName('sidebarSep')
        sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)
        
        # Settings button
        self.settings_btn = SidebarButton('⚙️', 'Settings', -1)
        self.settings_btn.clicked.connect(lambda: self._on_button_clicked(-1))
        layout.addWidget(self.settings_btn)
        
        # Set initial active
        self.buttons[0].set_active(True)
    
    def _on_button_clicked(self, index):
        # Deactivate all
        for btn in self.buttons:
            btn.set_active(False)
        self.settings_btn.set_active(False)
        
        # Activate selected
        if index >= 0:
            self.buttons[index].set_active(True)
            self.current_index = index
        else:
            self.settings_btn.set_active(True)
        
        self.moduleChanged.emit(index)
    
    def set_active_module(self, index):
        """Programmatically set the active module."""
        self._on_button_clicked(index)
