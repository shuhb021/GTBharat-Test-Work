"""
FAR Automation Tool — Sidebar Navigation
Left panel with module navigation buttons styled like VS Code's activity bar.
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
        self._update_style(False)
    
    def _update_style(self, active):
        if active:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #1E1E1E;
                    color: #FFFFFF;
                    border: none;
                    border-left: 3px solid #007ACC;
                    text-align: left;
                    padding-left: 12px;
                    font-size: 12px;
                    font-weight: 600;
                }
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #252526;
                    color: #CCCCCC;
                    border: none;
                    border-left: 3px solid transparent;
                    text-align: left;
                    padding-left: 12px;
                    font-size: 12px;
                    font-weight: 400;
                }
                QPushButton:hover {
                    background-color: #2A2D2E;
                    color: #FFFFFF;
                }
            """)
    
    def set_active(self, active):
        self.setChecked(active)
        self._update_style(active)


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
        
        # App title
        title_frame = QFrame()
        title_frame.setStyleSheet("""
            QFrame {
                background-color: #252526;
                border-bottom: 1px solid #3C3C3C;
                padding: 8px;
            }
        """)
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(16, 12, 16, 12)
        
        title_label = QLabel('FAR Tool')
        title_label.setStyleSheet("""
            QLabel {
                color: #007ACC;
                font-size: 18px;
                font-weight: 700;
                background: transparent;
            }
        """)
        title_layout.addWidget(title_label)
        
        version_label = QLabel('v1.0.0')
        version_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 10px;
                background: transparent;
            }
        """)
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
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("QFrame { background-color: #3C3C3C; max-height: 1px; }")
        layout.addWidget(sep)
        
        # Settings button
        self.settings_btn = SidebarButton('⚙️', 'Settings', -1)
        self.settings_btn.clicked.connect(lambda: self._on_button_clicked(-1))
        layout.addWidget(self.settings_btn)
        
        # Set initial active
        self.buttons[0].set_active(True)
        
        # Frame styling
        self.setStyleSheet("""
            QWidget#sidebarFrame {
                background-color: #252526;
                border-right: 1px solid #3C3C3C;
            }
        """)
    
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
