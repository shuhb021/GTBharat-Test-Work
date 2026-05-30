"""
FAR Automation Tool — Main Entry Point
Final Analytical Review Workpaper Desktop Application
"""

import sys
import os
import logging
from datetime import datetime

# Add project root to path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPalette, QColor, QFont
from PyQt6.QtCore import Qt


def setup_logging():
    """Configure application logging."""
    log_dir = os.path.join(project_root, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = os.path.join(log_dir, 'far_app.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info("FAR Automation Tool started at %s", datetime.now().isoformat())
    logger.info("Python %s", sys.version)
    logger.info("Project root: %s", project_root)
    return logger


def load_stylesheet():
    """Load the QSS stylesheet."""
    qss_path = os.path.join(project_root, 'src', 'assets', 'styles.qss')
    if os.path.exists(qss_path):
        with open(qss_path, 'r', encoding='utf-8') as f:
            return f.read()
    return ''


def setup_dark_palette(app):
    """Set up a dark color palette for the application."""
    palette = QPalette()
    
    palette.setColor(QPalette.ColorRole.Window, QColor('#1E1E1E'))
    palette.setColor(QPalette.ColorRole.WindowText, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.Base, QColor('#252526'))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor('#2A2D2E'))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor('#252526'))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor('#CCCCCC'))
    palette.setColor(QPalette.ColorRole.Text, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.Button, QColor('#3C3C3C'))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.BrightText, QColor('#007ACC'))
    palette.setColor(QPalette.ColorRole.Link, QColor('#007ACC'))
    palette.setColor(QPalette.ColorRole.Highlight, QColor('#007ACC'))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor('#FFFFFF'))
    
    # Disabled colors
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor('#666666'))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor('#666666'))
    
    app.setPalette(palette)


def main():
    """Application entry point."""
    logger = setup_logging()
    
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName('FAR Automation Tool')
    app.setApplicationVersion('1.0.0')
    app.setOrganizationName('FAR Automation')
    
    # Set default font
    font = QFont('Segoe UI', 11)
    app.setFont(font)
    
    # Apply dark palette
    setup_dark_palette(app)
    
    # Load QSS stylesheet
    stylesheet = load_stylesheet()
    if stylesheet:
        app.setStyleSheet(stylesheet)
        logger.info("Stylesheet loaded")
    
    # Global exception handler
    def exception_handler(exc_type, exc_value, exc_tb):
        logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    
    sys.excepthook = exception_handler
    
    # Create and show main window
    from src.gui.main_window import MainWindow
    
    window = MainWindow()
    window.show()
    
    logger.info("Main window shown")
    
    # Run event loop
    exit_code = app.exec()
    
    logger.info("Application exiting with code %d", exit_code)
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
