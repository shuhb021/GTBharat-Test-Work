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


def setup_light_palette(app):
    """Set up a light purple color palette for the application."""
    palette = QPalette()
    
    palette.setColor(QPalette.ColorRole.Window, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.WindowText, QColor('#2D2D2D'))
    palette.setColor(QPalette.ColorRole.Base, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor('#F9F7FC'))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor('#FFFFFF'))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor('#2D2D2D'))
    palette.setColor(QPalette.ColorRole.Text, QColor('#2D2D2D'))
    palette.setColor(QPalette.ColorRole.Button, QColor('#F0F0F0'))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor('#2D2D2D'))
    palette.setColor(QPalette.ColorRole.BrightText, QColor('#4A1A6B'))
    palette.setColor(QPalette.ColorRole.Link, QColor('#4A1A6B'))
    palette.setColor(QPalette.ColorRole.Highlight, QColor('#4A1A6B'))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor('#FFFFFF'))
    
    # Disabled colors
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor('#AAAAAA'))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor('#AAAAAA'))
    
    app.setPalette(palette)


def load_config():
    import json
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


def main():
    """Application entry point."""
    logger = setup_logging()
    
    config = load_config()
    
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName('FAR Automation Tool')
    app.setApplicationVersion('1.0.0')
    app.setOrganizationName('FAR Automation')
    
    # Set default font
    family = config['font_family']
    size = config['font_size']
    zoom = config['ui_zoom']
    zoom_factor = zoom / 100.0
    size_factor = size / 11.0
    total_factor = zoom_factor * size_factor
    
    scaled_size = int(round(size * zoom_factor))
    font = QFont(family, scaled_size)
    app.setFont(font)
    
    # Apply light purple palette
    setup_light_palette(app)
    
    # Load QSS stylesheet
    stylesheet = load_stylesheet()
    if stylesheet:
        import re
        try:
            # Replace font family in stylesheet
            stylesheet = re.sub(r"font-family:\s*[^;]+;", f"font-family: '{family}', sans-serif;", stylesheet)
            
            # Scale font size in stylesheet
            def scale_size(match):
                val = int(match.group(1))
                new_val = max(8, int(round(val * total_factor)))
                return f"font-size: {new_val}px"
            stylesheet = re.sub(r"font-size:\s*(\d+)\s*px", scale_size, stylesheet)
        except Exception as e:
            logger.warning("Failed to scale QSS at startup: %s", e)
            
        app.setStyleSheet(stylesheet)
        logger.info("Scaled stylesheet loaded")
    
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
