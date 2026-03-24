"""
Main entry point for Geo Processor ok for linux and windows
"""
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
from gui_handler import GeoProcessorApp

def main():
    """Main function to start the application"""
    app = QApplication(sys.argv)
    
    # Cross-platform font handling
    if sys.platform == "win32":
        font_family = "Pyidaungsu"
    elif sys.platform == "linux":
        font_family = "Noto Sans Myanmar"
    else:
        font_family = "Sans Serif"
    
    # Global application font
    app_font = QFont(font_family, 10)
    app.setFont(app_font)
    
    # Global stylesheet with dynamic font
    app.setStyleSheet(f"""
        * {{
            font-family: '{font_family}';
        }}
        QMainWindow {{
            background-color: #2b2b2b;
            color: #ffffff;
        }}
        QGroupBox {{
            font-weight: bold;
            border: 2px solid #555555;
            border-radius: 5px;
            margin-top: 1ex;
            padding-top: 10px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
            color: #ff9800;
        }}
        QLineEdit {{
            padding: 5px;
            border: 1px solid #555555;
            border-radius: 3px;
            background-color: #3c3c3c;
            color: #ffffff;
        }}
        QPushButton {{
            background-color: #00796b;
            color: white;
            padding: 5px 10px;
            border: none;
            border-radius: 3px;
        }}
        QPushButton:hover {{
            background-color: #004d40;
        }}
        QComboBox {{
            padding: 5px;
            border: 1px solid #555555;
            border-radius: 3px;
            background-color: #3c3c3c;
            color: #ffffff;
        }}
        QProgressBar {{
            border: 1px solid #555555;
            border-radius: 3px;
            text-align: center;
            color: white;
        }}
        QProgressBar::chunk {{
            background-color: #ff9800;
        }}
        QTextEdit {{
            background-color: #3c3c3c;
            color: #ffffff;
            border: 1px solid #555555;
            border-radius: 3px;
        }}
    """)
    
    # Create and show main window
    window = GeoProcessorApp()
    window.show()
    
    # Start application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()