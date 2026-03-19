"""
GUI components for Geo Processor
"""
import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QFileDialog,
    QMessageBox, QProgressBar, QTextEdit, QGroupBox, QRadioButton
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

# Import processing functions
from core_processor import process_vector
from utils import convert_to_mm_digits, round_coordinate_for_phrase

# ==================== PROCESSING THREAD ====================
class ProcessingThread(QThread):
    """Background thread for processing"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, input_file, output_file, proj_choice, zone_mode, manual_zone):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.proj_choice = proj_choice
        self.zone_mode = zone_mode
        self.manual_zone = manual_zone
    
    def run(self):
        """Run processing in background thread"""
        try:
            self.message.emit("ဒေတာဖိုင်ဖတ်နေသည်...")
            
            # Process vector file
            rows = process_vector(
                self.input_file, 
                self.proj_choice, 
                self.zone_mode, 
                self.manual_zone
            )
            
            self.progress.emit(30)
            
            if not rows:
                self.finished.emit(False, "ရွေးချယ်ထားသော ဖိုင်မှ Point, Line, သို့မဟုတ် Polygon Data များကို ရှာမတွေ့ပါ။")
                return
            
            self.message.emit("Excel ဖိုင်ထုတ်နေသည်...")
            
            # Import here to avoid circular imports
            from core_processor import ExcelExporter
            
            # Export to Excel
            exporter = ExcelExporter(self.proj_choice)
            exporter.export(rows, self.output_file)
            
            self.progress.emit(100)
            
            self.finished.emit(
                True, 
                f"အချက်အလက်များ အောင်မြင်စွာ ပြုပြင်ပြီး Excel ဖိုင် (၅)ခု ထုတ်ပြီးပါပြီ။\n"
                f"သိမ်းဆည်းထားသော လမ်းကြောင်း: {self.output_file}"
            )
            
        except Exception as e:
            self.finished.emit(False, f"Data ပြုပြင်နေစဉ် အမှားရှိခဲ့သည်:\n{str(e)}")

# ==================== MAIN APPLICATION WINDOW ====================
class GeoProcessorApp(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self._init_ui()
        self._setup_connections()
    
    def _init_ui(self):
        """Initialize UI components"""
        self.setWindowTitle("Geospatial Data Processor Tool")
        self.setGeometry(100, 100, 600, 500)
        
        # Set dark theme
        self._set_dark_theme()
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Add all UI components
        self._create_title(layout)
        self._create_input_section(layout)
        self._create_output_section(layout)
        self._create_projection_section(layout)
        self._create_zone_section(layout)
        self._create_progress_section(layout)
        self._create_process_button(layout)
    
    def _create_title(self, layout):
        """Create title label"""
        title = QLabel("Geospatial Data Processor Tool")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18pt; font-weight: bold; color: #ffffff; margin: 10px;")
        layout.addWidget(title)
    
    def _create_input_section(self, layout):
        """Create input file section"""
        input_group = QGroupBox("Input Data ဖိုင်")
        input_layout = QHBoxLayout(input_group)
        
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setPlaceholderText("Input data file path...")
        input_layout.addWidget(self.input_file_edit)
        
        self.browse_input_btn = QPushButton("ရွေးချယ်ရန်")
        self.browse_input_btn.clicked.connect(self._browse_input_file)
        input_layout.addWidget(self.browse_input_btn)
        
        layout.addWidget(input_group)
    
    def _create_output_section(self, layout):
        """Create output file section"""
        output_group = QGroupBox("Output Excel ဖိုင်")
        output_layout = QHBoxLayout(output_group)
        
        self.output_file_edit = QLineEdit()
        self.output_file_edit.setPlaceholderText("Output Excel file path...")
        output_layout.addWidget(self.output_file_edit)
        
        self.browse_output_btn = QPushButton("သိမ်းရန်")
        self.browse_output_btn.clicked.connect(self._browse_output_file)
        output_layout.addWidget(self.browse_output_btn)
        
        layout.addWidget(output_group)
    
    def _create_projection_section(self, layout):
        """Create projection selection section"""
        proj_group = QGroupBox("Coordinate စနစ်")
        proj_layout = QHBoxLayout(proj_group)
        
        self.proj_combo = QComboBox()
        self.proj_combo.addItem("MYANMAR Datum UTM (Easting/Northing)", "Custom_UTM")
        self.proj_combo.addItem("WGS84 UTM (Easting/Northing)", "WGS84_UTM")
        self.proj_combo.addItem("WGS84 (Lon/Lat)", "WGS84_LatLon")
        proj_layout.addWidget(self.proj_combo)
        
        layout.addWidget(proj_group)
    
    def _create_zone_section(self, layout):
        """Create zone selection section"""
        zone_group = QGroupBox("Zone Selection")
        zone_layout = QVBoxLayout(zone_group)
        
        # Zone mode selection
        zone_mode_layout = QHBoxLayout()
        self.zone_auto_radio = QRadioButton("Auto")
        self.zone_manual_46_radio = QRadioButton("Manual: 46")
        self.zone_manual_47_radio = QRadioButton("Manual: 47")
        
        self.zone_auto_radio.setChecked(True)  # Default to Auto
        
        zone_mode_layout.addWidget(self.zone_auto_radio)
        zone_mode_layout.addWidget(self.zone_manual_46_radio)
        zone_mode_layout.addWidget(self.zone_manual_47_radio)
        zone_layout.addLayout(zone_mode_layout)
        
        # Threshold display
        threshold_layout = QHBoxLayout()
        self.threshold_label = QLabel("MMD: 95.9968784133°")
        self.threshold_label.setStyleSheet("color: green; font-weight: bold;")
        threshold_layout.addWidget(QLabel("Threshold:"))
        threshold_layout.addWidget(self.threshold_label)
        threshold_layout.addStretch()
        zone_layout.addLayout(threshold_layout)
        
        layout.addWidget(zone_group)
    
    def _create_progress_section(self, layout):
        """Create progress section"""
        progress_group = QGroupBox("လုပ်ဆောင်ချက်")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setReadOnly(True)
        progress_layout.addWidget(self.log_text)
        
        layout.addWidget(progress_group)
    
    def _create_process_button(self, layout):
        """Create process button"""
        self.process_btn = QPushButton("စတင်ပြုပြင်မည်")
        self.process_btn.clicked.connect(self._process_data)
        self.process_btn.setStyleSheet(
            "QPushButton { background-color: #ff9800; color: white; font-weight: bold; padding: 10px; }"
        )
        layout.addWidget(self.process_btn)
    
    def _setup_connections(self):
        """Setup signal connections"""
        self.zone_auto_radio.toggled.connect(self._on_zone_mode_changed)
        self.zone_manual_46_radio.toggled.connect(self._on_zone_mode_changed)
        self.zone_manual_47_radio.toggled.connect(self._on_zone_mode_changed)
        self.proj_combo.currentTextChanged.connect(self._on_projection_changed)
    
    def _set_dark_theme(self):
        """Set dark theme stylesheet"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #555555;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #ff9800;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QPushButton {
                background-color: #00796b;
                color: white;
                padding: 5px 10px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #004d40;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QProgressBar {
                border: 1px solid #555555;
                border-radius: 3px;
                text-align: center;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #ff9800;
            }
            QTextEdit {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 3px;
            }
        """)
    
    def _on_zone_mode_changed(self, checked):
        """Update threshold display based on zone mode"""
        if not checked:
            return
        
        if self.zone_manual_46_radio.isChecked():
            self.threshold_label.setText("Zone 46")
            self.threshold_label.setStyleSheet("color: orange; font-weight: bold;")
        elif self.zone_manual_47_radio.isChecked():
            self.threshold_label.setText("Zone 47")
            self.threshold_label.setStyleSheet("color: orange; font-weight: bold;")
        else:  # Auto mode
            proj_data = self.proj_combo.currentData()
            if proj_data == "Custom_UTM":
                self.threshold_label.setText("MMD: 95.9968784133°")
                self.threshold_label.setStyleSheet("color: green; font-weight: bold;")
            else:
                self.threshold_label.setText("WGS: 96.0°")
                self.threshold_label.setStyleSheet("color: green; font-weight: bold;")
    
    def _on_projection_changed(self):
        """Update threshold display when projection changes"""
        if self.zone_auto_radio.isChecked():
            proj_data = self.proj_combo.currentData()
            if proj_data == "Custom_UTM":
                self.threshold_label.setText("MMD: 95.9968784133°")
            else:
                self.threshold_label.setText("WGS: 96.0°")
    
    def _browse_input_file(self):
        """Browse for input file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Input Data ဖိုင်ရွေးချယ်ရန်",
            "",
            "Geospatial Files (*.kmz *.kml *.shp *.geojson *.gpkg);;All Files (*.*)"
        )
        
        if file_path:
            self.input_file_edit.setText(file_path)
            
            # Suggest output filename
            dir_name = os.path.dirname(file_path)
            base_name = os.path.basename(file_path)
            file_name_without_ext = os.path.splitext(base_name)[0]
            default_output = os.path.join(dir_name, f"{file_name_without_ext}.xlsx")
            self.output_file_edit.setText(default_output)
    
    def _browse_output_file(self):
        """Browse for output file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Output Excel ဖိုင်သိမ်းဆည်းရန်",
            self.output_file_edit.text(),
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        
        if file_path:
            self.output_file_edit.setText(file_path)
    
    def _process_data(self):
        """Process data when button is clicked"""
        # Get input values
        input_file = self.input_file_edit.text()
        output_file = self.output_file_edit.text()
        proj_choice = self.proj_combo.currentData()
        
        # Validation
        if not input_file or not os.path.exists(input_file):
            QMessageBox.critical(self, "အမှား", "ကျေးဇူးပြု၍ Input Data ဖိုင်လမ်းကြောင်းကို မှန်ကန်စွာ ရွေးချယ်ပါ။")
            return
        
        if not output_file or not output_file.lower().endswith(('.xlsx', '.xls')):
            QMessageBox.critical(self, "အမှား", "ကျေးဇူးပြု၍ Output ဖိုင်အမည်ကို .xlsx (သို့) .xls ဖြင့် အဆုံးသတ်ပါ။")
            return
        
        # Check geospatial libraries
        try:
            import geopandas as gpd
            from pyproj import Transformer, CRS
            from openpyxl.styles import Font, Alignment
        except ImportError:
            QMessageBox.critical(
                self, 
                "Dependency Error", 
                "Geopandas, pyproj, သို့မဟုတ် openpyxl ကို ရှာမတွေ့ပါ။\n"
                "'pip install geopandas openpyxl pandas pyproj' ဖြင့် install လုပ်ရန် လိုအပ်ပါသည်။"
            )
            return
        
        # Show warning for manual zone
        if self.zone_manual_46_radio.isChecked() or self.zone_manual_47_radio.isChecked():
            zone_name = "46" if self.zone_manual_46_radio.isChecked() else "47"
            
            warning_msg = QMessageBox()
            warning_msg.setIcon(QMessageBox.Warning)
            warning_msg.setWindowTitle("Manual Zone Selected")
            warning_msg.setText(f"<b>Manual Zone {zone_name} ရွေးထားပါတယ်</b>")
            warning_msg.setInformativeText(
                f"သတိပြုရန်: Manual Zone {zone_name} ကို ရွေးထားပါတယ်။\n\n"
                f"ဒေတာထဲမှာရှိတဲ့ longitude တန်ဖိုးတွေကို မစစ်ဆေးတော့ပါ။\n"
                f"မှားယွင်းတဲ့ zone သုံးမိရင် coordinates တွေ မှားနိုင်ပါတယ်။\n\n"
                f"ဆက်လက်လုပ်ဆောင်ပါမည်လား?"
            )
            warning_msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            warning_msg.setDefaultButton(QMessageBox.No)
            
            reply = warning_msg.exec_()
            if reply != QMessageBox.Yes:
                return
        
        # Get zone parameters
        if self.zone_auto_radio.isChecked():
            zone_mode = "auto"
            manual_zone = None
        elif self.zone_manual_46_radio.isChecked():
            zone_mode = "manual"
            manual_zone = 46
        else:  # manual 47
            zone_mode = "manual"
            manual_zone = 47
        
        # Disable UI during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        # Start processing thread
        self.thread = ProcessingThread(input_file, output_file, proj_choice, zone_mode, manual_zone)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.message.connect(self.log_text.append)
        self.thread.finished.connect(self._processing_finished)
        self.thread.start()
    
    def _processing_finished(self, success, message):
        """Handle processing completion"""
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "အောင်မြင်ပါသည်", message)
        else:
            QMessageBox.critical(self, "လုပ်ဆောင်ချက် အမှား", message)