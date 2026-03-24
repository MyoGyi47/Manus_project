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
# ==================== PROCESSING THREAD ====================
class ProcessingThread(QThread):
    """Background thread for processing"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, input_file, output_file, projection, datum, zone_mode, manual_zone, detected_zone=None):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.projection = projection  # "UTM" or "Geographic"
        self.datum = datum  # "MMD2000" or "WGS84"
        self.zone_mode = zone_mode
        self.manual_zone = manual_zone
        self.detected_zone = detected_zone
    
    def run(self):
        try:
            self.message.emit("ဒေတာဖိုင်ဖတ်နေသည်...")
            
            # Use detected zone if available
            if self.zone_mode == "auto" and self.detected_zone is not None:
                actual_zone_mode = "manual"
                actual_manual_zone = self.detected_zone
                zone_msg = f"Zone {self.detected_zone}"
            else:
                actual_zone_mode = self.zone_mode
                actual_manual_zone = self.manual_zone
                zone_msg = f"{self.zone_mode} mode"
            
            self.message.emit(f"Processing with: {self.projection} | {self.datum} | {zone_msg}")
            
            # Process vector file with new parameters
            rows = process_vector(
                self.input_file, 
                self.projection,  # "UTM" or "Geographic"
                self.datum,  # "MMD2000" or "WGS84"
                actual_zone_mode, 
                actual_manual_zone
            )          
                        
            self.progress.emit(30)
            
            if not rows:
                self.finished.emit(False, "ရွေးချယ်ထားသော ဖိုင်မှ Point, Line, သို့မဟုတ် Polygon Data များကို ရှာမတွေ့ပါ။")
                return
            
            self.message.emit("Excel ဖိုင်ထုတ်နေသည်...")
            
            # Import here to avoid circular imports
            from core_processor import ExcelExporter
            
            # Determine proj_choice for ExcelExporter (legacy compatibility)
            # proj_choice တန်ဖိုးကို projection နဲ့ datum ပေါ်မူတည်ပြီးဆုံးဖြတ်
            if self.projection == "UTM" and self.datum == "WGS84":
                proj_choice = "WGS84_UTM"
            elif self.projection == "UTM" and self.datum == "MMD2000":
                proj_choice = "Custom_UTM"
            elif self.projection == "Geographic" and self.datum == "WGS84":
                proj_choice = "WGS84_LatLon"
            else:
                proj_choice = "Custom_LatLon"
            
            # Export to Excel
            exporter = ExcelExporter(proj_choice)  # proj_choice ပို့မယ်
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
# gui_handler.py - UI အသစ်
class GeoProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self._current_detected_zone = None
        self._init_ui()
        self._setup_connections()
        self._update_ui_based_on_selection()  # Initial update
        
    
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
        
        # Initialize UI state
        self._update_ui_based_on_selection()  # <-- ဒါကို သုံးပါ
    
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
        """Create 3-layer projection selection"""
        proj_group = QGroupBox("Coordinate System")
        proj_layout = QVBoxLayout(proj_group)
        
        # 1. Projection selection
        proj_row = QHBoxLayout()
        proj_row.addWidget(QLabel("Projection:"))
        
        self.projection_combo = QComboBox()
        self.projection_combo.addItem("UTM (Easting/Northing)", "UTM")
        self.projection_combo.addItem("Geographic (Lon/Lat)", "Geographic")
        proj_row.addWidget(self.projection_combo)
        proj_row.addStretch()
        proj_layout.addLayout(proj_row)
        
        # 2. Datum selection
        datum_row = QHBoxLayout()
        datum_row.addWidget(QLabel("Datum:"))
        
        self.datum_combo = QComboBox()        
        self.datum_combo.addItem("MMD2000", "MMD2000") 
        self.datum_combo.addItem("WGS84", "WGS84")       
        datum_row.addWidget(self.datum_combo)
        datum_row.addStretch()
        proj_layout.addLayout(datum_row)
        
        layout.addWidget(proj_group)
    
    def _create_zone_section(self, layout):
        """Create zone selection section"""
        zone_group = QGroupBox("Zone Selection")
        self.zone_group = zone_group
        zone_layout = QVBoxLayout(zone_group)
        
        # Zone mode selection
        zone_mode_layout = QHBoxLayout()
        self.zone_auto_radio = QRadioButton("Auto")
        self.zone_manual_46_radio = QRadioButton("Manual: 46")
        self.zone_manual_47_radio = QRadioButton("Manual: 47")
        
        self.zone_auto_radio.setChecked(True)
        
        zone_mode_layout.addWidget(self.zone_auto_radio)
        zone_mode_layout.addWidget(self.zone_manual_46_radio)
        zone_mode_layout.addWidget(self.zone_manual_47_radio)
        zone_layout.addLayout(zone_mode_layout)
        
        # Threshold/Info display
        threshold_layout = QHBoxLayout()
        self.info_label = QLabel("Select projection and datum")
        self.info_label.setStyleSheet("color: green; font-weight: bold;")
        threshold_layout.addWidget(QLabel("Info:"))
        threshold_layout.addWidget(self.info_label)
        threshold_layout.addStretch()
        zone_layout.addLayout(threshold_layout)
        
        layout.addWidget(zone_group)
        return zone_group
    
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
        self.projection_combo.currentTextChanged.connect(self._on_selection_changed)
        self.datum_combo.currentTextChanged.connect(self._on_selection_changed)
        self.zone_auto_radio.toggled.connect(self._on_zone_mode_changed)
        self.zone_manual_46_radio.toggled.connect(self._on_zone_mode_changed)
        self.zone_manual_47_radio.toggled.connect(self._on_zone_mode_changed)
    
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

    def _update_zone_info(self, filename):
        """Update zone info based on filename"""
        projection = self.projection_combo.currentData()
        datum = self.datum_combo.currentData()

        # BRIGHTER GREEN COLOR
        bright_green = "color: #32CD32; font-weight: bold;"
        
        # Geographic ဆိုရင် zone info မလို
        if projection == "Geographic":
            if datum == "WGS84":
                self.info_label.setText("WGS84 Geographic (EPSG:4326)")
            else:
                self.info_label.setText("MMD2000 Geographic")
            return
        
        # UTM အတွက် filename check
        import re
        import os
        
        base_name = os.path.basename(filename)
        name_without_ext = os.path.splitext(base_name)[0]
        clean_name = re.sub(r'[^a-zA-Z0-9\s\-]', ' ', name_without_ext)
        
        pattern = r'(\d{4})[\s\-_](\d{1,2})'
        match = re.search(pattern, clean_name)
        
        if match:
            year_part = match.group(1)
            number_part = match.group(2)
            
            if len(year_part) == 4:
                year_last_two = year_part[2:]
            else:
                year_last_two = year_part
            
            try:
                year_num = int(year_last_two)
                num = int(number_part)
                
                if year_num == 95 and num in [13, 14, 15, 16]:
                    zone_info = f"{datum} Zone 46 ဖြင့်အလုပ်လုပ်နေပါသည်"
                    self._current_detected_zone = 46
                elif year_num == 96 and num in [1, 2, 3, 4]:
                    zone_info = f"{datum} Zone 47 ဖြင့်အလုပ်လုပ်နေပါသည်"
                    self._current_detected_zone = 47
                else:
                    zone_info = f"{datum} UTM (Auto detection)"
                    self._current_detected_zone = None
                
                # When setting text, also set bright color
                self.info_label.setText(zone_info)
                self.info_label.setStyleSheet(bright_green)  # <-- ဒီမှာထည့်
                return
                
            except ValueError:
                pass
        
        # No pattern found
        self.info_label.setText(f"{datum} UTM (Auto detection)")
        self._current_detected_zone = None

    def _on_zone_mode_changed(self, checked):
        """Update when zone mode changes"""
        if not checked:
            return
        
        projection = self.projection_combo.currentData()
        if projection == "Geographic":
            return
        
        datum = self.datum_combo.currentData()
        
        if self.zone_manual_46_radio.isChecked():
            self.info_label.setText(f"{datum} Zone 46 (Manual)")
            self.info_label.setStyleSheet("color: orange; font-weight: bold;")
        elif self.zone_manual_47_radio.isChecked():
            self.info_label.setText(f"{datum} Zone 47 (Manual)")
            self.info_label.setStyleSheet("color: orange; font-weight: bold;")
        else:  # Auto mode
            bright_green = "color: #32CD32; font-weight: bold;"  # <-- Bright green
            filename = self.input_file_edit.text()
            if filename and os.path.exists(filename):
                self._update_zone_info(filename)
            else:
                self.info_label.setStyleSheet(bright_green)  # <-- Bright green
                self.info_label.setText(f"{datum} UTM (Auto detection)")
                self.info_label.setStyleSheet("color: green; font-weight: bold;")
    
    
    def _on_selection_changed(self):
        """When projection or datum changes"""
        self._update_ui_based_on_selection()

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
            
            # Update zone info
            self._update_zone_info(file_path)
    
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
        """Process data - updated for 3-layer system"""
        # Get values from UI
        input_file = self.input_file_edit.text()
        output_file = self.output_file_edit.text()
        projection = self.projection_combo.currentData()
        datum = self.datum_combo.currentData()
        
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
            detected_zone = self._current_detected_zone
        elif self.zone_manual_46_radio.isChecked():
            zone_mode = "manual"
            manual_zone = 46
            detected_zone = None
        else:  # manual 47
            zone_mode = "manual"
            manual_zone = 47
            detected_zone = None
        
        # Disable UI during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
                
        # Start processing thread with zone parameters       
        self.thread = ProcessingThread(
            input_file, output_file, 
            projection, datum,  # NEW: projection and datum
            zone_mode, manual_zone, detected_zone
        )
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

    def _update_ui_based_on_selection(self):
        """Update UI based on current selections"""
        projection = self.projection_combo.currentData()
        datum = self.datum_combo.currentData()
        
        # BRIGHTER GREEN COLOR
        bright_green = "color: #32CD32; font-weight: bold;"  # Lime Green
        
        if projection == "Geographic":
            # Geographic ဆိုရင် zone selection ကို disable
            self.zone_group.setEnabled(False)
            self.zone_group.setStyleSheet("color: gray;")
            
            # Update info label
            if datum == "WGS84":
                self.info_label.setText("WGS84 Geographic (EPSG:4326)")
                self.info_label.setStyleSheet(bright_green)  # <-- အရောင်တောက်
            else:
                self.info_label.setText("MMD2000 Geographic")
                self.info_label.setStyleSheet(bright_green)  # <-- အရောင်တောက်
                    
        else:  # UTM
            # UTM ဆိုရင် zone selection ကို enable
            self.zone_group.setEnabled(True)
            self.zone_group.setStyleSheet("")
            
            # Check filename for zone info
            filename = self.input_file_edit.text()
            if filename and os.path.exists(filename):
                self._update_zone_info(filename)
            else:
                # Default info with bright color
                if datum == "WGS84":
                    self.info_label.setText("WGS84 UTM (Auto detection)")
                else:
                    self.info_label.setText("MMD2000 UTM (Auto detection)")
                self.info_label.setStyleSheet(bright_green)  # <-- အရောင်တောက်
