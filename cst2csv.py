import sys
import math
import pandas as pd
import numpy as np
import json
import os
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, 
                           QTableWidget, QTableWidgetItem, QVBoxLayout, QCheckBox, QLineEdit, 
                           QWidget, QSpinBox, QScrollArea, QGroupBox, QHBoxLayout, QDoubleSpinBox,
                           QComboBox, QTabWidget, QTextEdit, QDialog)
from PyQt5.QtCore import Qt

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')

class LibraryPathDialog(QDialog):
    def __init__(self, parent=None, current_path=None):
        super().__init__(parent)
        self.setWindowTitle("CST Library Path Configuration")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        
        # Path input
        path_group = QGroupBox("CST Python Library Path")
        path_layout = QHBoxLayout()
        
        self.path_input = QLineEdit(current_path if current_path else "")
        self.browse_btn = QPushButton("Browse")
        self.browse_btn.clicked.connect(self.browse_path)
        
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_btn)
        path_group.setLayout(path_layout)
        layout.addWidget(path_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.accept)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)
        
        self.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
        """)
    
    def browse_path(self):
        path = QFileDialog.getExistingDirectory(self, "Select CST Python Libraries Directory")
        if path:
            self.path_input.setText(path)
    
    def get_path(self):
        return self.path_input.text()

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def setup_cst_path(parent=None):
    config = load_config()
    current_path = config.get('cst_library_path')
    try:
        import cst.results
    except ImportError:
        current_path = None
    
    # Check if path exists and is valid
    path_valid = False
    if current_path:
        try:
            sys.path.append(current_path)
            import cst.results
            path_valid = True
        except ImportError:
            sys.path.remove(current_path)
    
    # Show dialog if no path or invalid path
    if not path_valid:
        dialog = LibraryPathDialog(parent, current_path)
        if dialog.exec_() == QDialog.Accepted:
            new_path = dialog.get_path()
            try:
                sys.path.append(new_path)
                import cst.results
                config['cst_library_path'] = new_path
                save_config(config)
                return True
            except ImportError:
                sys.path.remove(new_path)
                QtWidgets.QMessageBox.critical(parent, "Error", 
                    "Could not find CST Python libraries in the selected path.\n"
                    "Please make sure you select the correct directory containing 'cst' module.")
                return setup_cst_path(parent)  # Try again
        return False
    return True

class CSTExportApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.parameter_checkboxes = {}
        self.freq_data = None
        self.s_parameters = None
        self.parameters = None
        self.full_data = None
        self.MAX_DISPLAY_ROWS = 500
        self.initUI()

    def initUI(self):
        self.setWindowTitle("CST to CSV Export Tool")
        self.setGeometry(100, 100, 1200, 800)

        # Create menu bar
        menubar = self.menuBar()
        settings_menu = menubar.addMenu('Settings')
        
        # Add CST Library Path action
        change_path_action = settings_menu.addAction('Change CST Library Path')
        change_path_action.triggered.connect(self.change_library_path)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Set application icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app_icon.svg')
        self.setWindowIcon(QtGui.QIcon(icon_path))

        # File selection section
        file_group = QGroupBox("File Selection")
        file_layout = QHBoxLayout()
        self.filePathLineEdit = QLineEdit()
        self.browseButton = QPushButton("Browse")
        self.browseButton.clicked.connect(self.browseFile)
        file_layout.addWidget(self.filePathLineEdit)
        file_layout.addWidget(self.browseButton)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # Create tab widget
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Data tab
        data_tab = QWidget()
        data_layout = QVBoxLayout(data_tab)

        # S11 Display Options
        display_group = QGroupBox("S11 Display Options")
        display_layout = QHBoxLayout()
        self.display_mode = QComboBox()
        self.display_mode.addItems([
            "Complex (Real + Imaginary)",
            "Magnitude",
            "Magnitude (dB)",
            "Magnitude and Phase"
        ])
        self.display_mode.currentIndexChanged.connect(self.update_display)
        display_layout.addWidget(self.display_mode)
        display_group.setLayout(display_layout)
        data_layout.addWidget(display_group)

        # Frequency range selection
        freq_group = QGroupBox("Frequency Range Selection")
        freq_layout = QHBoxLayout()
        
        self.freq_start = QLineEdit()
        self.freq_start.setPlaceholderText("Start Frequency (GHz)")
        
        self.freq_end = QLineEdit()
        self.freq_end.setPlaceholderText("End Frequency (GHz)")
        
        freq_layout.addWidget(QLabel("Start Frequency:"))
        freq_layout.addWidget(self.freq_start)
        freq_layout.addWidget(QLabel("End Frequency:"))
        freq_layout.addWidget(self.freq_end)
        
        self.use_all_freq = QCheckBox("Use All Frequencies")
        self.use_all_freq.setChecked(True)
        self.use_all_freq.stateChanged.connect(self.toggle_freq_range)
        freq_layout.addWidget(self.use_all_freq)
        
        self.apply_freq_range = QPushButton("Apply Range")
        self.apply_freq_range.clicked.connect(self.update_display)
        freq_layout.addWidget(self.apply_freq_range)
        
        freq_group.setLayout(freq_layout)
        data_layout.addWidget(freq_group)

        # Results table with header checkboxes
        table_group = QGroupBox("Results")
        table_layout = QVBoxLayout()
        
        self.header_widget = QWidget()
        self.header_layout = QHBoxLayout()
        self.header_layout.setSpacing(0)
        self.header_layout.setContentsMargins(0, 0, 0, 0)
        self.header_widget.setLayout(self.header_layout)
        table_layout.addWidget(self.header_widget)
        
        self.tableWidget = QTableWidget()
        self.tableWidget.setMinimumHeight(400)
        table_layout.addWidget(self.tableWidget)
        
        table_group.setLayout(table_layout)
        data_layout.addWidget(table_group)

        # Export section
        export_group = QGroupBox("Export Options")
        export_layout = QHBoxLayout()
        
        # File format selection
        export_layout.addWidget(QLabel("Export Format:"))
        self.export_format = QComboBox()
        self.export_format.addItems(["CSV", "Excel"])
        export_layout.addWidget(self.export_format)
        
        # Export button
        self.exportButton = QPushButton("Export...")
        self.exportButton.clicked.connect(self.exportData)
        export_layout.addWidget(self.exportButton)
        export_layout.addStretch()
        
        export_group.setLayout(export_layout)
        data_layout.addWidget(export_group)

        # Summary tab
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)
        
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text)

        # Add tabs to tab widget
        self.tab_widget.addTab(data_tab, "Data")
        self.tab_widget.addTab(summary_tab, "Summary")

        # Style everything
        self.apply_styles()

    def apply_styles(self):
        # Use monospace font
        font = QtGui.QFont("Consolas", 9)
        self.setFont(font)
        
        self.setStyleSheet("""
            QMainWindow {
                background: white;
            }
            QGroupBox {
                font-family: Consolas;
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 1ex;
                padding: 10px;
                background: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
                color: #2c3e50;
            }
            QPushButton {
                font-family: Consolas;
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QLineEdit, QTextEdit {
                font-family: Consolas;
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
            QComboBox {
                font-family: Consolas;
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
            QLabel {
                font-family: Consolas;
            }
            QTableWidget {
                font-family: Consolas;
            }
        """)

    def browse_export_path(self):
        file_format = self.export_format.currentText()
        file_filter = "CSV Files (*.csv)" if file_format == "CSV" else "Excel Files (*.xlsx)"
        file_ext = ".csv" if file_format == "CSV" else ".xlsx"
        
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            "",
            file_filter
        )
        if path:
            if not path.endswith(file_ext):
                path += file_ext
            self.export_path.setText(path)

    def update_summary(self):
        if not self.freq_data is None and not self.s_parameters is None:
            summary = []
            summary.append("=== Data Summary ===\n")
            
            # Frequency range
            summary.append(f"Frequency Range: {min(self.freq_data):.6f} GHz to {max(self.freq_data):.6f} GHz")
            summary.append(f"Number of Frequency Points: {len(self.freq_data)}")
            
            # Parameters
            if self.parameters and len(self.parameters) > 0:
                summary.append(f"\nParameter Combinations: {len(self.parameters)}")
                param_names = list(self.parameters[0].keys())
                summary.append(f"Parameters: {', '.join(param_names)}")
            
            # Find minimum S11 in dB
            min_s11_db = float('inf')
            min_s11_freq = 0
            min_s11_params = None
            
            for run_idx, s_param in enumerate(self.s_parameters):
                for freq_idx, freq in enumerate(self.freq_data):
                    s11_db = 20 * math.log10(abs(s_param[freq_idx]))
                    if s11_db < min_s11_db:
                        min_s11_db = s11_db
                        min_s11_freq = freq
                        min_s11_params = self.parameters[run_idx] if self.parameters else None
            
            summary.append(f"\nMinimum S11: {min_s11_db:.2f} dB at {min_s11_freq:.6f} GHz")
            if min_s11_params:
                params_str = ", ".join(f"{k}={v}" for k, v in min_s11_params.items())
                summary.append(f"Parameters at minimum S11: {params_str}")
            
            self.summary_text.setText("\n".join(summary))

    def exportData(self):
        if self.full_data is None:
            QtWidgets.QMessageBox.warning(self, "Error", "No data to export")
            return

        try:
            # Get file path from user
            file_format = self.export_format.currentText()
            file_filter = "CSV Files (*.csv)" if file_format == "CSV" else "Excel Files (*.xlsx)"
            file_ext = ".csv" if file_format == "CSV" else ".xlsx"
            
            export_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Data",
                "",
                file_filter
            )
            
            if not export_path:
                return
                
            if not export_path.endswith(file_ext):
                export_path += file_ext

            # Get selected columns
            selected_columns = [col for col, checkbox in self.parameter_checkboxes.items() if checkbox.isChecked()]
            
            # Create DataFrame
            data = []
            frequencies = self.full_data['frequencies']
            s_params = self.full_data['s_params']
            
            display_mode = self.display_mode.currentText()
            
            for run_idx, s_param in enumerate(s_params):
                for freq_idx, freq in enumerate(frequencies):
                    row_data = {}
                    row_data["Frequency (GHz)"] = freq
                    
                    # Parameters
                    if self.parameters:
                        for param_name in self.parameters[0].keys():
                            row_data[param_name] = self.parameters[run_idx][param_name]
                    
                    # S11 values
                    s11 = s_param[freq_idx]
                    if display_mode == "Complex (Real + Imaginary)":
                        row_data["S11 Real"] = s11.real
                        row_data["S11 Imaginary"] = s11.imag
                    elif display_mode == "Magnitude":
                        row_data["S11 Magnitude"] = abs(s11)
                    elif display_mode == "Magnitude (dB)":
                        row_data["S11 (dB)"] = 20 * math.log10(abs(s11))
                    else:  # Magnitude and Phase
                        row_data["S11 Magnitude"] = abs(s11)
                        row_data["Phase (degrees)"] = math.degrees(math.atan2(s11.imag, s11.real))
                    
                    # Only include selected columns
                    filtered_row = {col: row_data[col] for col in selected_columns if col in row_data}
                    data.append(filtered_row)

            df = pd.DataFrame(data)
            
            # Export based on selected format
            if file_format == "CSV":
                df.to_csv(export_path, index=False)
            else:  # Excel
                df.to_excel(export_path, index=False)
            
            QtWidgets.QMessageBox.information(self, "Success", f"Data exported to {export_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")

    def load_cst_data(self):
        try:
            import cst.results  # Import here to ensure it's available
            project = cst.results.ProjectFile(self.filePathLineEdit.text(), allow_interactive=True)
            s11_path = '1D Results\S-Parameters\S1,1'
            s11 = project.get_3d().get_result_item(s11_path, 1)
            self.freq_data = np.array(s11.get_xdata())
            
            # Update frequency range inputs
            min_freq = float(np.min(self.freq_data))
            max_freq = float(np.max(self.freq_data))
            self.freq_start.setText(f"{min_freq:.6f}")
            self.freq_end.setText(f"{max_freq:.6f}")

            # Get all parameters and S11 data for each run
            runids = project.get_3d().get_run_ids(s11_path)
            self.s_parameters = []
            self.parameters = []
            
            for run in runids:
                s11_data = project.get_3d().get_result_item(s11_path, run)
                self.s_parameters.append(np.array(s11_data.get_ydata()))
                params = project.get_3d().get_parameter_combination(run)
                self.parameters.append(params)

            self.update_display()
            self.update_summary()
            QtWidgets.QMessageBox.information(self, "Success", "CST file loaded successfully")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load CST file: {str(e)}")

    def toggle_freq_range(self, state):
        self.freq_start.setEnabled(not state)
        self.freq_end.setEnabled(not state)
        self.apply_freq_range.setEnabled(not state)
        self.update_display()

    def create_header_checkboxes(self, columns):
        # Clear existing header layout
        while self.header_layout.count():
            item = self.header_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        self.parameter_checkboxes = {}
        
        # Create a widget for each column with better styling
        for col in columns:
            container = QWidget()
            container_layout = QHBoxLayout()
            container_layout.setContentsMargins(5, 2, 5, 2)
            container_layout.setSpacing(5)
            
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            checkbox.setStyleSheet("""
                QCheckBox {
                    spacing: 5px;
                }
                QCheckBox::indicator {
                    width: 15px;
                    height: 15px;
                }
                QCheckBox::indicator:unchecked {
                    border: 2px solid #999;
                    background: white;
                }
                QCheckBox::indicator:checked {
                    border: 2px solid #4CAF50;
                    background: #4CAF50;
                }
            """)
            
            label = QLabel(col)
            label.setStyleSheet("""
                QLabel {
                    color: #2c3e50;
                    font-weight: bold;
                    padding: 2px;
                    background: #f8f9fa;
                    border-radius: 3px;
                }
            """)
            
            container_layout.addWidget(checkbox)
            container_layout.addWidget(label)
            container_layout.addStretch()
            
            container.setLayout(container_layout)
            container.setStyleSheet("""
                QWidget {
                    background: #f8f9fa;
                    border-bottom: 1px solid #dee2e6;
                }
            """)
            
            self.header_layout.addWidget(container)
            self.parameter_checkboxes[col] = checkbox
        
        # Style the table
        self.tableWidget.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                selection-background-color: #e3f2fd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: black;
            }
            QHeaderView::section {
                background-color: #f8f9fa;
                padding: 5px;
                border: 1px solid #dee2e6;
                font-weight: bold;
                color: #2c3e50;
            }
        """)
        
        # Hide default header and set alternating row colors
        self.tableWidget.horizontalHeader().hide()
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setStyleSheet(self.tableWidget.styleSheet() + """
            QTableWidget {
                alternate-background-color: #f8f9fa;
                background-color: white;
            }
        """)
        
        # Set fixed column widths
        header_width = 150  # Fixed width for each column
        for i in range(len(columns)):
            self.tableWidget.setColumnWidth(i, header_width)

    def get_freq_range(self):
        try:
            if self.use_all_freq.isChecked():
                return min(self.freq_data), max(self.freq_data)
            else:
                start = float(self.freq_start.text().replace(',', '.'))
                end = float(self.freq_end.text().replace(',', '.'))
                return start, end
        except:
            return min(self.freq_data), max(self.freq_data)

    def update_display(self):
        if self.freq_data is None or self.s_parameters is None:
            return

        # Get frequency range
        freq_start, freq_end = self.get_freq_range()
        
        # Convert frequency data to numpy array for comparison
        freq_array = np.array(self.freq_data)
        
        # Filter frequencies based on range selection
        freq_mask = (freq_array >= freq_start) & (freq_array <= freq_end)
        frequencies = freq_array[freq_mask]
        s_params = [s_param[freq_mask] for s_param in self.s_parameters]

        # Store full dataset
        self.full_data = {
            'frequencies': frequencies,
            's_params': s_params
        }

        # Limit display data
        display_step = max(1, len(frequencies) // self.MAX_DISPLAY_ROWS)
        frequencies = frequencies[::display_step]
        s_params = [s_param[::display_step] for s_param in s_params]

        # Create headers based on display mode and parameters
        param_headers = []
        if self.parameters and len(self.parameters) > 0:
            param_headers = list(self.parameters[0].keys())

        display_mode = self.display_mode.currentText()
        if display_mode == "Complex (Real + Imaginary)":
            value_headers = ["S11 Real", "S11 Imaginary"]
        elif display_mode == "Magnitude":
            value_headers = ["S11 Magnitude"]
        elif display_mode == "Magnitude (dB)":
            value_headers = ["S11 (dB)"]
        else:  # Magnitude and Phase
            value_headers = ["S11 Magnitude", "Phase (degrees)"]

        headers = ["Frequency (GHz)"] + param_headers + value_headers
        self.create_header_checkboxes(headers)

        # Prepare table
        self.tableWidget.setRowCount(len(frequencies) * len(s_params))
        self.tableWidget.setColumnCount(len(headers))

        # Fill table with data
        row = 0
        for run_idx, s_param in enumerate(s_params):
            for freq_idx, freq in enumerate(frequencies):
                col = 0
                # Frequency
                self.tableWidget.setItem(row, col, QTableWidgetItem(f"{freq:.6f}"))
                col += 1

                # Parameters
                if self.parameters:
                    for param_name in param_headers:
                        param_value = self.parameters[run_idx][param_name]
                        self.tableWidget.setItem(row, col, QTableWidgetItem(str(param_value)))
                        col += 1

                # S11 values based on display mode
                s11 = s_param[freq_idx]
                if display_mode == "Complex (Real + Imaginary)":
                    self.tableWidget.setItem(row, col, QTableWidgetItem(f"{s11.real:.6f}"))
                    self.tableWidget.setItem(row, col + 1, QTableWidgetItem(f"{s11.imag:.6f}"))
                elif display_mode == "Magnitude":
                    magnitude = abs(s11)
                    self.tableWidget.setItem(row, col, QTableWidgetItem(f"{magnitude:.6f}"))
                elif display_mode == "Magnitude (dB)":
                    magnitude_db = 20 * math.log10(abs(s11))
                    self.tableWidget.setItem(row, col, QTableWidgetItem(f"{magnitude_db:.6f}"))
                else:  # Magnitude and Phase
                    magnitude = abs(s11)
                    phase = math.degrees(math.atan2(s11.imag, s11.real))
                    self.tableWidget.setItem(row, col, QTableWidgetItem(f"{magnitude:.6f}"))
                    self.tableWidget.setItem(row, col + 1, QTableWidgetItem(f"{phase:.6f}"))
                row += 1

    def browseFile(self):
        filePath, _ = QFileDialog.getOpenFileName(self, "Select CST File", "", "CST Files (*.cst);;All Files (*)")
        if filePath:
            self.filePathLineEdit.setText(filePath)
            self.load_cst_data()

    def change_library_path(self):
        if setup_cst_path(self):
            QtWidgets.QMessageBox.information(self, "Success", 
                "CST Library path updated successfully.\nPlease restart the application for changes to take effect.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    if setup_cst_path():
        mainWin = CSTExportApp()
        mainWin.show()
        sys.exit(app.exec_())
    else:
        sys.exit(1)
