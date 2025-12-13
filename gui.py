from PyQt6.QtGui import QIcon
import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QComboBox, QTableView, QHeaderView, QMessageBox)
from PyQt6.QtCore import Qt, QAbstractTableModel
from PyQt6.QtGui import QPalette, QColor

import main as logic

class DataFrameModel(QAbstractTableModel):
    def __init__(self, data):
        super(DataFrameModel, self).__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Equipment Distributor")
        self.resize(800, 600)
        self.setWindowIcon(QIcon("logo.ico"))
        
        self.equipment_file = "equipment.xlsx"
        self.school_file = "care.json"
        self.school_models = {}
        self.result_df = None

        self.setup_ui()
        self.set_dark_theme()
        
        # Try to load default files if they exist
        if os.path.exists(self.school_file):
            self.load_school_models(self.school_file)
            
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # File Selection Area
        file_layout = QHBoxLayout()
        
        self.equip_btn = QPushButton(f"Equip DB: {self.equipment_file}")
        self.equip_btn.clicked.connect(self.select_equipment_file)
        file_layout.addWidget(self.equip_btn)
        
        self.school_btn = QPushButton(f"School Models: {self.school_file}")
        self.school_btn.clicked.connect(self.select_school_file)
        file_layout.addWidget(self.school_btn)
        
        layout.addLayout(file_layout)
        
        # Model Selection Area
        model_layout = QHBoxLayout()
        model_layout.addWidget(QLabel("Select Model:"))
        self.model_combo = QComboBox()
        model_layout.addWidget(self.model_combo)
        
        self.calc_btn = QPushButton("Calculate")
        self.calc_btn.clicked.connect(self.calculate)
        model_layout.addWidget(self.calc_btn)
        
        self.save_btn = QPushButton("Save Results")
        self.save_btn.clicked.connect(self.save_results)
        self.save_btn.setEnabled(False) # Disabled until calculation
        model_layout.addWidget(self.save_btn)
        
        layout.addLayout(model_layout)
        
        # Result Table
        self.table_view = QTableView()
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_view)
        
    def set_dark_theme(self):
        app = QApplication.instance()
        app.setStyle("Fusion")
        
        
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
        
        app.setPalette(palette)

    def select_equipment_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Equipment DB", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_name:
            self.equipment_file = file_name
            self.equip_btn.setText(f"Equip DB: {os.path.basename(file_name)}")

    def select_school_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select School Models", "", "JSON Files (*.json);;All Files (*)")
        if file_name:
            self.school_file = file_name
            self.school_btn.setText(f"School Models: {os.path.basename(file_name)}")
            self.load_school_models(file_name)

    def load_school_models(self, filepath):
        try:
            self.school_models = logic.get_school_models(filepath)
            self.model_combo.clear()
            self.model_combo.addItems(list(self.school_models.keys()))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load school models: {e}")

    def calculate(self):
        if not self.equipment_file or not os.path.exists(self.equipment_file):
            QMessageBox.warning(self, "Warning", "Please select a valid Equipment DB file.")
            return

        selected_model = self.model_combo.currentText()
        if not selected_model:
            QMessageBox.warning(self, "Warning", "Please select a model.")
            return

        try:
            model_data = logic.choose_model(self.school_models, selected_model)
            rooms = model_data['rooms']
            
            df = logic.load_equipment_db(self.equipment_file)
            final_df = logic.calculate_needed_equipment(df, rooms)
            
            self.result_df = final_df
            model = DataFrameModel(final_df)
            self.table_view.setModel(model)
            self.save_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Calculation failed: {e}")

    def save_results(self):
        if self.result_df is None or self.result_df.empty:
            QMessageBox.warning(self, "Warning", "No results to save.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Save Results", "results.xlsx", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if file_name:
            try:
                if file_name.endswith('.csv'):
                    self.result_df.to_csv(file_name, index=False)
                else:
                    self.result_df.to_excel(file_name, index=False)
                QMessageBox.information(self, "Success", f"Results saved to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
