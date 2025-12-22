
import sys
import math
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QComboBox, QSpinBox, QScrollArea, 
                             QTableView, QHeaderView, QMessageBox, QFrame, QDialog, 
                             QLineEdit, QFileDialog)
from PyQt6.QtCore import Qt, QAbstractTableModel
from PyQt6.QtGui import QPalette, QColor, QFont

import main as logic

# Данные о нагрузке по классам (часы в неделю)
GRADE_DATA = {}
for g in range(1, 12):
    d = {"universal": 0, "foreign_lang": 0, "biology_safety": 0, "physics": 0, "chemistry": 0, "it_material": 0, "informatics": 0}
    if g <= 4:
        d["universal"] = 5 + 4 + 4 + 2 + 2 
        if g >= 2: d["foreign_lang"] = 2 * 2
    elif g <= 9:
        rus = {5:5, 6:6, 7:4, 8:3, 9:3}[g]
        lit = {5:3, 6:3, 7:2, 8:2, 9:3}[g]
        mat = 5
        hist = {5:3, 6:3, 7:3, 8:3, 9:2}[g]
        soc = {5:0, 6:0, 7:0, 8:1, 9:1}[g]
        geo = {5:1, 6:1, 7:2, 8:2, 9:2}[g]
        pe = 2
        d["universal"] = rus + lit + mat + hist + soc + geo + pe
        d["foreign_lang"] = 3 * 2
        d["biology_safety"] = {5:1, 6:1, 7:2, 8:2, 9:2}[g] + {5:0, 6:0, 7:0, 8:1, 9:1}[g]
        d["physics"] = {5:0, 6:0, 7:2, 8:2, 9:3}[g]
        d["chemistry"] = {5:0, 6:0, 7:0, 8:2, 9:2}[g]
        d["informatics"] = {5:0, 6:0, 7:1, 8:1, 9:1}[g] * 2
        d["it_material"] = {5:2, 6:2, 7:2, 8:1, 9:0}[g] * 2
    else:
        d["universal"] = 2 + 3 + 6 + 2 + 2 + 2
        d["foreign_lang"] = 3 * 2
        d["biology_safety"] = 1
        d["physics"] = 2
        d["chemistry"] = 1
        d["informatics"] = 1 * 2
    GRADE_DATA[g] = d

ROOM_CAPACITY_WEEKLY = 34
AVG_CLASS_SIZE = 25

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setFixedWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Equipment File
        layout.addWidget(QLabel("Файл оборудования (Excel):"))
        file_layout = QHBoxLayout()
        self.file_edit = QLineEdit(settings['equipment_file'])
        file_layout.addWidget(self.file_edit)
        
        browse_btn = QPushButton("Обзор")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)

        # Capacity settings
        cap_layout = QHBoxLayout()
        cap_layout.addWidget(QLabel("Часов в кабинете (нед):"))
        self.capacity_spin = QSpinBox()
        self.capacity_spin.setRange(1, 100)
        self.capacity_spin.setValue(settings['room_capacity'])
        cap_layout.addWidget(self.capacity_spin)
        
        cap_layout.addWidget(QLabel("Учеников в классе:"))
        self.avg_size_spin = QSpinBox()
        self.avg_size_spin.setRange(1, 50)
        self.avg_size_spin.setValue(settings['avg_class_size'])
        cap_layout.addWidget(self.avg_size_spin)
        layout.addLayout(cap_layout)
        
        btns_layout = QHBoxLayout()
        save_btn = QPushButton("ОК")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        btns_layout.addWidget(save_btn)
        btns_layout.addWidget(cancel_btn)
        layout.addLayout(btns_layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл оборудования", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.file_edit.setText(file_path)

    def get_settings(self):
        return {
            'equipment_file': self.file_edit.text(),
            'room_capacity': self.capacity_spin.value(),
            'avg_class_size': self.avg_size_spin.value()
        }

class DataFrameModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole:
                val = self._data.iloc[index.row(), index.column()]
                if isinstance(val, (int, float)) and val % 1 == 0:
                    return str(int(val))
                return str(val)
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("School Infrastructure & Equipment Calculator")
        self.resize(1000, 800)
        
        self.equipment_file = "equipment.xlsx"
        self.room_capacity = 34
        self.avg_class_size = 25
        self.grade_data = GRADE_DATA
        self.result_df = None
        
        self.programs = {
            "1-4": range(1, 5),
            "1-6": range(1, 7),
            "5-11": range(5, 12),
            "7-11": range(7, 12),
            "1-9": range(1, 10),
            "1-11": range(1, 12)
        }
        
        self.parallels_widgets = {} # {grade: QSpinBox}
        self.setup_ui()
        self.set_dark_theme()
        self.update_grade_inputs()
        self.recalculate_free_rooms()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- Top Panel ---
        top_panel = QFrame()
        top_layout = QHBoxLayout(top_panel)
        
        # Program Selection
        top_layout.addWidget(QLabel("Программа:"))
        self.program_combo = QComboBox()
        self.program_combo.addItems(list(self.programs.keys()))
        self.program_combo.currentTextChanged.connect(self.update_grade_inputs)
        top_layout.addWidget(self.program_combo)
        
        # Cabinets Count
        top_layout.addWidget(QLabel("Всего кабинетов:"))
        self.total_rooms_spin = QSpinBox()
        self.total_rooms_spin.setRange(1, 200)
        self.total_rooms_spin.setValue(35)
        self.total_rooms_spin.valueChanged.connect(self.recalculate_free_rooms)
        top_layout.addWidget(self.total_rooms_spin)
        
        # Settings Button
        self.settings_btn = QPushButton("⚙ Настройки")
        self.settings_btn.clicked.connect(self.open_settings)
        top_layout.addWidget(self.settings_btn)
        
        top_layout.addStretch()
        
        # Free Rooms Indicator
        self.free_rooms_label = QLabel("Свободных мест: 0")
        self.free_rooms_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        top_layout.addWidget(self.free_rooms_label)
        
        main_layout.addWidget(top_panel)

        # --- Middle Panel (Parallels) ---
        parallels_group = QFrame()
        parallels_group.setFrameShape(QFrame.Shape.StyledPanel)
        self.parallels_layout = QHBoxLayout(parallels_group)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(parallels_group)
        scroll.setMaximumHeight(150)
        main_layout.addWidget(scroll)

        # --- Bottom Panel (Actions & Results) ---
        actions_layout = QHBoxLayout()
        self.calc_btn = QPushButton("Рассчитать оборудование")
        self.calc_btn.clicked.connect(self.calculate_equipment)
        actions_layout.addWidget(self.calc_btn)
        
        self.save_btn = QPushButton("Сохранить результаты")
        self.save_btn.clicked.connect(self.save_results)
        self.save_btn.setEnabled(False)
        actions_layout.addWidget(self.save_btn)
        
        main_layout.addLayout(actions_layout)

        self.table_view = QTableView()
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.table_view)

    def set_dark_theme(self):
        app = QApplication.instance()
        app.setStyle("Fusion")
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        app.setPalette(palette)

    def open_settings(self):
        settings = {
            'equipment_file': self.equipment_file,
            'room_capacity': self.room_capacity,
            'avg_class_size': self.avg_class_size
        }
        dlg = SettingsDialog(settings, self)
        if dlg.exec():
            new_settings = dlg.get_settings()
            self.equipment_file = new_settings['equipment_file']
            self.room_capacity = new_settings['room_capacity']
            self.avg_class_size = new_settings['avg_class_size']
            self.recalculate_free_rooms()

    def update_grade_inputs(self):
        # Clear existing
        for i in reversed(range(self.parallels_layout.count())): 
            widget = self.parallels_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.parallels_widgets = {}

        selected_program = self.program_combo.currentText()
        grade_range = self.programs[selected_program]

        for g in grade_range:
            v_layout = QVBoxLayout()
            v_layout.addWidget(QLabel(f"{g} класс"), alignment=Qt.AlignmentFlag.AlignCenter)
            
            spin = QSpinBox()
            spin.setRange(0, 100)
            spin.setValue(1)
            spin.valueChanged.connect(self.recalculate_free_rooms)
            v_layout.addWidget(spin)
            
            w = QWidget()
            w.setLayout(v_layout)
            self.parallels_layout.addWidget(w)
            self.parallels_widgets[g] = spin
        
        self.recalculate_free_rooms()

    def recalculate_free_rooms(self):
        total_available_rooms = self.total_rooms_spin.value()
        
        # Считаем часы по категориям
        final_hours = {k: 0 for k in self.grade_data[1].keys()}
        total_classes = 0
        for g, spin in self.parallels_widgets.items():
            count = spin.value()
            total_classes += count
            for cat, hrs in self.grade_data[g].items():
                final_hours[cat] += hrs * count
        
        # Считаем необходимые комнаты
        needed_rooms = sum(math.ceil(hrs / self.room_capacity) for hrs in final_hours.values())
        
        # Правка для it_info (как в test.py)
        selected_program = self.program_combo.currentText()
        if selected_program in ["5-11", "7-11", "1-9", "1-11"]:
            it_info_needed = math.ceil(final_hours["informatics"] / self.room_capacity) + \
                             math.ceil(final_hours["it_material"] / self.room_capacity)
            if it_info_needed == 0:
                needed_rooms += 1

        # Свободные кабинеты
        free_rooms = total_available_rooms - needed_rooms
        
        # Расчет свободных ученических мест
        # Логика: сколько полных классов (по 25 чел) можно еще впихнуть в свободные кабинеты.
        # Для одной полной параллели (1 класс в каждом году обучения) нужно parallel_rooms_needed кабинетов.
        # Но проще: свободные кабинеты / (среднее количество кабинетов на 1 класс).
        
        # Считаем "вес" одного класса в кабинетах для текущей программы
        grade_range = self.programs[selected_program]
        total_hrs_per_parallel = {k: 0 for k in self.grade_data[1].keys()}
        for g in grade_range:
            for cat, hrs in self.grade_data[g].items():
                total_hrs_per_parallel[cat] += hrs
        
        rooms_per_parallel = sum(math.ceil(h / self.room_capacity) for h in total_hrs_per_parallel.values())
        rooms_per_class = rooms_per_parallel / len(grade_range) if grade_range else 1
        
        # Свободные ученические места = (свободные кабинеты / кабинетов на класс) * 25
        free_students = int(free_rooms / rooms_per_class * self.avg_class_size) if rooms_per_class > 0 else 0

        self.free_rooms_label.setText(f"Свободных мест (учеников): {free_students}")
        
        if free_students < 0:
            self.free_rooms_label.setStyleSheet("color: #ff4d4d;")
        else:
            self.free_rooms_label.setStyleSheet("color: white;")
        
        return final_hours

    def calculate_equipment(self):
        try:
            final_hours = self.recalculate_free_rooms()
            rooms_allocation = {
                cat: math.ceil(hrs / self.room_capacity) for cat, hrs in final_hours.items()
            }
            
            # Специальная правка для it_info как специфика модели
            selected_program = self.program_combo.currentText()
            if selected_program in ["5-11", "7-11", "1-9", "1-11"]:
                if rooms_allocation.get("informatics", 0) + rooms_allocation.get("it_material", 0) == 0:
                    rooms_allocation["informatics"] = 1

            if not os.path.exists(self.equipment_file):
                QMessageBox.warning(self, "Ошибка", f"Файл {self.equipment_file} не найден. Проверьте настройки.")
                return

            df = logic.load_equipment_db(self.equipment_file)
            final_df = logic.calculate_needed_equipment(df, rooms_allocation)
            
            self.result_df = final_df
            model = DataFrameModel(final_df)
            self.table_view.setModel(model)
            self.save_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при расчете: {e}")

    def save_results(self):
        if self.result_df is None or self.result_df.empty:
            QMessageBox.warning(self, "Предупреждение", "Нет результатов для сохранения.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить результаты", "results.xlsx", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if file_name:
            try:
                if file_name.endswith('.csv'):
                    self.result_df.to_csv(file_name, index=False)
                else:
                    self.result_df.to_excel(file_name, index=False)
                QMessageBox.information(self, "Успех", f"Результаты сохранены в {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
