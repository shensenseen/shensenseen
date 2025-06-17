import sys
import zlib
from PIL import Image
import io
import pyodbc
import pandas as pd
from datetime import datetime
import re
from typing import Dict, Callable, Optional, Any
import json

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit,
    QPushButton, QComboBox, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QMessageBox, QStatusBar, QListWidget, QListWidgetItem, QFileDialog,
    QCheckBox, QGridLayout, QTabWidget, QGroupBox, QHeaderView, QTabBar,
    QSpinBox, QSizePolicy, QColorDialog, QMenuBar, QTextEdit,
    QDateEdit, QScrollArea, QFrame
)
from PySide6.QtGui import QIcon, QPixmap, QFont, QPainter, QBrush, QColor, QAction
from PySide6.QtCore import Qt, Signal, QDate

# معلومات الاتصال بقاعدة البيانات والبيانات الثابتة
connection_string = 'DRIVER={SQL Server};SERVER=10.20.20.206;DATABASE=sohag;UID=mas;PWD=m@s45Scwwq'
image_connection_string = 'DRIVER={SQL Server};SERVER=10.20.20.213;DATABASE=sohag_images;UID=mas;PWD=m@s45Scwwq'

branch_names = {
    1: "فرع طما", 2: "فرع طهطا", 3: "فرع المراغة", 4: "فرع جهينة",
    5: "فرع غرب", 6: "فرع شرق", 7: "فرع اخميم", 8: "فرع ساقلتة",
    9: "فرع المنشاة", 10: "فرع البلينا", 11: "فرع جرجا", 12: "فرع دار السلام"
}
branches = branch_names.copy()

def get_db_connection():
    try:
        return pyodbc.connect(connection_string, timeout=5)
    except Exception as e:
        QMessageBox.critical(None, "خطأ اتصال", f"فشل الاتصال بقاعدة البيانات الرئيسية:\n{e}")
        return None

def get_image_db_connection():
    try:
        return pyodbc.connect(image_connection_string, timeout=5)
    except Exception as e:
        QMessageBox.critical(None, "خطأ اتصال", f"فشل الاتصال بقاعدة بيانات الصور:\n{e}")
        return None

# الأنماط (Stylesheet)
MAIN_STYLESHEET = """
    QMainWindow, QDialog { background-color: #ecf0f1; }
    QTabWidget::pane { border-top: 2px solid #3498db; background: white; }
    QTabWidget::tab-bar { alignment: right; }
    QTabBar::tab { background: #bdc3c7; padding: 12px 25px; margin-left: 1px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
    QTabBar::tab:selected { background: #3498db; color: white; }
    QPushButton { background-color: #3498db; color: white; border: none; padding: 10px 18px; border-radius: 5px; }
    QPushButton:hover { background-color: #2980b9; }
    QPushButton[objectName="ExportButton"] { background-color: #27ae60; }
    QPushButton[objectName="ExportButton"]:hover { background-color: #229954; }
    QPushButton[objectName="ClearButton"], QPushButton[objectName="ResetButton"] { background-color: #e74c3c; }
    QPushButton[objectName="ClearButton"]:hover, QPushButton[objectName="ResetButton"]:hover { background-color: #c0392b; }
    QLineEdit, QComboBox, QSpinBox, QDateEdit { padding: 8px; border: 1px solid #bdc3c7; border-radius: 5px; }
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDateEdit:focus { border: 2px solid #3498db; }
    QGroupBox { border: 1px solid #bdc3c7; border-radius: 6px; margin-top: 10px; padding: 15px; }
    QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top right; padding: 5px 10px; background-color: #f2f2f2; border-radius: 4px; }
    QTableWidget { border: 1px solid #bdc3c7; gridline-color: #dde1e2; alternate-background-color: #f8f9f9; }
    QTableWidget::item { padding: 8px; text-align: right; }
    QHeaderView::section { background-color: #34495e; color: white; padding: 8px; border: 1px solid #4a6278; }
    QStatusBar { background-color: #2c3e50; color: white; font-weight: bold; }
    QLabel#TabTitleLabel { font-size: 18pt; font-weight: bold; margin-bottom: 10px; }
    QScrollArea { background: transparent; border: none; }
"""

class WelcomeTab(QWidget):
    settingsChanged = Signal(dict)

    def __init__(self, icons):
        super().__init__()
        self.icons = icons
        self.settings_file = "settings.json"
        self.default_settings = {
            'family': 'Segoe UI', 'size': 11, 'bold': False, 'italic': False, 'color': '#2c3e50'
        }
        self.current_settings = self._load_settings()
        self.setup_ui()
        self._update_ui_from_settings()

    def _load_settings(self) -> dict:
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                for key, value in self.default_settings.items():
                    if key not in settings:
                        settings[key] = value
                return settings
        except (FileNotFoundError, json.JSONDecodeError):
            return self.default_settings.copy()

    def _save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.current_settings, f, indent=4)
        except Exception as e:
            print(f"Could not save settings: {e}")

    def _update_ui_from_settings(self):
        self.font_combo.blockSignals(True)
        self.font_size_spinbox.blockSignals(True)
        self.bold_checkbox.blockSignals(True)
        self.italic_checkbox.blockSignals(True)
        self.font_combo.setCurrentText(self.current_settings.get('family', 'Segoe UI'))
        self.font_size_spinbox.setValue(self.current_settings.get('size', 11))
        self.bold_checkbox.setChecked(self.current_settings.get('bold', False))
        self.italic_checkbox.setChecked(self.current_settings.get('italic', False))
        self._update_color_preview(self.current_settings.get('color', '#2c3e50'))
        self.font_combo.blockSignals(False)
        self.font_size_spinbox.blockSignals(False)
        self.bold_checkbox.blockSignals(False)
        self.italic_checkbox.blockSignals(False)

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(20)
        top_section_layout = QHBoxLayout()
        top_section_layout.setSpacing(30)
        welcome_section = QWidget()
        layout = QVBoxLayout(welcome_section)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(25)
        logo_label = QLabel(self)
        logo_pixmap = self.icons.get("water_icon", QIcon()).pixmap(150, 150)
        logo_label.setPixmap(logo_pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)
        title_label = QLabel("نظام إدارة صور العدادات", self)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(title_label)
        version_label = QLabel("تــم بــحــول اللـــه و قـوتــه <br> ثم بأعداد / وائل احمد عبد الفتاح   <br>  ادارة التطبيقات الالكترونية<br>  الإصدار 3.4 | شركة مياه الشرب والصرف الصحي بسوهاج", self)
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_label.setStyleSheet("font-size: 14pt; color: #550000;")
        layout.addWidget(version_label)
        top_section_layout.addWidget(welcome_section, 2)

        settings_container = QWidget()
        settings_v_layout = QVBoxLayout(settings_container)
        settings_v_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        settings_box = QGroupBox("إعدادات العرض العام", self)
        settings_box.setFixedWidth(400)
        settings_layout = QGridLayout(settings_box)
        settings_layout.setSpacing(15)

        settings_layout.addWidget(QLabel("نوع الخط:", settings_box), 0, 0)
        self.font_combo = QComboBox(settings_box)
        self.font_combo.addItems(['Segoe UI', 'Tahoma', 'Arial', 'Times New Roman', 'Calibri'])
        self.font_combo.currentTextChanged.connect(self._on_setting_change)
        settings_layout.addWidget(self.font_combo, 0, 1, 1, 2)
        settings_layout.addWidget(QLabel("حجم الخط:", settings_box), 1, 0)
        self.font_size_spinbox = QSpinBox(settings_box)
        self.font_size_spinbox.setRange(9, 20)
        self.font_size_spinbox.setSuffix(" نقطة")
        self.font_size_spinbox.valueChanged.connect(self._on_setting_change)
        settings_layout.addWidget(self.font_size_spinbox, 1, 1, 1, 2)
        style_layout = QHBoxLayout()
        self.bold_checkbox = QCheckBox("خط عريض (Bold)", settings_box)
        self.bold_checkbox.stateChanged.connect(self._on_setting_change)
        self.italic_checkbox = QCheckBox("خط مائل (Italic)", settings_box)
        self.italic_checkbox.stateChanged.connect(self._on_setting_change)
        style_layout.addWidget(self.bold_checkbox)
        style_layout.addWidget(self.italic_checkbox)
        settings_layout.addLayout(style_layout, 2, 0, 1, 3)
        settings_layout.addWidget(QLabel("لون الخط:", settings_box), 3, 0)
        self.color_button = QPushButton("اختيار اللون", settings_box)
        self.color_button.clicked.connect(self._choose_color)
        settings_layout.addWidget(self.color_button, 3, 1)
        self.color_preview_label = QLabel(self)
        self.color_preview_label.setFixedSize(40, 40)
        self.color_preview_label.setAutoFillBackground(True)
        self.color_preview_label.setStyleSheet("border: 1px solid #bdc3c7; border-radius: 20px;")
        settings_layout.addWidget(self.color_preview_label, 3, 2)
        reset_button = QPushButton("إعادة للوضع الافتراضي", self)
        reset_button.setObjectName("ResetButton")
        reset_button.clicked.connect(self.reset_to_defaults)
        settings_layout.addWidget(reset_button, 4, 0, 1, 3)

        settings_v_layout.addWidget(settings_box)
        top_section_layout.addWidget(settings_container, 1)
        main_layout.addLayout(top_section_layout)
        about_box = QGroupBox("عن البرنامج", self)
        about_layout = QVBoxLayout(about_box)
        about_text = QTextEdit(self)
        about_text.setReadOnly(True)
        about_text.setHtml(
            "<p style='text-align: right; line-height: 1.6;'><b>نظام إدارة صور العدادات</b> هو تطبيق مكتبي متكامل تم تطويره لتسهيل عمليات متابعة وتحليل صور عدادات المياه التي يتم التقاطها بواسطة المحصلين في مختلف فروع الشركة.</p>"
            "<h4 style='text-align: right; color: #2980b9;'>الأهداف الرئيسية:</h4>"
            "<ul style='text-align: right; margin-right: 20px;'><li>توفير واجهة مركزية للبحث عن بيانات المشتركين وصور العدادات.</li>"
            "<li>إنشاء تقارير تحليلية دقيقة حول أداء المحصلين ونسب التصوير.</li>"
            "<li>تسهيل عملية تصدير البيانات والصور للاستخدام الخارجي.</li>"
            "<li>تحسين الرقابة على عمليات القراءة الميدانية وضمان جودتها.</li></ul>"
            "<h4 style='text-align: right; color: #2980b9;'>المميزات:</h4>"
            "<ul style='text-align: right; margin-right: 20px;'><li>واجهة مستخدم مرنة وقابلة للتخصيص (حجم ولون الخط).</li>"
            "<li>نظام تبويبات ديناميكي لإدارة مساحات العمل.</li>"
            "<li>اتصال مباشر وآمن بقواعد بيانات الشركة.</li>"
            "<li>تقارير شاملة مع إمكانية التصدير إلى Excel.</li></ul>"
        )
        about_layout.addWidget(about_text)
        main_layout.addWidget(about_box)
        main_layout.addStretch(1)

    def _choose_color(self):
        initial_color = QColor(self.current_settings.get('color', '#2c3e50'))
        color = QColorDialog.getColor(initial_color, self, "اختر لون الخط")
        if color.isValid():
            self.current_settings['color'] = color.name()
            self._update_color_preview(color.name())
            self._on_setting_change()

    def _update_color_preview(self, color_hex: str):
        self.color_preview_label.setStyleSheet(f"background-color: {color_hex}; border: 1px solid #bdc3c7; border-radius: 20px;")

    def _on_setting_change(self):
        self.current_settings['family'] = self.font_combo.currentText()
        self.current_settings['size'] = self.font_size_spinbox.value()
        self.current_settings['bold'] = self.bold_checkbox.isChecked()
        self.current_settings['italic'] = self.italic_checkbox.isChecked()
        self.settingsChanged.emit(self.current_settings)
        self._save_settings()

    def reset_to_defaults(self):
        self.current_settings = self.default_settings.copy()
        self._update_ui_from_settings()
        self._on_setting_change()

class SearchTab(QWidget):
    def __init__(self, main_window, icons):
        super().__init__()
        self.main_window = main_window
        self.icons = icons
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)
        title_label = QLabel("البحث عن المشتركين", self)
        title_label.setObjectName("TabTitleLabel")
        layout.addWidget(title_label)
        search_box = QGroupBox("إعدادات البحث", self)
        search_layout = QGridLayout(search_box)
        search_layout.setSpacing(15)
        search_layout.addWidget(QLabel("رقم الدورة:", search_box), 0, 0)
        cycle_layout = QHBoxLayout()
        cycle_layout.setContentsMargins(0, 0, 0, 0)
        cycle_layout.setSpacing(10)
        self.year_spinbox = QSpinBox(search_box)
        self.year_spinbox.setRange(2024, 2040)
        self.year_spinbox.setValue(datetime.now().year)
        self.year_spinbox.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.year_spinbox.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.month_combo = QComboBox(search_box)
        [self.month_combo.addItem(f"{m:02}") for m in range(1, 13)]
        self.month_combo.setCurrentIndex(datetime.now().month - 1)
        cycle_layout.addWidget(self.year_spinbox)
        cycle_layout.addWidget(self.month_combo)
        search_layout.addLayout(cycle_layout, 0, 1)
        search_layout.addWidget(QLabel("اختر الفرع:", search_box), 1, 0)
        self.branch_combo = QComboBox(search_box)
        self.branch_combo.addItem("كل الفروع", 0)
        [self.branch_combo.addItem(n, k) for k, n in branch_names.items()]
        search_layout.addWidget(self.branch_combo, 1, 1)
        search_layout.addWidget(QLabel("رقم المشترك (أو جزء منه):", search_box), 2, 0)
        self.entry_custkey = QLineEdit(search_box)
        self.entry_custkey.setPlaceholderText("أدخل رقم المشترك")
        search_layout.addWidget(self.entry_custkey, 2, 1)
        query_button = QPushButton("تنفيذ الاستعلام", search_box)
        query_button.setObjectName("QueryButton")
        query_button.clicked.connect(self.execute_query)
        btn_container_layout = QHBoxLayout()
        btn_container_layout.addWidget(query_button)
        btn_container_layout.addStretch(1)
        search_layout.addLayout(btn_container_layout, 3, 0, 1, 2)
        layout.addWidget(search_box)

        results_box = QGroupBox("نتائج البحث", self)
        results_box.setMinimumHeight(900)
        results_layout = QVBoxLayout(results_box)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["الفرع", "رقم الدورة", "رقم الحساب", "اسم العميل"])
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setAlternatingRowColors(True)

        self.scroll_area.setWidget(self.results_table)
        results_layout.addWidget(self.scroll_area)

        export_btn = QPushButton("تصدير إلى Excel", results_box)
        export_btn.setObjectName("ExportButton")
        export_btn.clicked.connect(self.export_results)
        export_btn_layout = QHBoxLayout()
        export_btn_layout.addWidget(export_btn)
        export_btn_layout.addStretch(1)
        results_layout.addLayout(export_btn_layout)
        
        # --- التعديل هنا ---
        layout.addWidget(results_box)
        layout.addStretch(1)

    def execute_query(self):
        try:
            custkey = self.entry_custkey.text()
            cycle_id = int(f"{self.year_spinbox.value()}{self.month_combo.currentText()}")
            branch_id = self.branch_combo.currentData()
            if not custkey.strip():
                QMessageBox.warning(self, "خطأ إدخال", "يرجى إدخال رقم المشترك.")
                return
            conn = get_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                query = "SELECT s.NAME, h.CYCLE_ID, h.CUSTKEY, h.tent_name FROM hand_mh_St h JOIN STATIONS s ON h.STATION_NO = s.STATION_NO WHERE h.CYCLE_ID = ? AND h.CUSTKEY LIKE ?"
                params = [cycle_id, f'%{custkey}%']
                if branch_id > 0:
                    query += " AND h.STATION_NO = ?"
                    params.append(branch_id)
                query += " ORDER BY h.CUSTKEY"
                cursor.execute(query, params)
                results = cursor.fetchall()
                self.display_results(results)
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ: {e}")

    def display_results(self, results):
        self.results_table.setRowCount(0)
        if results:
            self.results_table.setRowCount(len(results))
            for r, row in enumerate(results):
                for c, cell in enumerate(row):
                    self.results_table.setItem(r, c, QTableWidgetItem(str(cell)))
            self.main_window.statusBar().showMessage(f"تم العثور على {len(results)} نتيجة.", 5000)
        else:
            QMessageBox.information(self, "نتائج البحث", "لم يتم العثور على بيانات مطابقة.")

    def export_results(self):
        if self.results_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات لتصديرها.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "حفظ الملف", "بحث_المشتركين.xlsx", "Excel Files (*.xlsx)")
        if path:
            headers = [self.results_table.horizontalHeaderItem(i).text() for i in range(self.results_table.columnCount())]
            data = [[self.results_table.item(r, c).text() for c in range(self.results_table.columnCount())] for r in range(self.results_table.rowCount())]
            pd.DataFrame(data, columns=headers).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "تم", "تم التصدير بنجاح")

class ImageTab(QWidget):
    def __init__(self, icons):
        super().__init__()
        self.icons = icons
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)
        title_label = QLabel("معاينة وتصدير الصور", self)
        title_label.setObjectName("TabTitleLabel")
        layout.addWidget(title_label)
        input_box = QGroupBox("بيانات البحث", self)
        input_layout = QGridLayout(input_box)
        input_layout.setSpacing(15)
        input_layout.addWidget(QLabel("رقم الحساب:", input_box), 0, 0)
        self.custkey_entry = QLineEdit(input_box)
        self.custkey_entry.setPlaceholderText("أدخل رقم الحساب الكامل")
        input_layout.addWidget(self.custkey_entry, 0, 1)
        btn_layout = QHBoxLayout()
        execute_btn = QPushButton("بحث عن الصور", input_box)
        execute_btn.setObjectName("ExecuteButton")
        execute_btn.clicked.connect(self.execute_query)
        btn_layout.addWidget(execute_btn)
        clear_btn = QPushButton("مسح المدخلات", input_box)
        clear_btn.setObjectName("ClearButton")
        clear_btn.clicked.connect(self.clear_entries)
        btn_layout.addWidget(clear_btn)
        export_btn = QPushButton("تصدير الصور المحددة", input_box)
        export_btn.setObjectName("ExportButton")
        export_btn.clicked.connect(self.export_images)
        btn_layout.addWidget(export_btn)
        btn_layout.addStretch(1)
        input_layout.addLayout(btn_layout, 1, 0, 1, 2)
        layout.addWidget(input_box)

        results_box = QGroupBox("الصور المتاحة", self)
        results_layout = QVBoxLayout(results_box)
        self.select_all_checkbox = QCheckBox("تحديد الكل", results_box)
        self.select_all_checkbox.stateChanged.connect(lambda s: [self.results_list.item(i).setCheckState(Qt.CheckState.Checked if s else Qt.CheckState.Unchecked) for i in range(self.results_list.count())])
        results_layout.addWidget(self.select_all_checkbox, alignment=Qt.AlignmentFlag.AlignRight)
        self.results_list = QListWidget(results_box)
        self.results_list.itemDoubleClicked.connect(self.show_selected_image)
        results_layout.addWidget(self.results_list)

        # --- التعديل هنا ---
        layout.addWidget(results_box)
        layout.addStretch(1)

    def execute_query(self):
        custkey = self.custkey_entry.text().strip()
        if not custkey:
            QMessageBox.warning(self, "خطأ", "يرجى إدخال رقم الحساب.")
            return
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT CUSTKEY, NAME FROM METER_IMAGES WHERE CUSTKEY = ? ORDER BY STAMP_DATE DESC", [custkey])
                results = cursor.fetchall()
            self.results_list.clear()
            if not results:
                QMessageBox.information(self, "نتيجة", "لم يتم العثور على صور لهذا الحساب.")
                return
            for res in results:
                item = QListWidgetItem(f"حساب: {res.CUSTKEY} | اسم الصورة: {res.NAME or 'N/A'}")
                item.setData(Qt.ItemDataRole.UserRole, res.NAME)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked)
                self.results_list.addItem(item)
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ: {e}")

    def export_images(self):
        selected_names = [self.results_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.results_list.count()) if self.results_list.item(i).checkState() == Qt.CheckState.Checked]
        if not selected_names:
            QMessageBox.warning(self, "تنبيه", "يرجى تحديد صورة واحدة على الأقل.")
            return
        folder = QFileDialog.getExistingDirectory(self, "اختر مجلد الحفظ")
        if not folder:
            return
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                exported_count = 0
                for name in selected_names:
                    cursor.execute("SELECT IMAGE FROM METER_IMAGES WHERE NAME = ?", [name])
                    res = cursor.fetchone()
                    if res and res.IMAGE:
                        Image.open(io.BytesIO(zlib.decompress(res.IMAGE))).save(f"{folder}/{name}.png")
                        exported_count += 1
            QMessageBox.information(self, "نجاح", f"تم تصدير {exported_count} صورة بنجاح.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل التصدير: {e}")

    def clear_entries(self):
        self.custkey_entry.clear()
        self.results_list.clear()

    def show_selected_image(self, item):
        name = item.data(Qt.ItemDataRole.UserRole)
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                cursor.execute("SELECT IMAGE FROM METER_IMAGES WHERE NAME = ?", [name])
                res = cursor.fetchone()
            if res and res.IMAGE:
                Image.open(io.BytesIO(zlib.decompress(res.IMAGE))).show(title=name)
            else:
                QMessageBox.warning(self, "خطأ", "لم يتم العثور على بيانات الصورة.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"لا يمكن عرض الصورة: {e}")

class ImageSearchApp(QWidget):
    def __init__(self, icons):
        super().__init__()
        self.icons = icons
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)
        title_label = QLabel("تقرير الصور برقم البلوك", self)
        title_label.setObjectName("TabTitleLabel")
        layout.addWidget(title_label)
        search_box = QGroupBox("إعدادات البحث", self)
        search_layout = QGridLayout(search_box)
        search_layout.setSpacing(15)
        search_layout.addWidget(QLabel("اختيار الفرع:", search_box), 0, 0)
        self.branch_combo = QComboBox(search_box)
        [self.branch_combo.addItem(v, k) for k, v in branches.items()]
        search_layout.addWidget(self.branch_combo, 0, 1)
        search_layout.addWidget(QLabel("اختار رقم الدورة:", search_box), 1, 0)
        cycle_layout = QHBoxLayout()
        cycle_layout.setContentsMargins(0, 0, 0, 0)
        self.year_spinbox = QSpinBox(search_box)
        self.year_spinbox.setRange(2024, 2040)
        self.year_spinbox.setValue(datetime.now().year)
        self.month_combo = QComboBox(search_box)
        [self.month_combo.addItem(f"{m:02}") for m in range(1, 13)]
        self.month_combo.setCurrentIndex(datetime.now().month - 1)
        cycle_layout.addWidget(self.year_spinbox)
        cycle_layout.addWidget(self.month_combo)
        search_layout.addLayout(cycle_layout, 1, 1)
        search_layout.addWidget(QLabel("رقم البلوك (واحد على الأقل):", search_box), 2, 0)
        walk_layout = QHBoxLayout()
        self.walk_inputs = [QLineEdit(search_box) for _ in range(6)]
        [ip.setPlaceholderText("بلوك") or ip.setMaximumWidth(110) or walk_layout.addWidget(ip) for ip in self.walk_inputs]
        search_layout.addLayout(walk_layout, 2, 1)
        btn_layout = QHBoxLayout()
        self.search_btn = QPushButton("بحث", search_box)
        self.search_btn.setObjectName("SearchButton")
        self.search_btn.clicked.connect(self.search)
        btn_layout.addWidget(self.search_btn)
        self.export_btn = QPushButton("تصدير إلى Excel", search_box)
        self.export_btn.setObjectName("ExportButton")
        self.export_btn.clicked.connect(self.export_to_excel)
        btn_layout.addWidget(self.export_btn)
        btn_layout.addStretch(1)
        search_layout.addLayout(btn_layout, 3, 0, 1, 2)
        layout.addWidget(search_box)

        results_box = QGroupBox("نتائج البحث", self)
        results_box.setMinimumHeight(900)
        results_layout = QVBoxLayout(results_box)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["رقم الحساب", "الدفعة", "وقت التصوير"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)

        self.scroll_area.setWidget(self.table)
        results_layout.addWidget(self.scroll_area)
        
        # --- التعديل هنا ---
        layout.addWidget(results_box)
        layout.addStretch(1)

    def search(self):
        cycle_id = f"{self.year_spinbox.value()}{self.month_combo.currentText()}"
        station_no = self.branch_combo.currentData()
        walk_nos = [w.text() for w in self.walk_inputs if w.text().strip().isdigit()]
        if not walk_nos:
            QMessageBox.warning(self, "تحذير", "يرجى إدخال رقم بلوك واحد على الأقل.")
            return
        placeholders = ','.join(['?'] * len(walk_nos))
        query = f"SELECT custkey, cycle_id, STAMP_DATE FROM METER_IMAGES WHERE STATION_NO = ? AND CYCLE_ID = ? AND WALK_NO IN ({placeholders}) ORDER BY STAMP_DATE DESC"
        params = [station_no, cycle_id] + [int(w) for w in walk_nos]
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                self.rows = [list(r) for r in cursor.fetchall()]
                self.columns = ["رقم الحساب", "الدفعة", "وقت التصوير"]
                self.table.setRowCount(len(self.rows))
                self.table.setColumnCount(len(self.columns))
                self.table.setHorizontalHeaderLabels(self.columns)
                for r_idx, row_data in enumerate(self.rows):
                    for c_idx, val in enumerate(row_data):
                        self.table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))
            if not self.rows:
                QMessageBox.information(self, "نتيجة البحث", "لم يتم العثور على بيانات مطابقة.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"خطأ في البحث: {e}")

    def export_to_excel(self):
        if not hasattr(self, 'rows') or not self.rows:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات لتصديرها")
            return
        cycle_id = f"{self.year_spinbox.value()}{self.month_combo.currentText()}"
        path, _ = QFileDialog.getSaveFileName(self, "حفظ الملف", f"تقرير_صور_البلوك_{self.branch_combo.currentText()}_{cycle_id}.xlsx", "Excel Files (*.xlsx)")
        if path:
            pd.DataFrame(self.rows, columns=self.columns).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "تم", "تم التصدير بنجاح")

class CountImagesApp(QWidget):
    def __init__(self, icons):
        super().__init__()
        self.icons = icons
        self.results = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(20)
        title_label = QLabel("تقرير إجمالي الصور بالفروع", self)
        title_label.setObjectName("TabTitleLabel")
        layout.addWidget(title_label)
        search_box = QGroupBox("إعدادات البحث", self)
        search_layout = QGridLayout(search_box)
        search_layout.setSpacing(15)
        search_layout.addWidget(QLabel("اختار الفرع:", search_box), 0, 0)
        self.branch_combo = QComboBox(search_box)
        self.branch_combo.addItem("كل الفروع", 0)
        [self.branch_combo.addItem(v, k) for k, v in branches.items()]
        search_layout.addWidget(self.branch_combo, 0, 1, 1, 2)
        search_layout.addWidget(QLabel("نوع البحث:", search_box), 1, 0)
        self.search_type_combo = QComboBox(search_box)
        self.search_type_combo.addItem("بحث بالدورة", "cycle")
        self.search_type_combo.addItem("بحث بنطاق التاريخ", "date")
        self.search_type_combo.currentIndexChanged.connect(self.search_type_changed)
        search_layout.addWidget(self.search_type_combo, 1, 1, 1, 2)
        self.cycle_search_widget = QWidget()
        cycle_layout = QHBoxLayout(self.cycle_search_widget)
        cycle_layout.setContentsMargins(0, 0, 0, 0)
        self.year_spinbox = QSpinBox(search_box)
        self.year_spinbox.setRange(2024, 2040)
        self.year_spinbox.setValue(datetime.now().year)
        self.month_combo = QComboBox(search_box)
        [self.month_combo.addItem(f"{m:02}") for m in range(1, 13)]
        self.month_combo.setCurrentIndex(datetime.now().month - 1)
        cycle_layout.addWidget(QLabel("اختر الدورة:"))
        cycle_layout.addWidget(self.year_spinbox)
        cycle_layout.addWidget(self.month_combo)
        search_layout.addWidget(self.cycle_search_widget, 2, 0, 1, 3)
        self.date_range_search_widget = QWidget()
        date_range_layout = QGridLayout(self.date_range_search_widget)
        date_range_layout.setContentsMargins(0, 0, 0, 0)
        self.start_date_edit = QDateEdit(self, calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))
        self.end_date_edit = QDateEdit(self, calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        date_range_layout.addWidget(QLabel("من تاريخ:"), 0, 0)
        date_range_layout.addWidget(self.start_date_edit, 0, 1)
        date_range_layout.addWidget(QLabel("إلى تاريخ:"), 1, 0)
        date_range_layout.addWidget(self.end_date_edit, 1, 1)
        search_layout.addWidget(self.date_range_search_widget, 3, 0, 1, 3)
        btn_layout = QHBoxLayout()
        self.search_btn = QPushButton("بحث", search_box)
        self.search_btn.clicked.connect(self.search)
        btn_layout.addWidget(self.search_btn)
        self.export_btn = QPushButton("تصدير إلى Excel", search_box)
        self.export_btn.clicked.connect(self.export_to_excel)
        btn_layout.addWidget(self.export_btn)
        btn_layout.addStretch(1)
        search_layout.addLayout(btn_layout, 4, 0, 1, 3)
        layout.addWidget(search_box)

        results_box = QGroupBox("نتائج البحث", self)
        results_box.setMinimumHeight(900)
        results_layout = QVBoxLayout(results_box)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.scroll_area.setWidget(self.table)
        results_layout.addWidget(self.scroll_area)
        
        # --- التعديل هنا ---
        layout.addWidget(results_box)
        layout.addStretch(1)
        
        self.search_type_changed()

    def search_type_changed(self):
        search_type = self.search_type_combo.currentData()
        self.cycle_search_widget.setVisible(search_type == "cycle")
        self.date_range_search_widget.setVisible(search_type != "cycle")
        self.table.setRowCount(0)
        if search_type == 'date':
            self.table.setColumnCount(3)
            self.table.setHorizontalHeaderLabels(["اسم الفرع", "الشهر", "عدد الصور"])
        else:
            self.table.setColumnCount(2)
            self.table.setHorizontalHeaderLabels(["اسم الفرع", "عدد الصور"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def search(self):
        station_no = self.branch_combo.currentData()
        search_type = self.search_type_combo.currentData()
        params, conditions = [], []
        self.search_type_changed()
        if search_type == 'cycle':
            cycle_id = f"{self.year_spinbox.value()}{self.month_combo.currentText()}"
            conditions.append("CYCLE_ID=?")
            params.append(cycle_id)
            query_template = "SELECT STATION_NO, COUNT(*) FROM METER_IMAGES {where_clause} GROUP BY STATION_NO ORDER BY STATION_NO"
        else:
            start_date_str = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd")
            conditions.append("CONVERT(date, STAMP_DATE) BETWEEN ? AND ?")
            params.extend([start_date_str, end_date_str])
            query_template = "SELECT STATION_NO, FORMAT(STAMP_DATE, 'yyyy-MM') AS ImageMonth, COUNT(*) FROM METER_IMAGES {where_clause} GROUP BY STATION_NO, FORMAT(STAMP_DATE, 'yyyy-MM') ORDER BY STATION_NO, ImageMonth"
        if station_no != 0:
            conditions.append("STATION_NO=?")
            params.append(station_no)
        where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""
        query = query_template.format(where_clause=where_clause)
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                results = cursor.fetchall()
                self.table.setRowCount(0)
                self.table.setRowCount(len(results))
                for r, row in enumerate(results):
                    self.table.setItem(r, 0, QTableWidgetItem(branch_names.get(row[0], f"فرع {row[0]}")))
                    if search_type == 'date':
                        self.table.setItem(r, 1, QTableWidgetItem(str(row[1])))
                        self.table.setItem(r, 2, QTableWidgetItem(str(row[2])))
                    else:
                        self.table.setItem(r, 1, QTableWidgetItem(str(row[1])))
                if not results:
                    QMessageBox.information(self, "نتيجة البحث", "لم يتم العثور على بيانات.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في البحث", f"حدث خطأ: {e}")

    def export_to_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات لتصديرها.")
            return

        branch_name_for_file = self.branch_combo.currentText().replace(" ", "_")
        search_type = self.search_type_combo.currentData()
        if search_type == 'cycle':
            filename = f"تقرير_الصور_{branch_name_for_file}_دفعة_{self.year_spinbox.value()}{self.month_combo.currentText()}.xlsx"
        else:
            filename = f"تقرير_الصور_{branch_name_for_file}_من_{self.start_date_edit.date().toString('yyyy-MM-dd')}_الى_{self.end_date_edit.date().toString('yyyy-MM-dd')}.xlsx"

        path, _ = QFileDialog.getSaveFileName(self, "حفظ ملف Excel", filename, "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            data = []
            for row in range(self.table.rowCount()):
                row_data = [self.table.item(row, col).text() for col in range(self.table.columnCount())]
                data.append(row_data)

            df = pd.DataFrame(data, columns=headers)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "تم بنجاح", "تم تصدير التقرير بنجاح.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", f"فشل التصدير:\n{e}")

class EmployeeImageApp(QWidget):
    def __init__(self, icons):
        super().__init__()
        self.icons = icons
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)
        title_label = QLabel("تقرير أداء المحصلين", self)
        title_label.setObjectName("TabTitleLabel")
        layout.addWidget(title_label)
        search_box = QGroupBox("إعدادات البحث", self)
        search_layout = QGridLayout(search_box)
        search_layout.setSpacing(15)
        search_layout.addWidget(QLabel("اختيار الفرع:", search_box), 0, 0)
        self.branch_combo = QComboBox(search_box)
        self.branch_combo.addItem("كل الفروع", 0)
        [self.branch_combo.addItem(v, k) for k, v in branches.items()]
        search_layout.addWidget(self.branch_combo, 0, 1)
        search_layout.addWidget(QLabel("بحث بكود الموظف:", search_box), 1, 0)
        self.emp_id_input = QLineEdit(search_box)
        self.emp_id_input.setPlaceholderText("اترك فارغًا للكل")
        search_layout.addWidget(self.emp_id_input, 1, 1)
        search_layout.addWidget(QLabel("فلتر العرض:", search_box), 2, 0)
        self.image_filter_combo = QComboBox(search_box)
        self.image_filter_combo.addItem("عرض الكل", "all")
        self.image_filter_combo.addItem("من التقطوا صورًا فقط", "has_images")
        self.image_filter_combo.addItem("من لم يلتقطوا صورًا فقط", "no_images")
        search_layout.addWidget(self.image_filter_combo, 2, 1)
        search_layout.addWidget(QLabel("نوع البحث:", search_box), 3, 0)
        self.search_type_combo = QComboBox(search_box)
        self.search_type_combo.addItem("بحث بالدورة", "cycle")
        self.search_type_combo.addItem("بحث بنطاق التاريخ", "date")
        self.search_type_combo.currentIndexChanged.connect(self.search_type_changed)
        search_layout.addWidget(self.search_type_combo, 3, 1)
        self.cycle_search_widget = QWidget()
        cycle_layout = QHBoxLayout(self.cycle_search_widget)
        cycle_layout.setContentsMargins(0, 0, 0, 0)
        self.year_spinbox = QSpinBox(search_box)
        self.year_spinbox.setRange(2024, 2040)
        self.year_spinbox.setValue(datetime.now().year)
        self.month_combo = QComboBox(search_box)
        [self.month_combo.addItem(f"{m:02}") for m in range(1, 13)]
        self.month_combo.setCurrentIndex(datetime.now().month - 1)
        cycle_layout.addWidget(QLabel("اختر الدورة:"))
        cycle_layout.addWidget(self.year_spinbox)
        cycle_layout.addWidget(self.month_combo)
        search_layout.addWidget(self.cycle_search_widget, 4, 0, 1, 2)
        self.date_range_search_widget = QWidget()
        date_range_layout = QGridLayout(self.date_range_search_widget)
        date_range_layout.setContentsMargins(0, 0, 0, 0)
        self.start_date_edit = QDateEdit(self, calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))
        self.end_date_edit = QDateEdit(self, calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        date_range_layout.addWidget(QLabel("من تاريخ:"), 0, 0)
        date_range_layout.addWidget(self.start_date_edit, 0, 1)
        date_range_layout.addWidget(QLabel("إلى تاريخ:"), 1, 0)
        date_range_layout.addWidget(self.end_date_edit, 1, 1)
        search_layout.addWidget(self.date_range_search_widget, 5, 0, 1, 2)
        btn_layout = QHBoxLayout()
        self.search_btn = QPushButton("بحث", search_box)
        self.search_btn.clicked.connect(self.search)
        btn_layout.addWidget(self.search_btn)
        self.export_btn = QPushButton("تصدير إلى Excel", search_box)
        self.export_btn.clicked.connect(self.export_to_excel)
        btn_layout.addWidget(self.export_btn)
        btn_layout.addStretch(1)
        search_layout.addLayout(btn_layout, 6, 0, 1, 2)
        layout.addWidget(search_box)

        results_box = QGroupBox("نتائج البحث", self)
        results_box.setMinimumHeight(900)
        results_layout = QVBoxLayout(results_box)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.scroll_area.setWidget(self.table)
        results_layout.addWidget(self.scroll_area)
        
        # --- التعديل هنا ---
        layout.addWidget(results_box)
        layout.addStretch(1)
        
        self.search_type_changed()

    def search_type_changed(self):
        search_type = self.search_type_combo.currentData()
        self.cycle_search_widget.setVisible(search_type == "cycle")
        self.date_range_search_widget.setVisible(search_type != "cycle")
        self.table.setRowCount(0)
        if search_type == 'date':
            self.table.setColumnCount(5)
            self.table.setHorizontalHeaderLabels(["اسم الفرع", "الشهر", "كود الموظف", "اسم الموظف", "عدد الصور"])
        else:
            self.table.setColumnCount(4)
            self.table.setHorizontalHeaderLabels(["اسم الفرع", "كود الموظف", "اسم الموظف", "عدد الصور"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def search(self):
        self.search_type_changed()
        branch_id = self.branch_combo.currentData()
        emp_id = self.emp_id_input.text().strip()
        filter_option = self.image_filter_combo.currentData()
        search_type = self.search_type_combo.currentData()
        params, where_conditions, group_by_columns, select_columns, order_by_clause = [], [], [], [], ""
        base_join = "ON m.DEVICE_ID=e.DEVICE_ID AND m.STATION_NO=e.BRANCH_ID"
        base_group = "GROUP BY e.BRANCH_ID, e.id, e.FULL_NAME"
        base_order = "ORDER BY e.BRANCH_ID, e.FULL_NAME, IMAGE_COUNT DESC"
        select_columns = ["e.BRANCH_ID", "e.id", "e.FULL_NAME"]
        group_by_columns = ["e.BRANCH_ID", "e.id", "e.FULL_NAME"]
        if search_type == 'cycle':
            where_conditions.append("m.CYCLE_ID=?")
            params.append(f"{self.year_spinbox.value()}{self.month_combo.currentText()}")
        else:
            where_conditions.append("CONVERT(date, m.STAMP_DATE) BETWEEN ? AND ?")
            start_date_str = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd")
            params.extend([start_date_str, end_date_str])
            select_columns.insert(1, "FORMAT(m.STAMP_DATE, 'yyyy-MM') AS ImageMonth")
            group_by_columns.append("FORMAT(m.STAMP_DATE, 'yyyy-MM')")
            order_by_clause = "ORDER BY e.BRANCH_ID, e.FULL_NAME, ImageMonth"
        if branch_id != 0:
            where_conditions.append("e.BRANCH_ID=?")
            params.append(branch_id)
        if emp_id:
            where_conditions.append("e.id=?")
            params.append(emp_id)
        select_clause = "SELECT " + ", ".join(select_columns) + ", COUNT(m.CUSTKEY) as IMAGE_COUNT"
        from_clause = "FROM [emp].[dbo].[EMP] e LEFT JOIN [sohag_images].[dbo].[METER_IMAGES] m " + base_join
        where_clause = "WHERE " + " AND ".join(where_conditions)
        group_clause = "GROUP BY " + ", ".join(group_by_columns)
        having_clause = ""
        if filter_option == "has_images":
            having_clause = " HAVING COUNT(m.CUSTKEY)>0"
        elif filter_option == "no_images":
            having_clause = " HAVING COUNT(m.CUSTKEY)=0"
        final_order = order_by_clause if order_by_clause else base_order
        base_query = f"{select_clause} {from_clause} {where_clause} {group_clause} {having_clause} {final_order}"
        try:
            conn = get_image_db_connection()
            if not conn:
                return
            with conn:
                cursor = conn.cursor()
                cursor.execute(base_query, params)
                results = cursor.fetchall()
                self.table.setRowCount(0)
                self.table.setRowCount(len(results))
                for r, row in enumerate(results):
                    branch_name = branch_names.get(row[0], f"غير معروف {row[0]}")
                    if search_type == 'date':
                        self.table.setItem(r, 0, QTableWidgetItem(branch_name))
                        self.table.setItem(r, 1, QTableWidgetItem(str(row[1])))
                        self.table.setItem(r, 2, QTableWidgetItem(str(row[2])))
                        self.table.setItem(r, 3, QTableWidgetItem(str(row[3])))
                        self.table.setItem(r, 4, QTableWidgetItem(str(row[4])))
                    else:
                        self.table.setItem(r, 0, QTableWidgetItem(branch_name))
                        self.table.setItem(r, 1, QTableWidgetItem(str(row[1])))
                        self.table.setItem(r, 2, QTableWidgetItem(str(row[2])))
                        self.table.setItem(r, 3, QTableWidgetItem(str(row[3])))
                if not results:
                    QMessageBox.information(self, "نتيجة البحث", "لم يتم العثور على بيانات مطابقة.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"خطأ في الاستعلام: {e}\n\n{base_query}")

    def export_to_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير.")
            return

        search_type = self.search_type_combo.currentData()
        if search_type == 'cycle':
            filename = f"تقرير_المحصلين_دفعة_{self.year_spinbox.value()}{self.month_combo.currentText()}.xlsx"
        else:
            filename = f"تقرير_المحصلين_من_{self.start_date_edit.date().toString('yyyy-MM-dd')}_الى_{self.end_date_edit.date().toString('yyyy-MM-dd')}.xlsx"

        path, _ = QFileDialog.getSaveFileName(self, "حفظ", filename, "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            data = []
            for row in range(self.table.rowCount()):
                row_data = [self.table.item(row, col).text() for col in range(self.table.columnCount())]
                data.append(row_data)

            df = pd.DataFrame(data, columns=headers)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "تم", "تم التصدير بنجاح")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", f"فشل تصدير البيانات إلى Excel:\n{e}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.original_stylesheet = MAIN_STYLESHEET
        self.tab_accent_colors = ["#3498db", "#2ecc71", "#f1c40f", "#e67e22", "#9b59b6", "#1abc9c"]
        self.open_tabs = {}
        self.tab_actions = {}
        self.tab_metadata = {}
        self.setup_icons()
        self.setup_window()
        self.setup_menu()
        self.setup_tabs()
        self.setup_status_bar()
        if hasattr(self, 'welcome_tab'):
            self.apply_global_font_settings(self.welcome_tab.current_settings)

    def setup_icons(self):
        self.icons = {"app_icon": self.create_icon_glyph("#3498db", "🏢", size=64), "water_icon": self.create_icon_glyph("#3498db", "💧", size=64)}
        self.setWindowIcon(self.icons.get("app_icon", QIcon()))

    def create_icon_glyph(self, bg_color_hex, text_symbol, size=32):
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if bg_color_hex and bg_color_hex != "transparent":
            painter.setBrush(QBrush(QColor(bg_color_hex)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(0, 0, size, size)
        font = QFont('Arial', int(size * 0.5))
        painter.setFont(font)
        text_col = QColor("#ffffff" if bg_color_hex != "transparent" else "#333333")
        painter.setPen(text_col)
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, text_symbol)
        painter.end()
        return QIcon(pixmap)

    def setup_window(self):
        self.setWindowTitle("نظام إدارة صور العدادات - نسخة 3.4")
        self.resize(800, 600)
        self.setMinimumSize(640, 480)

    def setup_menu(self):
        menu_bar = self.menuBar()
        tabs_menu = menu_bar.addMenu("التبويبات")
        self.tab_metadata = {
            'welcome': {'title': "الرئيسية", 'factory': None},
            'search': {'title': "بحث مشتركين", 'factory': lambda: SearchTab(self, self.icons)},
            'image_view': {'title': "معاينة صور", 'factory': lambda: ImageTab(self.icons)},
            'block_search': {'title': "صور البلوك", 'factory': lambda: ImageSearchApp(self.icons)},
            'branch_total': {'title': "إجمالي صور الفرع", 'factory': lambda: CountImagesApp(self.icons)},
            'employee_report': {'title': "تقرير المحصلين", 'factory': lambda: EmployeeImageApp(self.icons)},
        }
        for key, meta in self.tab_metadata.items():
            if meta['factory'] is None:
                continue
            action = QAction(meta['title'], self)
            action.triggered.connect(lambda checked=False, k=key: self.open_tab_by_key(k))
            tabs_menu.addAction(action)
            self.tab_actions[key] = action

    def setup_tabs(self):
        self.tabs = QTabWidget(self)
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(True)
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.on_tab_close_requested)
        self.setCentralWidget(self.tabs)

        self.welcome_tab = WelcomeTab(self.icons)
        self.welcome_scroll_area = QScrollArea()
        self.welcome_scroll_area.setWidget(self.welcome_tab)
        self.welcome_scroll_area.setWidgetResizable(True)
        self.welcome_scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        welcome_index = self.tabs.addTab(self.welcome_scroll_area, "الرئيسية")
        self.tabs.tabBar().setTabButton(welcome_index, QTabBar.ButtonPosition.RightSide, None)
        self.welcome_tab.settingsChanged.connect(self.apply_global_font_settings)

        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.open_tab_by_key('search')
        self.open_tab_by_key('employee_report')

    def open_tab_by_key(self, key: str):
        if key in self.open_tabs:
            self.tabs.setCurrentWidget(self.open_tabs[key])
            return
        if key in self.tab_metadata:
            meta = self.tab_metadata[key]
            if meta['factory'] is None:
                self.tabs.setCurrentWidget(self.welcome_scroll_area)
                return

            content_widget = meta['factory']()
            scroll_area = QScrollArea()
            scroll_area.setWidget(content_widget)
            scroll_area.setWidgetResizable(True)
            scroll_area.setFrameShape(QFrame.Shape.NoFrame)

            index = self.tabs.addTab(scroll_area, meta['title'])
            self.open_tabs[key] = scroll_area
            self.tab_actions[key].setEnabled(False)
            self.tabs.setCurrentIndex(index)

    def on_tab_close_requested(self, index: int):
        widget_to_close = self.tabs.widget(index)
        if widget_to_close == self.welcome_scroll_area:
            return

        key_to_remove = next((k for k, v in self.open_tabs.items() if v == widget_to_close), None)
        if key_to_remove:
            self.tabs.removeTab(index)
            del self.open_tabs[key_to_remove]
            self.tab_actions[key_to_remove].setEnabled(True)

    def apply_global_font_settings(self, settings: dict):
        font_family = settings.get('family', 'Segoe UI')
        font_size = settings.get('size', 11)
        font_weight = 'bold' if settings.get('bold', False) else 'normal'
        font_style = 'italic' if settings.get('italic', False) else 'normal'
        font_color = settings.get('color', '#2c3e50')
        tab_font_size = font_size + 2
        dynamic_stylesheet = f"""
            QWidget {{ font-family: "{font_family}"; font-size: {font_size}pt; font-weight: {font_weight}; font-style: {font_style}; color: {font_color}; direction: rtl; }}
            QPushButton, QGroupBox::title {{ font-weight: bold; font-style: normal; }}
            QTabBar::tab {{ font-size: {tab_font_size}pt; font-style: normal; }}"""
        final_stylesheet = dynamic_stylesheet + self.original_stylesheet
        QApplication.instance().setStyleSheet(final_stylesheet)
        self.on_tab_changed(self.tabs.currentIndex())

    def on_tab_changed(self, index: int):
        self.apply_tab_colors()
        if index < 0 or self.tabs.count() == 0:
            return
        tab_bar = self.tabs.tabBar()
        accent_color = self.tab_accent_colors[index % len(self.tab_accent_colors)]
        tab_bar.setProperty("selectedTabAccentColor", QColor(accent_color))
        tab_bar.style().unpolish(tab_bar)
        tab_bar.style().polish(tab_bar)

    def apply_tab_colors(self):
        tab_bar = self.tabs.tabBar()
        base_color = QColor(self.welcome_tab.current_settings.get('color', '#2c3e50'))
        for i in range(tab_bar.count()):
            tab_bar.setTabTextColor(i, base_color)
        current_index = self.tabs.currentIndex()
        if current_index >= 0:
            accent_color = self.tab_accent_colors[current_index % len(self.tab_accent_colors)]
            tab_bar.setTabTextColor(current_index, QColor(accent_color))

    def setup_status_bar(self):
        status_bar = self.statusBar()
        date_label = QLabel(f"تاريخ اليوم: {datetime.now().strftime('%A, %d %B %Y')}", status_bar)
        status_bar.addPermanentWidget(date_label)
        status_bar.showMessage("جاهز", 5000)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("نظام إدارة صور العدادات")
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())