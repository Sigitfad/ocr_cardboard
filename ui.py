"""
Semua komponen UI menggunakan PySide6 untuk aplikasi QC_GS-Battery.
"""

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QRadioButton, QCheckBox, QGroupBox, QSpinBox,
    QMessageBox, QFileDialog, QTreeWidget, QTreeWidgetItem, QHeaderView, QDialog,
    QComboBox, QDateEdit, QAbstractItemView, QCompleter
)  # PySide6 GUI components | UI widgets dan layouts
from PySide6.QtCore import (
    Qt, QTimer, Signal, QThread, QDateTime, QDate, QLocale, QMetaObject
)  # PySide6 core | Core signal/slot dan threading
from PySide6.QtGui import (
    QPixmap, QImage, QFont, QColor
)  # PySide6 GUI utilities | Untuk image handling dan styling
from config import (
    APP_NAME, WINDOW_WIDTH, WINDOW_HEIGHT, CONTROL_PANEL_WIDTH, RIGHT_PANEL_WIDTH,
    JIS_TYPES, DIN_TYPES, MONTHS, MONTH_MAP
)  # Import konfigurasi dari config.py
from datetime import datetime  # Date/time operations | Modul untuk date/time
from ui_export import create_export_dialog  # Import fungsi export dialog | Fungsi untuk membuat export dialog
import os  # File operations | Modul untuk file operations


class LogicSignals(QThread):
    # Class wrapper QThread | Tujuan: Wrapper untuk DetectionLogic instance dengan signal/slot capability
    
    update_signal = Signal(object)  # Signal untuk update video frame | Emit ketika frame baru siap
    code_detected_signal = Signal(str)  # Signal untuk code detected | Emit ketika kode terdeteksi
    camera_status_signal = Signal(str, bool)  # Signal untuk camera status | Emit status kamera (on/off)
    data_reset_signal = Signal()  # Signal untuk reset data | Emit untuk reset display
    all_text_signal = Signal(list)  # Signal untuk OCR text output | Emit list teks yang terdeteksi

    def __init__(self):
        # Fungsi inisialisasi QThread | Tujuan: Setup thread dan buat DetectionLogic instance
        super().__init__()
        from ocr import DetectionLogic
        self.logic = DetectionLogic(
            self.update_signal,
            self.code_detected_signal,
            self.camera_status_signal,
            self.data_reset_signal,
            self.all_text_signal
        )
        
    def run(self):
        # Fungsi jalankan thread | Tujuan: Start thread event loop
        self.exec()


class MainWindow(QMainWindow):
    # Class jendela utama aplikasi | Tujuan: Main application window dengan semua UI components
    
    export_result_signal = Signal(str)  # Signal untuk export result | Emit hasil export file
    export_status_signal = Signal(str, str)  # Signal untuk export status | Emit status export operation
    file_scan_result_signal = Signal(str)  # Signal untuk file scan result | Emit hasil scan dari file

    def __init__(self):
        # Fungsi inisialisasi main window | Tujuan: Setup UI dan inisialisasi semua components
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setGeometry(100, 100, WINDOW_WIDTH, WINDOW_HEIGHT)
        
        self.logic_thread = None
        self.logic = None
        
        self.export_result_signal.connect(self._handle_export_result)
        self.export_status_signal.connect(self._update_export_button_ui)
        self.file_scan_result_signal.connect(self._handle_file_scan_result)

        self._setup_logic_thread(initial_setup=True)
        
        self.setup_ui()
        self.setup_timer()
    
    def _setup_logic_thread(self, initial_setup=False):
        #Helper untuk membuat instance baru LogicSignals dan Logic, lalu menghubungkan sinyal.
        
        if self.logic_thread:
            if self.logic:
                 self.logic.stop_detection()
            
            if self.logic_thread.isRunning():
                 self.logic_thread.quit()
                 self.logic_thread.wait(5000)
                 try:
                     self.logic_thread.update_signal.disconnect(self.update_video_frame)
                     self.logic_thread.code_detected_signal.disconnect(self.handle_code_detection)
                     self.logic_thread.camera_status_signal.disconnect(self.update_camera_status)
                     self.logic_thread.data_reset_signal.disconnect(self.update_code_display)
                     self.logic_thread.all_text_signal.disconnect(self.update_all_text_display)
                 except TypeError:
                     pass
                     
            self.logic_thread = None
            self.logic = None
        
        self.logic_thread = LogicSignals()
        self.logic = self.logic_thread.logic
        
        self.logic_thread.update_signal.connect(self.update_video_frame)
        self.logic_thread.code_detected_signal.connect(self.handle_code_detection)
        self.logic_thread.camera_status_signal.connect(self.update_camera_status)
        self.logic_thread.data_reset_signal.connect(self.update_code_display)
        self.logic_thread.all_text_signal.connect(self.update_all_text_display)
    
    def closeEvent(self, event):
        #Menangani penutupan jendela (tombol X).
        reply = QMessageBox.question(self, 'Quit Confirmation',
                                     "Are you sure for quit?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            if self.logic:
                self.logic.stop_detection()
            if self.logic_thread and self.logic_thread.isRunning():
                self.logic_thread.quit()
                self.logic_thread.wait()
            event.accept()
        else:
            event.ignore()

    def setup_timer(self):
        #Mengatur timer untuk jam real-time dan pengecekan reset harian.
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_realtime_clock)
        self.timer.start(1000)

    def update_realtime_clock(self):
        #Update the date/time label and check for daily reset.
        now = QDateTime.currentDateTime()
        
        if self.logic and self.logic.check_daily_reset():
             QMessageBox.information(self, "Reset Data", f"Data deteksi telah di-reset untuk hari baru: {self.logic.current_date.strftime('%d-%m-%Y')}")
             
        locale = QLocale(QLocale.Indonesian, QLocale.Indonesia)
        formatted_time = locale.toString(now, "dddd, d MMMM yyyy, HH:mm:ss")
        
        self.date_time_label.setText(formatted_time)

    def setup_ui(self):
        #Fungsi setup seluruh UI aplikasi | Tujuan: Create dan arrange semua UI components di window
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        main_layout = QHBoxLayout(main_widget)
        
        control_frame = self._create_control_panel()
        main_layout.addWidget(control_frame)
        
        self.video_label = QLabel("CAMERA OFF")
        self.video_label.setAlignment(Qt.AlignCenter)
        self.video_label.setStyleSheet("background-color: black; color: white; font-size: 14pt;")
        main_layout.addWidget(self.video_label, 1)
        
        right_panel = self._create_right_panel()
        main_layout.addWidget(right_panel)

        control_frame.setFixedWidth(CONTROL_PANEL_WIDTH)
        right_panel.setFixedWidth(RIGHT_PANEL_WIDTH)

    def _create_control_panel(self):
        #Fungsi buat panel kontrol | Tujuan: Create left control panel dengan buttons dan combos
        frame = QWidget()
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 15, 15, 15)

        top_control_widget = QWidget()
        top_control_layout = QHBoxLayout(top_control_widget)
        top_control_layout.setContentsMargins(0, 0, 0, 0)
        
        preset_group = QGroupBox("Tipe:")
        preset_layout = QVBoxLayout(preset_group)
        preset_layout.setAlignment(Qt.AlignTop)
        preset_layout.setContentsMargins(8, 12, 8, 8)
        preset_layout.setSpacing(6)

        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["JIS", "DIN"])
        self.preset_combo.setCurrentIndex(0)
        
        preset_layout.addWidget(self.preset_combo)
        preset_layout.addStretch()
        
        set_options = lambda: self.logic.set_camera_options(
            self.preset_combo.currentText(),
            False,  # flip_h disabled (was: self.cb_flip_h.isChecked())
            False,  # flip_v disabled (was: self.cb_flip_v.isChecked())
            self.cb_binary.isChecked(),
            self.cb_split.isChecked(),
            2.0
        ) if self.logic else None
        
        self.preset_combo.currentTextChanged.connect(set_options)
        self.preset_combo.currentTextChanged.connect(self._update_label_options)
        top_control_layout.addWidget(preset_group)
        
        options_group = QGroupBox("Option:")
        options_layout = QVBoxLayout(options_group)
        # self.cb_flip_h = QCheckBox("HORIZONAL FLIP")
        # self.cb_flip_v = QCheckBox("VERTICAL FLIP")
        self.cb_binary = QCheckBox("BINARY COLOR")  # Checkbox untuk mode binary (grayscale threshold)
        self.cb_split = QCheckBox("SHOW SPLIT SCREEN")  # Checkbox untuk split screen mode
        
        option_change = set_options
        # self.cb_flip_h.toggled.connect(option_change)
        # self.cb_flip_v.toggled.connect(option_change)
        self.cb_binary.toggled.connect(option_change)  # Trigger set_options saat binary mode berubah
        self.cb_split.toggled.connect(option_change)  # Trigger set_options saat split mode berubah

        # options_layout.addWidget(self.cb_flip_h)
        # options_layout.addWidget(self.cb_flip_v)
        options_layout.addWidget(self.cb_binary)
        options_layout.addWidget(self.cb_split)
        top_control_layout.addWidget(options_group)
        
        layout.addWidget(top_control_widget)
        
        self.camera_label = QLabel("Camera: Not Selected")
        self.camera_label.setFont(QFont("Arial", 9))
        self.camera_label.setStyleSheet("color: blue;")
        layout.addWidget(self.camera_label)

        jis_type_label = QLabel("Select Label:")
        jis_type_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(jis_type_label)
        
        self.jis_type_combo = QComboBox()
        self.jis_type_combo.addItems(JIS_TYPES)
        
        self.jis_type_combo.setEditable(True)
        self.jis_type_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.jis_type_combo.setCompleter(QCompleter(self.jis_type_combo.model()))
        self.jis_type_combo.currentTextChanged.connect(self.on_jis_type_changed)
        layout.addWidget(self.jis_type_combo)
        
        self.selected_type_label = QLabel("No Type Selected - Pilih Label untuk memulai")
        self.selected_type_label.setFont(QFont("Arial", 9))
        self.selected_type_label.setStyleSheet("color: #FF6600; font-weight: bold;")
        self.selected_type_label.setWordWrap(True)
        layout.addWidget(self.selected_type_label)

        camera_btn_container = QWidget()
        camera_btn_layout = QHBoxLayout(camera_btn_container)
        camera_btn_layout.setContentsMargins(0, 0, 0, 0)
        camera_btn_layout.setSpacing(5)

        self.btn_start = QPushButton("START")
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 35px;")
        self.btn_start.clicked.connect(self.start_detection)
        
        self.btn_stop = QPushButton("STOP")
        self.btn_stop.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 35px;")
        self.btn_stop.clicked.connect(self.stop_detection)
        self.btn_stop.setEnabled(False)

        camera_btn_layout.addWidget(self.btn_start)
        camera_btn_layout.addWidget(self.btn_stop)
        layout.addWidget(camera_btn_container)
        
        self.btn_file = QPushButton("SCAN FROM FILE")
        self.btn_file.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")
        self.btn_file.clicked.connect(self.open_file_scan_dialog)
        layout.addWidget(self.btn_file)
        
        self.success_container = QWidget()
        self.success_layout = QVBoxLayout(self.success_container)
        self.success_container.setFixedHeight(50)
        layout.addWidget(self.success_container)
        
        all_text_group = QGroupBox("Detection Output:")
        all_text_layout = QVBoxLayout(all_text_group)
        
        self.all_text_tree = QTreeWidget()
        self.all_text_tree.setHeaderLabels(["Element Text"])
        self.all_text_tree.header().setVisible(False)
        self.all_text_tree.setStyleSheet("font-size: 9pt; background-color: #f9f9f9;")
        self.all_text_tree.setMinimumHeight(150)
        all_text_layout.addWidget(self.all_text_tree)
        
        layout.addWidget(all_text_group, 2)

        return frame

    def update_all_text_display(self, text_list):
        #Update list elemen teks yang terdeteksi.
        self.all_text_tree.clear()
        for text in text_list:
            item = QTreeWidgetItem([text])
            self.all_text_tree.addTopLevelItem(item)

    def on_jis_type_changed(self, text):
        #Handler ketika user memilih JIS Type.
        if text == "Select Label . . ." or not text.strip():
            self.selected_type_label.setText("No Type Selected - Pilih Label untuk memulai")
            self.selected_type_label.setStyleSheet("color: #FF6600; font-weight: bold;")
            if self.logic:
                self.logic.set_target_label("")  # Changed from selected_type = None to set_target_label("")
        else:
            self.selected_type_label.setText(f"Selected: {text}")
            self.selected_type_label.setStyleSheet("color: green; font-weight: bold;")
            if self.logic:
                self.logic.set_target_label(text)  # Changed from selected_type = text to set_target_label(text)
        
        self.update_code_display()

    def _create_right_panel(self):
        #Fungsi buat panel kanan | Tujuan: Create right panel dengan export button dan data display table
        frame = QWidget()
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 15, 15, 15)
        
        self.btn_export = QPushButton("EXPORT DATA")
        self.btn_export.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")
        self.btn_export.clicked.connect(self.open_export_dialog)
        layout.addWidget(self.btn_export)
        
        self.date_time_label = QLabel("Memuat Tanggal...")
        self.date_time_label.setFont(QFont("Arial", 10))
        layout.addWidget(self.date_time_label)
        
        label_barang = QLabel("Data Barang :")
        label_barang.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(label_barang)
        
        self.code_tree = QTreeWidget()
        self.code_tree.setHeaderLabels(["Waktu", "Label", "Status", "Path Gambar", "ID"])
        self.code_tree.setColumnCount(5)
        self.code_tree.header().setDefaultAlignment(Qt.AlignCenter)
        self.code_tree.setColumnWidth(0, 80)
        self.code_tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.code_tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.code_tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.code_tree.header().setSectionResizeMode(3, QHeaderView.Fixed)
        self.code_tree.header().setSectionResizeMode(4, QHeaderView.Fixed)

        self.code_tree.setSelectionMode(QAbstractItemView.MultiSelection)

        self.code_tree.setColumnHidden(3, True)
        self.code_tree.setColumnHidden(4, True)
        
        self.code_tree.itemDoubleClicked.connect(self.view_selected_image)

        layout.addWidget(self.code_tree)
        
        self.btn_delete_selected = QPushButton("CLEAR")
        self.btn_delete_selected.setStyleSheet("background-color: #Ff0000; color: white; font-weight: bold; height: 30px;")
        self.btn_delete_selected.clicked.connect(self.delete_selected_codes)
        layout.addWidget(self.btn_delete_selected)
        
        self.update_code_display()

        return frame

    def _reset_file_scan_button(self):
        #Meriset teks dan status tombol 'Scan from File'.
        self.btn_file.setText("SCAN FROM FILE")
        self.btn_file.setEnabled(True)

    def _update_camera_button_styles(self, active_button):
        #Mengatur warna tombol kamera.
        style_active = "background-color: #8B0000; color: white; font-weight: bold; height: 35px;"
        style_inactive = "background-color: #4CAF50; color: white; font-weight: bold; height: 35px;"

        if active_button == "START":
            self.btn_start.setStyleSheet(style_active)
            self.btn_stop.setStyleSheet(style_inactive)
        else:
            self.btn_stop.setStyleSheet(style_active)
            self.btn_start.setStyleSheet(style_inactive)
    
    def start_detection(self):
        #Handler untuk tombol START.
        import threading
        
        selected_type = self.jis_type_combo.currentText()
        if selected_type == "Select Label . . ." or not selected_type.strip():
            QMessageBox.warning(self, "Warning",
                "Harap pilih label terlebih dahulu sebelum memulai scanning!")
            return
            
        self._setup_logic_thread()
        
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self._update_camera_button_styles("START")
        
        if self.logic:
            self.logic.set_camera_options(
                self.preset_combo.currentText(),
                False,  # flip_h disabled
                False,  # flip_v disabled
                self.cb_binary.isChecked(),
                self.cb_split.isChecked(),
                2.0
            )
            self.logic.set_target_label(selected_type)  # Changed from set_target_jis_type to set_target_label
            
            self.logic.start_detection()
            self.logic_thread.start()
        
        self._hide_success_popup()

    def stop_detection(self):
        #Handler untuk tombol STOP.
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self._update_camera_button_styles("STOP")
        
        if self.logic:
            self.logic.stop_detection()
        
        if self.logic_thread and self.logic_thread.isRunning():
            self.logic_thread.quit()
            self.logic_thread.wait()
            
        self._hide_success_popup()

    def update_camera_status(self, status_text, is_running):
        #Update status kamera.
        self.camera_label.setText(status_text)
        if not is_running:
            self.video_label.setText("CAMERA STOP")

    def update_video_frame(self, pil_image):
        #Update frame video dari kamera.
        if not self.video_label.size().isValid():
            return
            
        qimage = QImage(pil_image.tobytes(), pil_image.width, pil_image.height,
                        pil_image.width * 3, QImage.Format_RGB888)
        
        pixmap = QPixmap.fromImage(qimage)
        scaled_pixmap = pixmap.scaled(self.video_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        
        self.video_label.setPixmap(scaled_pixmap)
        self.video_label.setText("")

    def handle_code_detection(self, detected_code):
        #Menangani sinyal kode terdeteksi dari Logic.
        self.update_code_display()
        
        if detected_code.startswith("ERROR:"):
            QMessageBox.critical(self, "Error Pemindaian File", f"Terjadi kesalahan saat pemindaian OCR/Regex:\n{detected_code[6:]}")
            self._reset_file_scan_button()
        elif detected_code == "FAILED":
            QMessageBox.critical(self, "Gagal Deteksi", "Tidak ada label yang terdeteksi pada gambar.")
            self._reset_file_scan_button()
        else:
            self.show_detection_success(detected_code)
            if self.logic and not self.logic.running:
                self._reset_file_scan_button()

    def update_code_display(self):
        #Update tampilan data kode yang terdeteksi.
        if not self.logic:
            return
            
        self.code_tree.clear()
        
        selected_session = self.jis_type_combo.currentText()
        show_nothing = (selected_session == "Select Label . . ." or not selected_session.strip())
        
        if show_nothing:
            self.selected_type_label.setText("No Type Selected - Pilih label untuk memulai")
            self.selected_type_label.setStyleSheet("color: #FF6600; font-weight: bold;")
            return
        
        displayed_count = 0
        ok_count = 0
        not_ok_count = 0
        
        for i, record in enumerate(reversed(self.logic.detected_codes)):
            target_session = record.get('TargetSession', record['Code'])
            
            if target_session != selected_session:
                continue
            
            displayed_count += 1
            
            time_str = record['Time'][11:19]
            code_str = f"{record['Code']} ({record['Type']})"
            status_str = record.get('Status', 'OK')
            image_path = record.get('ImagePath', '')
            record_id = record.get('ID', '')
            
            item = QTreeWidgetItem([time_str, code_str, status_str, image_path, str(record_id)])
            self.code_tree.addTopLevelItem(item)
            
            if status_str == "OK":
                ok_count += 1
            elif status_str == "Not OK":
                not_ok_count += 1
                # Set background merah dan text putih untuk setiap kolom
                for col in range(item.columnCount()):
                    item.setBackground(col, QColor(255, 0, 0))  # Red background
                    item.setForeground(col, QColor(255, 255, 255))  # White text
        
        status_text = f"Total: {displayed_count} | OK: {ok_count} | Not OK: {not_ok_count}"
        self.selected_type_label.setText(f"Label: {selected_session} | {status_text}")

    def view_selected_image(self, item, column):
        #Handler untuk membuka gambar double-click.
        import sys
        import subprocess
        
        try:
            image_path = item.text(3)
            
            if not image_path or image_path == 'N/A' or not os.path.exists(image_path):
                QMessageBox.warning(self, "Gambar Tidak Ditemukan",
                                    f"File gambar tidak ditemukan atau path tidak valid:\n{image_path}")
                return
            
            if sys.platform == "win32":
                os.startfile(image_path)
            elif sys.platform == "darwin":
                subprocess.call(('open', image_path))
            else:
                subprocess.call(('xdg-open', image_path))

        except Exception as e:
            QMessageBox.critical(self, "Error Membuka Gambar",
                                 f"Gagal membuka file gambar:\n{e}")

    def show_detection_success(self, detected_code):
        #Tampilkan popup sukses deteksi.
        self._hide_success_popup()

        success_widget = QWidget()
        success_widget.setStyleSheet(
            "background-color: #F70D0D; "
            "border: 2px solid #D00; "
            "border-radius: 5px;"
        )
        success_widget.setFixedHeight(42)   #  POPUP 

        layout = QVBoxLayout(success_widget)
        layout.setContentsMargins(6, 4, 6, 4)

        label = QLabel(f"SCAN BERHASIL !\n{detected_code}")
        label.setAlignment(Qt.AlignCenter)
        label.setWordWrap(True)
        label.setStyleSheet("""
            color: white;
            font-weight: bold;
            font-size: 12px;   
            line-height: 12px;
        """)

        layout.addWidget(label)

        self.success_layout.addWidget(success_widget)
        self.current_success_popup = success_widget

        QTimer.singleShot(3000, self._hide_success_popup)


    def _hide_success_popup(self):
        #Sembunyikan popup sukses.
        if hasattr(self, 'current_success_popup') and self.current_success_popup:
            self.current_success_popup.deleteLater()
            self.current_success_popup = None

    def delete_selected_codes(self):
        #Handler untuk tombol CLEAR.
        selected_items = self.code_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Harap pilih data terlebih dahulu!")
            return
        
        reply = QMessageBox.question(
            self, "Konfirmasi",
            f"Apakah Anda yakin ingin menghapus {len(selected_items)} item?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes and self.logic:
            for item in selected_items:
                record_id = int(item.text(4))
                self.logic.delete_codes(record_id)
            self.update_code_display()

    def open_file_scan_dialog(self):
        #Membuka dialog file untuk scan.
        import threading
        
        selected_type = self.jis_type_combo.currentText()
        if selected_type == "Select Label . . ." or not selected_type.strip():
            QMessageBox.warning(self, "Warning",
                "Harap pilih label terlebih dahulu sebelum memulai scanning!")
            return
            
        if self.logic and self.logic.running:
             QMessageBox.information(self, "Info", "Harap hentikan Live Detection sebelum memindai dari file.")
             return
             
        self._hide_success_popup()

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Image Files (*.png *.jpg *.jpeg *.bmp *.tiff)"
        )
        
        if file_path:
            self.btn_file.setText("SCANNING . . .")
            self.btn_file.setEnabled(False)
            threading.Thread(target=self._scan_file_thread, args=(file_path,), daemon=True).start()

    def _scan_file_thread(self, file_path):
        #Thread untuk scan image file tanpa block UI.
        if self.logic:
            self.logic.scan_file(file_path)

    def _handle_file_scan_result(self, result):
        #Handle hasil scan file dari thread.
        self.update_code_display()

    def open_export_dialog(self):
        #Fungsi: Buka dialog export data dengan berbagai filter option.
        dialog = create_export_dialog(self, self.logic, self.preset_combo, self.jis_type_combo)
        
        if not dialog:
            return
        
        # Store state for access from export functions
        export_range_var = "All"

        def handle_export_click():
            # Fungsi: Handle click export button dan validate filters
            # Tujuan: Collect filter parameters dan start export thread
            import threading
            from datetime import timedelta, time as py_time

            start_date = None
            end_date = None
            sql_filter = ""
            date_range_desc = ""

            try:
                current_time = datetime.now()
                range_key = dialog.export_range_var
                
                if range_key == "All":
                    sql_filter = ""
                    date_range_desc = "Semua Data Tersimpan"
                elif range_key == "Today":
                    start_date = datetime(current_time.year, current_time.month, current_time.day, 0, 0, 0)
                    end_date = datetime(current_time.year, current_time.month, current_time.day, 23, 59, 59)
                elif range_key == "24H":
                    start_date = current_time - timedelta(days=1)
                    end_date = current_time
                elif range_key == "7D":
                    start_date = current_time - timedelta(weeks=1)
                    end_date = current_time
                elif range_key == "1Y":
                    start_date = current_time - timedelta(days=365)
                    end_date = current_time
                elif range_key == "Month":
                    month_name = dialog.month_combo.currentText()
                    year = int(dialog.year_combo.currentText())
                    month_num = MONTH_MAP.get(month_name)
                    start_date = datetime(year, month_num, 1, 0, 0, 0)
                    if month_num == 12:
                        end_date = datetime(year + 1, 1, 1, 0, 0, 0) - timedelta(microseconds=1)
                    else:
                        end_date = datetime(year, month_num + 1, 1, 0, 0, 0) - timedelta(microseconds=1)
                elif range_key == "CustomDate":
                    selected_start_date = dialog.start_date_entry.date().toPython()
                    selected_end_date = dialog.end_date_entry.date().toPython()
                    start_date = datetime(selected_start_date.year, selected_start_date.month, selected_start_date.day, 0, 0, 0)
                    end_date = datetime(selected_end_date.year, selected_end_date.month, selected_end_date.day, 23, 59, 59)
                    if start_date > end_date:
                        raise ValueError("Tanggal Mulai tidak boleh setelah Tanggal Akhir.")

                if start_date:
                    start_date_str_db = start_date.strftime("%Y-%m-%d %H:%M:%S")
                    end_date_str_db = end_date.strftime("%Y-%m-%d %H:%M:%S")
                    start_date_str_id = start_date.strftime("%d-%m-%Y %H:%M:%S")
                    end_date_str_id = end_date.strftime("%d-%m-%Y %H:%M:%S")
                    if range_key == "Today":
                         date_range_desc = start_date.strftime('%d-%m-%Y')
                    elif range_key in ["CustomDate", "Month"] and start_date.time() == py_time.min and end_date.time() == py_time(23, 59, 59):
                         date_range_desc = f"{start_date.strftime('%d-%m-%Y')} s/d {end_date.strftime('%d-%m-%Y')}"
                    else:
                         date_range_desc = f"{start_date_str_id} s/d {end_date_str_id}"
                    sql_filter = f"WHERE timestamp BETWEEN '{start_date_str_db}' AND '{end_date_str_db}'"

                selected_export_preset = dialog.export_preset_combo.currentText()
                if selected_export_preset == "Preset":
                    selected_export_preset = self.preset_combo.currentText()
                
                if sql_filter:
                    sql_filter += f" AND preset = '{selected_export_preset}'"
                else:
                    sql_filter = f"WHERE preset = '{selected_export_preset}'"

                # Filter label
                if dialog.export_label_filter_enabled.isChecked():
                    selected_export_label = dialog.export_label_type_combo.currentText()
                    if selected_export_label and selected_export_label != "All Label":
                        if sql_filter:
                            sql_filter += f" AND target_session = '{selected_export_label}'"
                        else:
                            sql_filter = f"WHERE target_session = '{selected_export_label}'"
                        # Karena Label sudah ada di A3-B3, A1-B1 hanya untuk Date saja

                dialog.accept()

                # Start export in background thread
                threading.Thread(
                    target=self._execute_export_thread, 
                    args=(
                        sql_filter, 
                        date_range_desc, 
                        dialog.export_label_type_combo.currentText() if dialog.export_label_filter_enabled.isChecked() else "",
                        selected_export_preset
                    ), 
                    daemon=True
                ).start()

            except Exception as e:
                QMessageBox.critical(self, "Error Filter", f"Gagal menentukan rentang waktu:\n{e}")
                dialog.reject()

        export_handler = lambda: handle_export_click()
        dialog.export_btn.clicked.connect(export_handler)
        
        self.btn_export.setEnabled(False)
        try:
            dialog.exec()
        finally:
            try:
                dialog.export_btn.clicked.disconnect(export_handler)
            except RuntimeError:
                pass
            self.btn_export.setEnabled(True)

    def _execute_export_thread(self, sql_filter, date_range_desc, export_label="", current_preset=""):
        #Thread untuk proses export data ke Excel.
        from export import execute_export
    
        self.export_status_signal.emit("Exporting...", "#FF9800")
        
        if not self.logic:
             self.export_result_signal.emit("EXPORT_ERROR: Logic Object not found")
             return
        
        result = execute_export(sql_filter, date_range_desc, export_label, current_preset)
        
        self.export_result_signal.emit(result)

    def _handle_export_result(self, result):
        #"""Handle hasil export dan tampilkan feedback kepada user.
        self.btn_export.setText("EXPORT DATA")
        self.btn_export.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")
        
        if result == "NO_DATA":
            QMessageBox.information(self, "Info", "Tidak ada data yang ditemukan dalam rentang waktu atau filter yang dipilih.")
            self._update_export_button_ui("Export Gagal!", "#f44336")
        elif result.startswith("EXPORT_ERROR:"):
            QMessageBox.critical(self, "Error Export", f"Gagal mengekspor data ke Excel:\n{result[13:]}")
            self._update_export_button_ui("Export Gagal!", "#f44336")
        else:
            QMessageBox.information(self, "Sukses", f"Data berhasil diekspor ke:\n{result}")
            self._update_export_button_ui("Export Berhasil!", "#4CAF50")
            
    def _update_export_button_ui(self, text, bg_color):
        #Update styling dan teks export button untuk menunjukkan status.
        self.btn_export.setText(text)
        self.btn_export.setStyleSheet(f"background-color: {bg_color}; color: white; font-weight: bold; height: 30px;")
        QTimer.singleShot(3000, self._reset_export_button_ui)

    def _reset_export_button_ui(self):
        #Reset export button ke kondisi default.
        self.btn_export.setText("EXPORT DATA")
        self.btn_export.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")

    def _update_label_options(self, preset):
        #Update daftar label/type sesuai preset yang dipilih.
        current_selection = self.jis_type_combo.currentText()
        self.jis_type_combo.blockSignals(True)
        self.jis_type_combo.clear()
        
        if preset == "DIN":
            self.jis_type_combo.addItems(DIN_TYPES)
        else:  # JIS
            self.jis_type_combo.addItems(JIS_TYPES)
        
        # Restore previous selection jika masih valid
        index = self.jis_type_combo.findText(current_selection)
        if index >= 0:
            self.jis_type_combo.setCurrentIndex(index)
        else:
            self.jis_type_combo.setCurrentIndex(0)
        
        self.jis_type_combo.blockSignals(False)