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
    #Class wrapper QThread untuk DetectionLogic
    #Tujuan: Wrapper untuk DetectionLogic instance dengan signal/slot capability
    #Fungsi: Menyediakan interface komunikasi antara UI thread dan detection thread
    
    update_signal = Signal(object)  # Signal untuk update video frame | Emit ketika frame baru siap dari kamera
    code_detected_signal = Signal(str)  # Signal untuk code detected | Emit ketika kode berhasil terdeteksi OCR
    camera_status_signal = Signal(str, bool)  # Signal untuk camera status | Emit status kamera (on/off) dengan info
    data_reset_signal = Signal()  # Signal untuk reset data | Emit untuk reset display saat ganti hari
    all_text_signal = Signal(list)  # Signal untuk OCR text output | Emit list semua teks yang terdeteksi OCR

    def __init__(self):
        #Fungsi inisialisasi QThread
        #Tujuan: Setup thread dan buat DetectionLogic instance
        #Fungsi: Membuat instance DetectionLogic dan menghubungkan semua signals

        super().__init__()
        from ocr import DetectionLogic
        # Buat instance DetectionLogic dengan semua signals yang diperlukan
        self.logic = DetectionLogic(
            self.update_signal,
            self.code_detected_signal,
            self.camera_status_signal,
            self.data_reset_signal,
            self.all_text_signal
        )
        
    def run(self):
        """
        Fungsi jalankan thread
        Tujuan: Start thread event loop
        Fungsi: Menjalankan Qt event loop untuk thread ini
        """
        self.exec()


class MainWindow(QMainWindow):
    # Class jendela utama aplikasi | Tujuan: Main application window dengan semua UI components
    
    export_result_signal = Signal(str)  # Signal untuk export result | Emit hasil export file
    export_status_signal = Signal(str, str)  # Signal untuk export status | Emit status export operation
    file_scan_result_signal = Signal(str)  # Signal untuk file scan result | Emit hasil scan dari file

    def __init__(self):
        """
        Fungsi inisialisasi main window
        Tujuan: Setup UI dan inisialisasi semua components
        Fungsi: Membuat window, setup signals, dan inisialisasi semua UI elements
        """
        super().__init__()
        self.setWindowTitle(APP_NAME)  # Set judul window dari config
        self.setGeometry(100, 100, WINDOW_WIDTH, WINDOW_HEIGHT)  # Set ukuran dan posisi window
        
        self.logic_thread = None  # Instance LogicSignals thread | Akan diisi saat _setup_logic_thread
        self.logic = None  # Instance DetectionLogic | Akan diisi saat _setup_logic_thread
        
        # Connect internal signals untuk handling asynchronous operations
        self.export_result_signal.connect(self._handle_export_result)  # Handle hasil export
        self.export_status_signal.connect(self._update_export_button_ui)  # Update UI button export
        self.file_scan_result_signal.connect(self._handle_file_scan_result)  # Handle hasil scan file

        self._setup_logic_thread(initial_setup=True)  # Setup thread pertama kali
        
        self.setup_ui()  # Buat semua UI components
        self.setup_timer()  # Setup timer untuk jam real-time
    
    def _setup_logic_thread(self, initial_setup=False):
        #Helper untuk membuat instance baru LogicSignals dan Logic, lalu menghubungkan sinyal
        #Tujuan: Setup atau reset detection logic thread
        #Fungsi: Membersihkan thread lama, membuat instance baru, dan connect semua signals
        #Parameter: initial_setup (bool) - True jika setup pertama kali
        
        # Cleanup thread lama jika ada
        if self.logic_thread:
            if self.logic:
                 self.logic.stop_detection()  # Stop detection jika sedang berjalan
            
            if self.logic_thread.isRunning():
                 self.logic_thread.quit()  # Request thread untuk stop
                 self.logic_thread.wait(5000)  # Wait maksimal 5 detik
                 # Disconnect semua signals untuk mencegah memory leak
                 try:
                     self.logic_thread.update_signal.disconnect(self.update_video_frame)
                     self.logic_thread.code_detected_signal.disconnect(self.handle_code_detection)
                     self.logic_thread.camera_status_signal.disconnect(self.update_camera_status)
                     self.logic_thread.data_reset_signal.disconnect(self.update_code_display)
                     self.logic_thread.all_text_signal.disconnect(self.update_all_text_display)
                 except TypeError:
                     pass  # Ignore jika signal sudah disconnected
                     
            self.logic_thread = None  # Clear reference
            self.logic = None  # Clear reference
        
        # Buat instance baru
        self.logic_thread = LogicSignals()  # Buat thread wrapper baru
        self.logic = self.logic_thread.logic  # Ambil reference ke DetectionLogic instance
        
        # Connect semua signals ke handler functions
        self.logic_thread.update_signal.connect(self.update_video_frame)  # Update frame kamera
        self.logic_thread.code_detected_signal.connect(self.handle_code_detection)  # Handle deteksi kode
        self.logic_thread.camera_status_signal.connect(self.update_camera_status)  # Update status kamera
        self.logic_thread.data_reset_signal.connect(self.update_code_display)  # Reset display data
        self.logic_thread.all_text_signal.connect(self.update_all_text_display)  # Update OCR output
    
    def closeEvent(self, event):
        #Menangani penutupan jendela (tombol X)
        #MODIFIED: Tambah warning jika kamera sedang aktif
        #Tujuan: Prevent user close app saat kamera masih running
        #Fungsi: Validasi status kamera sebelum allow close

        # CHECK: Apakah kamera sedang aktif?
        if self.logic and self.logic.running:
            # Kamera masih aktif - tampilkan warning
            QMessageBox.warning(
                self, 
                'Warning !',
                "Kamera sedang aktif!\nHarap STOP kamera terlebih dahulu sebelum keluar aplikasi!",
                QMessageBox.Ok
            )
            # Ignore close event - aplikasi tidak akan tertutup
            event.ignore()
            return
        
        # Kamera tidak aktif - tampilkan konfirmasi normal
        reply = QMessageBox.question(
            self, 
            'Quit Confirmation',
            "Are you sure you want to quit?",
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # User konfirmasi keluar
            if self.logic:
                self.logic.stop_detection()  # Pastikan stop detection
            if self.logic_thread and self.logic_thread.isRunning():
                self.logic_thread.quit()  # Quit thread
                self.logic_thread.wait()  # Wait sampai thread selesai
            event.accept()  # Accept close event
        else:
            # User cancel keluar
            event.ignore()  # Ignore close event

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
        #Fungsi setup seluruh UI aplikasi
        #Tujuan: Create dan arrange semua UI components di window
        #Fungsi: Membuat layout utama dengan 3 panel (control, video, data)

        main_widget = QWidget()  # Buat central widget
        self.setCentralWidget(main_widget)  # Set sebagai central widget window
        
        main_layout = QHBoxLayout(main_widget)  # Buat horizontal layout utama
        
        # Panel kiri - Control panel dengan buttons dan settings
        control_frame = self._create_control_panel()
        main_layout.addWidget(control_frame)
        
        # Panel tengah - Video display dari kamera
        self.video_label = QLabel("CAMERA OFF")  # Label untuk tampilkan video
        self.video_label.setAlignment(Qt.AlignCenter)  # Center alignment
        self.video_label.setStyleSheet("background-color: black; color: white; font-size: 14pt;")
        main_layout.addWidget(self.video_label, 1)  # Stretch factor 1 untuk expand
        
        # Panel kanan - Data display dan export
        right_panel = self._create_right_panel()
        main_layout.addWidget(right_panel)

        # Set fixed width untuk side panels
        control_frame.setFixedWidth(CONTROL_PANEL_WIDTH)  # Fixed width control panel
        right_panel.setFixedWidth(RIGHT_PANEL_WIDTH)  # Fixed width data panel

    def _create_control_panel(self):
        """
        Fungsi buat panel kontrol
        Tujuan: Create left control panel dengan buttons dan combos
        Fungsi: Membuat semua UI controls (preset, options, label selector, buttons)
        Return: QWidget - Widget panel kontrol yang sudah lengkap
        """
        frame = QWidget()  # Buat widget container
        layout = QVBoxLayout(frame)  # Vertical layout
        layout.setContentsMargins(15, 15, 15, 15)  # Set margins

        # === Top Control Widget (Preset dan Options) ===
        top_control_widget = QWidget()
        top_control_layout = QHBoxLayout(top_control_widget)
        top_control_layout.setContentsMargins(0, 0, 0, 0)
        
        # === Group Box untuk Preset Selection ===
        preset_group = QGroupBox("Tipe:")  # Group box dengan label "Tipe:"
        preset_layout = QVBoxLayout(preset_group)
        preset_layout.setAlignment(Qt.AlignTop)
        preset_layout.setContentsMargins(8, 12, 8, 8)
        preset_layout.setSpacing(6)

        # Combo box untuk pilih preset (JIS atau DIN)
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["JIS", "DIN"])  # Tambah pilihan JIS dan DIN
        self.preset_combo.setCurrentIndex(0)  # Default ke JIS
        
        preset_layout.addWidget(self.preset_combo)
        preset_layout.addStretch()
        
        # Lambda function untuk set camera options saat ada perubahan
        set_options = lambda: self.logic.set_camera_options(
            self.preset_combo.currentText(),  # Preset yang dipilih
            False,  # flip_h disabled (horizontal flip)
            False,  # flip_v disabled (vertical flip)
            self.cb_edge.isChecked(),  # Binary mode status
            self.cb_split.isChecked(),  # Split screen status
            2.0  # Scan interval 2 detik
        ) if self.logic else None
        
        # Connect signals
        self.preset_combo.currentTextChanged.connect(set_options)  # Update options saat preset berubah
        self.preset_combo.currentTextChanged.connect(self._update_label_options)  # Update label options
        top_control_layout.addWidget(preset_group)
        
        # === Group Box untuk Options (Binary dan Split Screen) ===
        options_group = QGroupBox("Option:")
        options_layout = QVBoxLayout(options_group)
        
        # Checkbox untuk binary color mode (convert image ke hitam putih)
        self.cb_edge = QCheckBox("BINARY COLOR")  # Checkbox untuk mode binary (grayscale threshold)
        # Checkbox untuk split screen mode (tampilkan binary dan original)
        self.cb_split = QCheckBox("SHOW SPLIT SCREEN")  # Checkbox untuk split screen mode
        
        option_change = set_options  # Handler untuk perubahan option
        
        # Connect checkbox changes ke set_options
        self.cb_edge.toggled.connect(option_change)  # Trigger set_options saat binary mode berubah
        self.cb_split.toggled.connect(option_change)  # Trigger set_options saat split mode berubah

        # Tambah checkboxes ke layout
        options_layout.addWidget(self.cb_edge)
        options_layout.addWidget(self.cb_split)
        top_control_layout.addWidget(options_group)
        
        layout.addWidget(top_control_widget)
        
        # === Camera Status Label ===
        self.camera_label = QLabel("Camera: Not Selected")  # Label untuk tampilkan status kamera
        self.camera_label.setFont(QFont("Arial", 9))
        self.camera_label.setStyleSheet("color: blue;")
        layout.addWidget(self.camera_label)

        # === Label Selection ===
        jis_type_label = QLabel("Select Label:")  # Label instruksi
        jis_type_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(jis_type_label)
        
        # Combo box untuk pilih label/type (JIS atau DIN)
        self.jis_type_combo = QComboBox()
        self.jis_type_combo.addItems(JIS_TYPES)  # Default ke JIS_TYPES
        
        # Set combo box agar editable (user bisa ketik manual)
        self.jis_type_combo.setEditable(True)
        self.jis_type_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)  # Tidak allow insert item baru
        self.jis_type_combo.setCompleter(QCompleter(self.jis_type_combo.model()))  # Enable autocomplete
        self.jis_type_combo.currentTextChanged.connect(self.on_jis_type_changed)  # Connect ke handler
        layout.addWidget(self.jis_type_combo)
        
        # Label untuk tampilkan label yang sedang dipilih
        self.selected_type_label = QLabel("No Type Selected - Pilih Label untuk memulai")
        self.selected_type_label.setFont(QFont("Arial", 9))
        self.selected_type_label.setStyleSheet("color: #FF6600; font-weight: bold;")
        self.selected_type_label.setWordWrap(True)  # Allow text wrapping
        layout.addWidget(self.selected_type_label)

        # === Camera Control Buttons (START dan STOP) ===
        camera_btn_container = QWidget()
        camera_btn_layout = QHBoxLayout(camera_btn_container)
        camera_btn_layout.setContentsMargins(0, 0, 0, 0)
        camera_btn_layout.setSpacing(5)

        # Button START untuk mulai detection
        self.btn_start = QPushButton("START")
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 35px;")
        self.btn_start.clicked.connect(self.start_detection)  # Connect ke start_detection function
        
        # Button STOP untuk stop detection
        self.btn_stop = QPushButton("STOP")
        self.btn_stop.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 35px;")
        self.btn_stop.clicked.connect(self.stop_detection)  # Connect ke stop_detection function
        self.btn_stop.setEnabled(False)  # Disabled by default

        camera_btn_layout.addWidget(self.btn_start)
        camera_btn_layout.addWidget(self.btn_stop)
        layout.addWidget(camera_btn_container)
        
        # === Button untuk Scan dari File ===
        self.btn_file = QPushButton("SCAN FROM FILE")
        self.btn_file.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")
        self.btn_file.clicked.connect(self.open_file_scan_dialog)  # Connect ke file scan dialog
        layout.addWidget(self.btn_file)
        
        # === Container untuk Success Popup ===
        self.success_container = QWidget()
        self.success_layout = QVBoxLayout(self.success_container)
        self.success_container.setFixedHeight(50)  # Fixed height untuk popup
        layout.addWidget(self.success_container)
        
        # === Group Box untuk Detection Output (OCR Results) ===
        all_text_group = QGroupBox("Detection Output:")
        all_text_layout = QVBoxLayout(all_text_group)
        
        # Tree widget untuk tampilkan semua teks yang terdeteksi OCR
        self.all_text_tree = QTreeWidget()
        self.all_text_tree.setHeaderLabels(["Element Text"])  # Header kolom
        self.all_text_tree.header().setVisible(False)  # Hide header
        self.all_text_tree.setStyleSheet("font-size: 9pt; background-color: #f9f9f9;")
        self.all_text_tree.setMinimumHeight(150)  # Minimum height
        all_text_layout.addWidget(self.all_text_tree)
        
        layout.addWidget(all_text_group, 2)  # Stretch factor 2

        return frame

    def update_all_text_display(self, text_list):
        """
        Update list elemen teks yang terdeteksi
        Tujuan: Tampilkan semua hasil OCR di tree widget
        Fungsi: Clear tree dan populate dengan hasil OCR terbaru
        Parameter: text_list (list) - List of strings dari OCR detection
        """
        self.all_text_tree.clear()  # Clear semua items
        for text in text_list:
            item = QTreeWidgetItem([text])  # Buat tree item dengan text
            self.all_text_tree.addTopLevelItem(item)  # Tambah ke tree


    def _is_valid_label(self, label_text, current_preset):
        """
        Validasi apakah label yang diinput user valid sesuai dengan preset
        Tujuan: Memastikan label yang diketik user ada dalam daftar valid
        Fungsi: Check apakah label ada dalam JIS_TYPES atau DIN_TYPES
        Parameter: 
            label_text (str) - Label yang diinput user
            current_preset (str) - Preset aktif ("JIS" atau "DIN")
        Return: bool - True jika valid, False jika tidak valid
        """
        # Jika label kosong atau placeholder, return False
        if not label_text or label_text.strip() == "" or label_text == "Select Label . . .":
            return False
        
        # Check berdasarkan preset aktif
        if current_preset == "JIS":
            # Check apakah label ada dalam daftar JIS_TYPES (skip index 0 yang merupakan placeholder)
            return label_text in JIS_TYPES[1:]
        elif current_preset == "DIN":
            # Check apakah label ada dalam daftar DIN_TYPES (skip index 0 yang merupakan placeholder)
            return label_text in DIN_TYPES[1:]
        
        return False  # Default return False jika preset tidak dikenali

    def on_jis_type_changed(self, text):
        #Handler ketika user memilih JIS Type
        #Tujuan: Update UI dan logic saat label berubah
        #Fungsi: Validasi label, update display label, dan set target di logic
        #Parameter: text (str) - Text dari combo box yang dipilih/diketik user

        # Get preset aktif
        current_preset = self.preset_combo.currentText()
        
        # Check apakah label valid
        if not self._is_valid_label(text, current_preset):
            # Label tidak valid (kosong, placeholder, atau tidak ada dalam daftar)
            self.selected_type_label.setText("No Type Selected - Pilih Label untuk memulai")
            self.selected_type_label.setStyleSheet("color: #FF6600; font-weight: bold;")
            if self.logic:
                self.logic.set_target_label("")  # Clear target label di logic
        else:
            # Label valid
            self.selected_type_label.setText(f"Selected: {text}")
            self.selected_type_label.setStyleSheet("color: green; font-weight: bold;")
            if self.logic:
                self.logic.set_target_label(text)  # Set target label di logic
        
        self.update_code_display()  # Update tampilan data



    def _create_right_panel(self):
        #Fungsi buat panel kanan
        #Tujuan: Create right panel dengan export button dan data display table
        #Fungsi: Membuat panel untuk tampilkan data deteksi dan export controls
        #Return: QWidget - Widget panel kanan yang sudah lengkap

        frame = QWidget()  # Buat widget container
        layout = QVBoxLayout(frame)  # Vertical layout
        layout.setContentsMargins(15, 15, 15, 15)  # Set margins
        
        # === Button untuk Export Data ===
        self.btn_export = QPushButton("EXPORT DATA")
        self.btn_export.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px;")
        self.btn_export.clicked.connect(self.open_export_dialog)  # Connect ke export dialog
        layout.addWidget(self.btn_export)
        
        # === Label untuk tampilkan Date/Time ===
        self.date_time_label = QLabel("Memuat Tanggal...")  # Placeholder text
        self.date_time_label.setFont(QFont("Arial", 10))
        layout.addWidget(self.date_time_label)
        
        # === Label Header untuk Data Barang ===
        label_barang = QLabel("Data Barang :")
        label_barang.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(label_barang)
        
        # === Tree Widget untuk tampilkan data deteksi ===
        self.code_tree = QTreeWidget()
        self.code_tree.setHeaderLabels(["Waktu", "Label", "Status", "Path Gambar", "ID"])  # Set headers
        self.code_tree.setColumnCount(5)  # 5 kolom
        self.code_tree.header().setDefaultAlignment(Qt.AlignCenter)  # Center alignment header
        
        # Set column widths dan resize modes
        self.code_tree.setColumnWidth(0, 80)  # Column Waktu
        self.code_tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.code_tree.header().setSectionResizeMode(1, QHeaderView.Stretch)  # Column Label (stretch)
        self.code_tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Column Status
        self.code_tree.header().setSectionResizeMode(3, QHeaderView.Fixed)  # Column Path (hidden)
        self.code_tree.header().setSectionResizeMode(4, QHeaderView.Fixed)  # Column ID (hidden)

        self.code_tree.setSelectionMode(QAbstractItemView.MultiSelection)  # Allow multiple selection

        # Hide kolom Path dan ID (untuk internal use saja)
        self.code_tree.setColumnHidden(3, True)
        self.code_tree.setColumnHidden(4, True)
        
        # Connect double click ke view image
        self.code_tree.itemDoubleClicked.connect(self.view_selected_image)

        layout.addWidget(self.code_tree)
        
        # === Button untuk Delete Selected Data ===
        self.btn_delete_selected = QPushButton("CLEAR")
        self.btn_delete_selected.setStyleSheet("background-color: #Ff0000; color: white; font-weight: bold; height: 30px;")
        self.btn_delete_selected.clicked.connect(self.delete_selected_codes)  # Connect ke delete function
        layout.addWidget(self.btn_delete_selected)
        
        self.update_code_display()  # Initial populate data

        return frame

    def _reset_file_scan_button(self):
        """
        Meriset teks dan status tombol 'Scan from File'
        Tujuan: Reset button ke state default setelah scan selesai
        Fungsi: Ubah text ke "SCAN FROM FILE" dan enable button
        """
        self.btn_file.setText("SCAN FROM FILE")
        self.btn_file.setEnabled(True)

    def _update_camera_button_styles(self, active_button):
        #Mengatur warna tombol kamera
        #Tujuan: Update styling button berdasarkan status (START/STOP)
        #Fungsi: Ubah warna button untuk indicate button mana yang aktif
        #Parameter: active_button (str) - "START" atau "STOP"

        style_active = "background-color: #8B0000; color: white; font-weight: bold; height: 35px;"  # Red untuk aktif
        style_inactive = "background-color: #4CAF50; color: white; font-weight: bold; height: 35px;"  # Green untuk inactive

        if active_button == "START":
            self.btn_start.setStyleSheet(style_active)
            self.btn_stop.setStyleSheet(style_inactive)
        else:
            self.btn_stop.setStyleSheet(style_active)
            self.btn_start.setStyleSheet(style_inactive)
    
    def _lock_label_and_type_controls(self):
        """
        Nonaktifkan kontrol Label dan Tipe saat kamera START
        Tujuan: Prevent user mengubah preset/label saat detection sedang berjalan
        Fungsi: Disable preset combo dan label combo
        """
        self.preset_combo.setEnabled(False)
        self.jis_type_combo.setEnabled(False)

    def _unlock_label_and_type_controls(self):
        """
        Aktifkan kembali kontrol Label dan Tipe saat kamera STOP
        Tujuan: Allow user mengubah preset/label setelah detection berhenti
        Fungsi: Enable preset combo dan label combo
        """
        self.preset_combo.setEnabled(True)
        self.jis_type_combo.setEnabled(True)

    def start_detection(self):
        #Handler untuk tombol START
        #Tujuan: Mulai camera detection dan OCR scanning
        #Fungsi: Validasi label, setup logic thread, start detection, update UI

        import threading
        
        # Ambil label yang dipilih
        selected_type = self.jis_type_combo.currentText()
        current_preset = self.preset_combo.currentText()
        
        # VALIDASI: Check apakah label valid sebelum start
        if not self._is_valid_label(selected_type, current_preset):
            QMessageBox.warning(self, "Warning",
                "Tolong pilih label dengan benar!")
            return
            
        self._setup_logic_thread()  # Setup fresh logic thread
        
        # Update button states
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self._update_camera_button_styles("START")
        
        self._lock_label_and_type_controls()  # Lock preset dan label controls

        # Set camera options dan start detection
        if self.logic:
            self.logic.set_camera_options(
                self.preset_combo.currentText(),
                False,  # flip_h disabled
                False,  # flip_v disabled
                self.cb_edge.isChecked(),
                self.cb_split.isChecked(),
                2.0
            )
            self.logic.set_target_label(selected_type)  # Set target label yang sudah divalidasi

            self.logic.start_detection()  # Start detection thread
            self.logic_thread.start()  # Start Qt thread

        self._hide_success_popup()  # Hide success popup jika ada

    def stop_detection(self):
        #Handler untuk tombol STOP.
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self._update_camera_button_styles("STOP")
        
        self._unlock_label_and_type_controls()

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
            # FIXED: Kumpulkan semua ID yang akan dihapus
            record_ids = []
            for item in selected_items:
                try:
                    record_id = int(item.text(4))  # Kolom ke-4 adalah ID
                    record_ids.append(record_id)
                except (ValueError, IndexError):
                    print(f"Warning: Invalid ID untuk item: {item.text(1)}")
                    continue
            
            # FIXED: Panggil delete_codes dengan list IDs, bukan satu-satu
            if record_ids:
                success = self.logic.delete_codes(record_ids)
                if success:
                    QMessageBox.information(self, "Sukses", f"{len(record_ids)} data berhasil dihapus!")
                    self.update_code_display()  # Refresh tampilan
                else:
                    QMessageBox.critical(self, "Error", "Gagal menghapus data dari database!")
            else:
                QMessageBox.warning(self, "Warning", "Tidak ada data valid yang bisa dihapus!")

    def open_file_scan_dialog(self):
        #Membuka dialog file untuk scan
        #Tujuan: Allow user scan image dari file
        #Fungsi: Validasi label, open file dialog, start scan thread
        import threading
    
        # Ambil label yang dipilih dan preset aktif
        selected_type = self.jis_type_combo.currentText()
        current_preset = self.preset_combo.currentText()

        # VALIDASI BARU: Check apakah label valid sebelum scan file
        if not self._is_valid_label(selected_type, current_preset):
            QMessageBox.warning(self, "Warning",
                "Tolong pilih label dengan benar!")
            return

        # Check apakah live detection sedang berjalan
        if self.logic and self.logic.running:
             QMessageBox.information(self, "Info", "Harap hentikan Live Detection sebelum memindai dari file.")
             return

        self._hide_success_popup()  # Hide popup jika ada

        # Open file dialog untuk pilih image
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Image Files (*.png *.jpg *.jpeg *.bmp *.tiff)"
        )

        if file_path:
            # User memilih file - start scan
            self.btn_file.setText("SCANNING . . .")  # Update button text
            self.btn_file.setEnabled(False)  # Disable button saat scanning
            # Start scan di background thread agar tidak block UI
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