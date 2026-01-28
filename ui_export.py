# Komponen UI dan fungsi terkait export data ke Excel
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QPushButton,
    QRadioButton, QCheckBox, QComboBox, QDateEdit, QMessageBox, QCompleter
)
from PySide6.QtCore import Qt, QTimer, QDate
from PySide6.QtGui import QFont
from datetime import datetime, timedelta, time as py_time
from config import MONTHS, MONTH_MAP

def create_export_dialog(parent, logic, preset_combo, jis_type_combo):
    from database import get_detection_count
    
    if not logic:
        QMessageBox.critical(parent, "Error", "Logic belum diinisialisasi. Coba mulai dan hentikan deteksi kamera sekali.")
        return None

    count = get_detection_count(logic.db_file if hasattr(logic, 'db_file') else None)

    if count == 0:
        QMessageBox.information(parent, "Info", "Tidak ada data !")
        return None

    # Buat dialog dengan ukuran COMPACT seperti di foto
    dialog = QDialog(parent)
    dialog.setWindowTitle("EXPORT DATA OPTION")
    dialog.setFixedSize(300, 350)  # Ukuran lebih kecil dan compact
    
    # Main layout
    main_layout = QVBoxLayout(dialog)
    main_layout.setSpacing(6)  # Spacing lebih kecil
    main_layout.setContentsMargins(10, 10, 10, 10)

    # FIXED: Simpan export_range_var sebagai attribute dialog agar bisa diakses dari luar
    dialog._export_range_value = "Today"

    def set_range(r):
        dialog._export_range_value = r
        toggle_selectors()

    # Style untuk GroupBox - TANPA background putih, border tipis seperti foto
    group_style = """
        QGroupBox {
            border: 1px solid #999;
            border-radius: 5px;
            margin-top: 8px;
            padding-top: 10px;
            font-size: 12px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 13px;
            padding: 0 1px;
        }
    """
    
    combo_style = """
        QComboBox {
            border: 1px solid #ccc;
            border-radius: 3px;
            padding: 3px 6px;
            min-height: 22px;
            font-size: 12px;
        }
    """
    
    date_style = """
        QDateEdit {
            border: 1px solid #ccc;
            border-radius: 3px;
            padding: 3px 6px;
            min-height: 22px;
            font-size: 12px;
        }
    """

    # ===== ROW 1: Date (kiri) dan Pilih Tipe dan Label (kanan) =====
    top_row = QHBoxLayout()
    top_row.setSpacing(6)
    
    # KOLOM KIRI: Date
    date_group = QGroupBox("Date")
    date_group.setFont(QFont("Arial", 9, QFont.Bold))
    date_group.setStyleSheet(group_style)
    date_layout = QVBoxLayout(date_group)
    date_layout.setSpacing(4)
    date_layout.setContentsMargins(8, 8, 8, 8)
    
    rb_all = QRadioButton("Semua Data")
    rb_all.setStyleSheet("font-size: 12px;")
    rb_today = QRadioButton("Data Hari Ini")
    rb_today.setStyleSheet("font-size: 12px;")
    rb_today.setChecked(True)
    
    date_layout.addWidget(rb_all)
    date_layout.addWidget(rb_today)
    date_layout.addStretch()
    
    top_row.addWidget(date_group, 1)
    
    # KOLOM KANAN: Pilih Tipe dan Label
    preset_group = QGroupBox("Pilih Tipe dan Label")
    preset_group.setFont(QFont("Arial", 9, QFont.Bold))
    preset_group.setStyleSheet(group_style)
    preset_layout = QVBoxLayout(preset_group)
    preset_layout.setSpacing(5)
    preset_layout.setContentsMargins(8, 8, 8, 8)
    
    # Tipe row
    tipe_row = QHBoxLayout()
    tipe_row.setSpacing(5)
    tipe_label = QLabel("Tipe:")
    tipe_label.setStyleSheet("font-size: 12px;")
    tipe_label.setFixedWidth(35)
    tipe_row.addWidget(tipe_label)
    
    export_preset_combo = QComboBox()
    export_preset_combo.setStyleSheet(combo_style)
    export_preset_combo.addItems(["Preset", "JIS", "DIN"])
    export_preset_combo.setCurrentText(preset_combo.currentText())
    
    def update_label_options_for_export(preset_choice):
        if preset_choice == "Preset":
            actual_preset = preset_combo.currentText()
        else:
            actual_preset = preset_choice
        
        if actual_preset == "DIN":
            from config import DIN_TYPES
            export_types = ["All Label"] + DIN_TYPES[1:]
        else:
            from config import JIS_TYPES
            export_types = ["All Label"] + JIS_TYPES[1:]
        
        export_label_type_combo.blockSignals(True)
        current_selection = export_label_type_combo.currentText()
        export_label_type_combo.clear()
        export_label_type_combo.addItems(export_types)
        
        if current_selection in export_types:
            export_label_type_combo.setCurrentText(current_selection)
        else:
            export_label_type_combo.setCurrentIndex(0)
        
        export_label_type_combo.blockSignals(False)
    
    export_preset_combo.currentTextChanged.connect(update_label_options_for_export)
    tipe_row.addWidget(export_preset_combo)
    preset_layout.addLayout(tipe_row)
    
    # Checkbox Pilih Label
    export_label_filter_enabled = QCheckBox("Pilih Label")
    export_label_filter_enabled.setStyleSheet("font-size: 12px;")
    preset_layout.addWidget(export_label_filter_enabled)
    
    # Label dropdown
    export_label_type_combo = QComboBox()
    export_label_type_combo.setStyleSheet(combo_style)
    
    initial_preset = preset_combo.currentText()
    if initial_preset == "DIN":
        from config import DIN_TYPES
        export_types = ["All Label"] + DIN_TYPES[1:]
    else:
        from config import JIS_TYPES
        export_types = ["All Label"] + JIS_TYPES[1:]
    
    export_label_type_combo.addItems(export_types)
    export_label_type_combo.setEditable(True)
    export_label_type_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    
    export_completer = export_label_type_combo.completer()
    if export_completer:
        export_completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        export_completer.setFilterMode(Qt.MatchFlag.MatchContains)
        export_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    
    export_label_type_combo.setMaxVisibleItems(15)
    export_label_type_combo.setEnabled(False)
    
    current_session = jis_type_combo.currentText()
    if current_session and current_session != "Select Label . . .":
        index = export_label_type_combo.findText(current_session)
        if index >= 0:
            export_label_type_combo.setCurrentIndex(index)
    
    export_label_filter_enabled.toggled.connect(
        lambda checked: export_label_type_combo.setEnabled(checked)
    )
    
    preset_layout.addWidget(export_label_type_combo)
    
    top_row.addWidget(preset_group, 1)
    
    main_layout.addLayout(top_row)
    
    # Connect radio buttons
    rb_today.toggled.connect(lambda: set_range("Today") if rb_today.isChecked() else None)
    rb_all.toggled.connect(lambda: set_range("All") if rb_all.isChecked() else None)

    # ===== ROW 2: Pilih Bulan (dengan checkbox di kiri) =====
    month_group = QGroupBox("Pilih Bulan")
    month_group.setFont(QFont("Arial", 9, QFont.Bold))
    month_group.setStyleSheet(group_style)
    month_layout = QHBoxLayout(month_group)
    month_layout.setContentsMargins(8, 12, 8, 8)
    month_layout.setSpacing(6)
    
    # Checkbox untuk enable/disable pilihan bulan
    cb_month = QCheckBox()
    cb_month.setStyleSheet("font-size: 12px;")
    month_layout.addWidget(cb_month)
    
    # Month combo
    current_year = datetime.now().year
    years = [str(y) for y in range(current_year, current_year - 5, -1)]

    month_combo = QComboBox()
    month_combo.setStyleSheet(combo_style)
    month_combo.addItems(MONTHS)
    
    current_month_name = datetime.now().strftime("%B")
    if current_month_name in MONTHS:
         month_combo.setCurrentText(current_month_name)
    else:
         month_combo.setCurrentIndex(datetime.now().month - 1)

    month_combo.setDisabled(True)
    month_layout.addWidget(month_combo)

    # Year combo
    year_combo = QComboBox()
    year_combo.setStyleSheet(combo_style)
    year_combo.addItems(years)
    year_combo.setCurrentText(str(current_year))
    year_combo.setDisabled(True)
    month_layout.addWidget(year_combo)
    
    main_layout.addWidget(month_group)

    # ===== ROW 3: Pilih Tanggal (dengan checkbox di kiri) =====
    date_range_group = QGroupBox("Pilih Tanggal")
    date_range_group.setFont(QFont("Arial", 9, QFont.Bold))
    date_range_group.setStyleSheet(group_style)
    date_range_layout = QHBoxLayout(date_range_group)
    date_range_layout.setContentsMargins(8, 12, 8, 8)
    date_range_layout.setSpacing(6)
    
    # Checkbox untuk enable/disable pilihan tanggal
    cb_custom = QCheckBox()
    cb_custom.setStyleSheet("font-size: 12px;")
    date_range_layout.addWidget(cb_custom)
    
    # Start date
    start_date_entry = QDateEdit()
    start_date_entry.setStyleSheet(date_style)
    start_date_entry.setCalendarPopup(True)
    start_date_entry.setDisplayFormat("dd-MM-yyyy")
    start_date_entry.setDate(QDate.currentDate())
    start_date_entry.setDisabled(True)
    date_range_layout.addWidget(start_date_entry)
    
    # Dash label
    dash_label = QLabel(" ──")
    dash_label.setStyleSheet("font-size: 13px; font-weight: bold; color: black;")
    date_range_layout.addWidget(dash_label)
    
    # End date
    end_date_entry = QDateEdit()
    end_date_entry.setStyleSheet(date_style)
    end_date_entry.setCalendarPopup(True)
    end_date_entry.setDisplayFormat("dd-MM-yyyy")
    end_date_entry.setDate(QDate.currentDate())
    end_date_entry.setDisabled(True)
    date_range_layout.addWidget(end_date_entry)
    
    main_layout.addWidget(date_range_group)

    # ===== LOGIKA CHECKBOX =====
    # Fungsi untuk handle checkbox "Pilih Bulan"
    def on_month_checkbox_toggled(checked):
        if checked:
            # Jika diceklis, pilih "Semua Data" dan disable "Data Hari Ini"
            rb_all.setChecked(True)
            rb_today.setEnabled(False)
            # Aktifkan combo bulan/tahun
            month_combo.setEnabled(True)
            year_combo.setEnabled(True)
            # Nonaktifkan checkbox "Pilih Tanggal"
            cb_custom.setChecked(False)
            set_range("Month")
        else:
            # Jika unchecked, enable kembali "Data Hari Ini"
            rb_today.setEnabled(True)
            # Nonaktifkan combo bulan/tahun
            month_combo.setEnabled(False)
            year_combo.setEnabled(False)
            # Kembali ke "Data Hari Ini" jika tidak ada pilihan lain
            if not cb_custom.isChecked():
                rb_today.setChecked(True)
                set_range("Today")
    
    # Fungsi untuk handle checkbox "Pilih Tanggal"
    def on_custom_checkbox_toggled(checked):
        if checked:
            # Jika diceklis, pilih "Semua Data" dan disable "Data Hari Ini"
            rb_all.setChecked(True)
            rb_today.setEnabled(False)
            # Aktifkan date pickers
            start_date_entry.setEnabled(True)
            end_date_entry.setEnabled(True)
            # Nonaktifkan checkbox "Pilih Bulan"
            cb_month.setChecked(False)
            set_range("CustomDate")
        else:
            # Jika unchecked, enable kembali "Data Hari Ini"
            rb_today.setEnabled(True)
            # Nonaktifkan date pickers
            start_date_entry.setEnabled(False)
            end_date_entry.setEnabled(False)
            # Kembali ke "Data Hari Ini" jika tidak ada pilihan lain
            if not cb_month.isChecked():
                rb_today.setChecked(True)
                set_range("Today")
    
    # Connect checkbox signals
    cb_month.toggled.connect(on_month_checkbox_toggled)
    cb_custom.toggled.connect(on_custom_checkbox_toggled)

    # ===== Toggle Function (updated untuk support checkbox) =====
    def toggle_selectors():
        # Fungsi ini sekarang tidak perlu melakukan apa-apa
        # karena logika sudah di-handle oleh checkbox toggled
        pass

    # ===== EXPORT BUTTON (BIRU BESAR seperti di foto) =====
    export_btn = QPushButton("EXPORT DATA")
    export_btn.setStyleSheet("""
        QPushButton {
            background-color: #0000FF;
            color: white;
            font-weight: bold;
            font-size: 10px;
            border-radius: 5px;
            padding: 8px;
            margin-top: 6px;
            border: none;
            min-height: 15px;
        }
        QPushButton:hover {
            background-color: #0000DD;
        }
        QPushButton:pressed {
            background-color: #0000BB;
        }
    """)
    
    main_layout.addWidget(export_btn)

    # Store references - FIXED: Gunakan _export_range_value attribute
    dialog.export_preset_combo = export_preset_combo
    dialog.export_label_filter_enabled = export_label_filter_enabled
    dialog.export_label_type_combo = export_label_type_combo
    # export_range_var tidak perlu lagi karena sudah ada di dialog._export_range_value
    dialog.export_btn = export_btn
    dialog.month_combo = month_combo
    dialog.year_combo = year_combo
    dialog.start_date_entry = start_date_entry
    dialog.end_date_entry = end_date_entry
    dialog.toggle_selectors = toggle_selectors
    dialog.cb_month = cb_month
    dialog.cb_custom = cb_custom

    return dialog