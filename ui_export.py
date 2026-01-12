# Komponen UI dan fungsi terkait export data ke Excel | Tujuan: Separated UI export logic dari ui.py

from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QPushButton,
    QRadioButton, QCheckBox, QComboBox, QDateEdit, QMessageBox, QCompleter
)  # PySide6 widgets | UI components untuk export dialog
from PySide6.QtCore import Qt, QTimer, QDate  # PySide6 core | Core classes
from PySide6.QtGui import QFont  # PySide6 GUI | Font untuk styling
from datetime import datetime, timedelta, time as py_time  # Date/time operations | Modul date/time
from config import MONTHS, MONTH_MAP  # Import dari config | Konfigurasi months


def create_export_dialog(parent, logic, preset_combo, jis_type_combo):
    # Fungsi membuat dan mengkonfigurasi dialog export | Tujuan: Create dialog dengan filter tanggal, label, preset untuk export data
    # Parameter: parent (parent window), logic (DetectionLogic instance), preset_combo, jis_type_combo
    # Return: Dialog object yang siap ditampilkan
    
    from database import get_detection_count
    
    if not logic:
        QMessageBox.critical(parent, "Error", "Logic belum diinisialisasi. Coba mulai dan hentikan deteksi kamera sekali.")
        return None

    count = get_detection_count(logic.db_file if hasattr(logic, 'db_file') else None)

    if count == 0:
        QMessageBox.information(parent, "Info", "Tidak ada data sama sekali di database untuk diekspor.")
        return None

    dialog = QDialog(parent)
    dialog.setWindowTitle("EXPORT DATA OPTION")
    dialog.setMinimumWidth(500)

    layout = QVBoxLayout(dialog)

    # Header label untuk instruksi
    header_label = QLabel("Ekspor Data Berdasarkan:", font=QFont("Arial", 12, QFont.Bold))
    header_label.setStyleSheet("margin-bottom: 10px;")
    layout.addWidget(header_label)

    # ===== Group Box Pilih Preset dan Sesi =====
    preset_select_group = QGroupBox(f"Pilih Tipe dan Label:")
    preset_select_layout = QVBoxLayout(preset_select_group)
    
    # Layout untuk Preset dropdown
    preset_inner_layout = QHBoxLayout()
    preset_inner_layout.addWidget(QLabel("Tipe:"))
    
    export_preset_combo = QComboBox()
    export_preset_combo.addItems(["Preset", "JIS", "DIN"])
    export_preset_combo.setCurrentText(preset_combo.currentText())
    
    def update_label_options_for_export(preset_choice):
        # Fungsi: Update daftar label/sesi ketika user mengubah preset di dialog export
        # Tujuan: Menampilkan label yang sesuai dengan preset yang dipilih (JIS atau DIN)
        if preset_choice == "Preset":
            # Gunakan preset yang aktif dari interface utama
            actual_preset = preset_combo.currentText()
        else:
            actual_preset = preset_choice
        
        # Update combo box label type sesuai preset terpilih
        if actual_preset == "DIN":
            from config import DIN_TYPES
            export_types = ["All Label"] + DIN_TYPES[1:]
        else:  # JIS
            from config import JIS_TYPES
            export_types = ["All Label"] + JIS_TYPES[1:]
        
        # Block signals untuk mencegah trigger event saat update
        export_label_type_combo.blockSignals(True)
        current_selection = export_label_type_combo.currentText()
        export_label_type_combo.clear()
        export_label_type_combo.addItems(export_types)
        
        # Restore previous selection jika masih valid
        if current_selection in export_types:
            export_label_type_combo.setCurrentText(current_selection)
        else:
            export_label_type_combo.setCurrentIndex(0)
        
        export_label_type_combo.blockSignals(False)
    
    # Hubungkan signal preset change ke update label
    export_preset_combo.currentTextChanged.connect(update_label_options_for_export)
    preset_inner_layout.addWidget(export_preset_combo)
    preset_inner_layout.addStretch()
    preset_select_layout.addLayout(preset_inner_layout)
    
    # Checkbox untuk mengaktifkan filter label/sesi
    export_label_filter_enabled = QCheckBox(f"Filter Label")
    preset_select_layout.addWidget(export_label_filter_enabled)
    
    # Layout untuk Pilih Sesi dropdown
    label_type_select_layout = QHBoxLayout()
    label_type_select_layout.addWidget(QLabel("Pilih Label:"))
    
    export_label_type_combo = QComboBox()
    
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
    
    # Setup completer untuk autocomplete label
    export_completer = export_label_type_combo.completer()
    if export_completer:
        export_completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        export_completer.setFilterMode(Qt.MatchFlag.MatchContains)
        export_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    
    export_label_type_combo.setMaxVisibleItems(15)
    export_label_type_combo.setEnabled(False)
    
    # Set default label dari main interface jika ada
    current_session = jis_type_combo.currentText()
    if current_session and current_session != "Select Label . . .":
        index = export_label_type_combo.findText(current_session)
        if index >= 0:
            export_label_type_combo.setCurrentIndex(index)
    
    # Toggle enable/disable label combo saat checkbox berubah
    export_label_filter_enabled.toggled.connect(
        lambda checked: export_label_type_combo.setEnabled(checked)
    )
    
    label_type_select_layout.addWidget(export_label_type_combo)
    preset_select_layout.addLayout(label_type_select_layout)
    
    layout.addWidget(preset_select_group)

    # ===== Date Range Selection =====
    export_range_var = "All"

    def set_range(value):
        # Fungsi: Update variabel range export saat radio button berubah | Tujuan: Menyimpan pilihan range yang dipilih user
        nonlocal export_range_var
        export_range_var = value
        dialog.export_range_var = value  # Update dialog.export_range_var agar nilai tersimpan dan bisa diakses dari parent
        toggle_selectors()

    # Radio buttons untuk pilihan tanggal
    rb_today = QRadioButton("Tanggal Hari Ini")
    rb_all = QRadioButton("Semua Data")
    rb_all.setChecked(True)

    rb_today.toggled.connect(lambda: set_range("Today") if rb_today.isChecked() else None)
    rb_all.toggled.connect(lambda: set_range("All") if rb_all.isChecked() else None)

    layout.addWidget(rb_today)
    layout.addWidget(rb_all)

    # ===== Month Selection =====
    layout.addWidget(QLabel("Pilih Bulan:", font=QFont("Arial", 12, QFont.Bold), styleSheet="margin-top: 15px;"))
    month_select_frame = QWidget()
    month_select_layout = QHBoxLayout(month_select_frame)
    month_select_layout.setContentsMargins(0, 0, 0, 0)

    rb_month = QRadioButton("Pilih Bulan:")
    rb_month.toggled.connect(lambda: set_range("Month") if rb_month.isChecked() else None)
    month_select_layout.addWidget(rb_month)

    current_year = datetime.now().year
    selected_month_var = datetime.now().strftime("%B")
    selected_year_var = str(current_year)

    years = [str(y) for y in range(current_year, current_year - 5, -1)]

    month_combo = QComboBox()
    month_combo.addItems(MONTHS)
    current_month_name = datetime.now().strftime("%B")
    if current_month_name in MONTHS:
         month_combo.setCurrentText(current_month_name)
    else:
         month_combo.setCurrentIndex(datetime.now().month - 1)

    month_combo.setDisabled(True)
    month_select_layout.addWidget(month_combo)

    year_combo = QComboBox()
    year_combo.addItems(years)
    year_combo.setCurrentText(str(current_year))
    year_combo.setDisabled(True)
    month_select_layout.addWidget(year_combo)
    layout.addWidget(month_select_frame)

    # ===== Custom Date Range Selection =====
    layout.addWidget(QLabel("Pilih Berdasarkan Tanggal:", font=QFont("Arial", 12, QFont.Bold), styleSheet="margin-top: 15px;"))
    custom_date_frame = QWidget()
    custom_date_layout = QHBoxLayout(custom_date_frame)
    custom_date_layout.setContentsMargins(0, 0, 0, 0)

    rb_custom = QRadioButton("Pilih Tanggal:")
    rb_custom.toggled.connect(lambda: set_range("CustomDate") if rb_custom.isChecked() else None)
    custom_date_layout.addWidget(rb_custom)

    custom_date_layout.addWidget(QLabel("Mulai:"))
    start_date_entry = QDateEdit()
    start_date_entry.setCalendarPopup(True)
    start_date_entry.setDisplayFormat("dd-MM-yyyy")
    start_date_entry.setDate(QDate.currentDate())
    start_date_entry.setDisabled(True)
    custom_date_layout.addWidget(start_date_entry)

    custom_date_layout.addWidget(QLabel("Akhir:"))
    end_date_entry = QDateEdit()
    end_date_entry.setCalendarPopup(True)
    end_date_entry.setDisplayFormat("dd-MM-yyyy")
    end_date_entry.setDate(QDate.currentDate())
    end_date_entry.setDisabled(True)
    custom_date_layout.addWidget(end_date_entry)
    layout.addWidget(custom_date_frame)

    # ===== Toggle Selectors Function =====
    def toggle_selectors():
        # Fungsi: Enable/disable date selector widgets berdasarkan pilihan range yang dipilih
        month_state = False
        date_state = False
        if export_range_var == "Month":
            month_state = True
        elif export_range_var == "CustomDate":
            date_state = True
        if month_combo:
            month_combo.setEnabled(month_state)
            year_combo.setEnabled(month_state)
        if start_date_entry:
            start_date_entry.setEnabled(date_state)
            end_date_entry.setEnabled(date_state)

    # ===== Export Button =====
    export_btn = QPushButton("EXPORT DATA")
    export_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px; margin-top: 20px;")
    layout.addWidget(export_btn)

    # Store dialog components untuk akses dari parent
    dialog.export_preset_combo = export_preset_combo
    dialog.export_label_filter_enabled = export_label_filter_enabled
    dialog.export_label_type_combo = export_label_type_combo
    dialog.export_range_var = export_range_var
    dialog.export_btn = export_btn
    dialog.month_combo = month_combo
    dialog.year_combo = year_combo
    dialog.start_date_entry = start_date_entry
    dialog.end_date_entry = end_date_entry
    dialog.toggle_selectors = toggle_selectors

    return dialog