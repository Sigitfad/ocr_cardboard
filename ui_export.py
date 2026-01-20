# Komponen UI dan fungsi terkait export data ke Excel | Tujuan: Separated UI export logic dari ui.py
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QPushButton,
    QRadioButton, QCheckBox, QComboBox, QDateEdit, QMessageBox, QCompleter
)  #Import PySide6 widgets untuk komponen UI dialog export
from PySide6.QtCore import Qt, QTimer, QDate #Import PySide6 core untuk enumerasi, timer, dan date handling
from PySide6.QtGui import QFont #Import PySide6 GUI untuk styling font
from datetime import datetime, timedelta, time as py_time  #Import datetime modules untuk manipulasi tanggal dan waktu
from config import MONTHS, MONTH_MAP  #Import dari config | Konfigurasi months

def create_export_dialog(parent, logic, preset_combo, jis_type_combo):
    # Fungsi membuat dan mengkonfigurasi dialog export | Tujuan: Create dialog dengan filter tanggal, label, preset untuk export data
    # Parameter: parent (parent window), logic (DetectionLogic instance), preset_combo, jis_type_combo
    # Return: Dialog object yang siap ditampilkan

    from database import get_detection_count #Import fungsi untuk mengecek jumlah data di database
    
    # Validasi logic object sudah diinisialisasi
    # Logic diperlukan untuk akses database dan data detection
    if not logic:
        QMessageBox.critical(parent, "Error", "Logic belum diinisialisasi. Coba mulai dan hentikan deteksi kamera sekali.")
        return None

    # Hitung jumlah data yang tersedia di database untuk export
    # Ambil db_file path dari logic object jika ada
    count = get_detection_count(logic.db_file if hasattr(logic, 'db_file') else None)

    # Jika database kosong (tidak ada data), tampilkan pesan info dan batalkan export
    if count == 0:
        QMessageBox.information(parent, "Info", "Tidak ada data sama sekali di database untuk diekspor.")
        return None

    # Buat dialog window baru dengan parent yang diberikan
    dialog = QDialog(parent)
    dialog.setWindowTitle("EXPORT DATA OPTION")  # Set judul window dialog
    dialog.setMinimumWidth(500)  # Set lebar minimum dialog 500px

    # Buat vertical layout sebagai main layout dialog
    layout = QVBoxLayout(dialog)

    # Header label untuk instruksi | Memberitahu user tentang fungsi dialog
    header_label = QLabel("Ekspor Data Berdasarkan:", font=QFont("Arial", 12, QFont.Bold))
    header_label.setStyleSheet("margin-bottom: 10px;")  # Tambah margin bawah untuk spacing
    layout.addWidget(header_label)

    # ===== Group Box Pilih Preset dan Sesi =====
    # Group box untuk pemilihan tipe battery (JIS/DIN) dan filter label/sesi
    preset_select_group = QGroupBox(f"Pilih Tipe dan Label:")
    preset_select_layout = QVBoxLayout(preset_select_group)
    
    # Layout horizontal untuk dropdown pemilihan preset (JIS/DIN)
    preset_inner_layout = QHBoxLayout()
    preset_inner_layout.addWidget(QLabel("Tipe:"))  # Label "Tipe:"
    
    # ComboBox untuk memilih preset battery (Preset/JIS/DIN)
    export_preset_combo = QComboBox()
    export_preset_combo.addItems(["Preset", "JIS", "DIN"])  # Tambahkan opsi preset
    export_preset_combo.setCurrentText(preset_combo.currentText())  # Set default sesuai preset aktif di main UI
    
    def update_label_options_for_export(preset_choice):
        # Fungsi: Update daftar label/sesi ketika user mengubah preset di dialog export
        # Tujuan: Menampilkan label yang sesuai dengan preset yang dipilih (JIS atau DIN)
        # Parameter: preset_choice (string) - pilihan preset dari combo box
        
        # Jika user pilih "Preset", gunakan preset yang sedang aktif di main interface
        if preset_choice == "Preset":
            actual_preset = preset_combo.currentText() #Gunakan preset yang aktif dari interface utama
        else:
            actual_preset = preset_choice #Jika user pilih JIS atau DIN langsung, gunakan pilihan itu
        
        # Update combo box label type sesuai preset terpilih
        # Jika preset DIN, load DIN_TYPES, jika JIS load JIS_TYPES
        if actual_preset == "DIN":
            from config import DIN_TYPES
            export_types = ["All Label"] + DIN_TYPES[1:]  # "All Label" + semua tipe DIN kecuali index 0
        else:  # JIS
            from config import JIS_TYPES
            export_types = ["All Label"] + JIS_TYPES[1:]  # "All Label" + semua tipe JIS kecuali index 0
        
        # Block signals untuk mencegah trigger event saat update, ini mencegah infinite loop saat mengubah items di combo box
        export_label_type_combo.blockSignals(True)
        current_selection = export_label_type_combo.currentText()  # Simpan pilihan saat ini
        export_label_type_combo.clear()  # Hapus semua item lama
        export_label_type_combo.addItems(export_types)  # Tambah item baru sesuai preset
        
        # Jika pilihan sebelumnya masih ada di list baru, kembalikan
        if current_selection in export_types:
            export_label_type_combo.setCurrentText(current_selection)
        else:
            export_label_type_combo.setCurrentIndex(0) #Jika tidak ada, set ke index 0 (All Label)
          
        export_label_type_combo.blockSignals(False) #Unblock signals setelah selesai update
    
    # Ketika user ubah preset, otomatis update daftar label yang tersedia
    export_preset_combo.currentTextChanged.connect(update_label_options_for_export)
    preset_inner_layout.addWidget(export_preset_combo)  # Tambah combo box ke layout
    preset_inner_layout.addStretch()  # Tambah stretch untuk align kiri
    preset_select_layout.addLayout(preset_inner_layout)  # Tambah layout preset ke group
    
    #Checkbox untuk mengaktifkan filter label/sesi
    #Jika unchecked, export semua label. Jika checked, filter berdasarkan label tertentu
    export_label_filter_enabled = QCheckBox(f"Filter Label")
    preset_select_layout.addWidget(export_label_filter_enabled)
    
    # Layout horizontal untuk dropdown pemilihan label/sesi
    label_type_select_layout = QHBoxLayout()
    label_type_select_layout.addWidget(QLabel("Pilih Label:"))  # Label "Pilih Label:"
    
    # ComboBox untuk memilih label/sesi spesifik (editable dengan autocomplete)
    export_label_type_combo = QComboBox()
    
    # Inisialisasi daftar label berdasarkan preset yang aktif saat ini
    initial_preset = preset_combo.currentText()
    if initial_preset == "DIN":
        from config import DIN_TYPES
        export_types = ["All Label"] + DIN_TYPES[1:]  # Load DIN types
    else:
        from config import JIS_TYPES
        export_types = ["All Label"] + JIS_TYPES[1:]  # Load JIS types
    
    # Tambahkan items ke combo box dan set editable untuk autocomplete
    export_label_type_combo.addItems(export_types)
    export_label_type_combo.setEditable(True)  # Bisa diketik untuk searching
    export_label_type_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)  # Tidak bisa insert item baru
    
    # Setup completer untuk autocomplete label
    # Completer membantu user mencari label dengan mengetik sebagian nama
    export_completer = export_label_type_combo.completer()
    if export_completer:
        export_completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)  # Tampilkan popup suggestions
        export_completer.setFilterMode(Qt.MatchFlag.MatchContains)  # Match jika mengandung substring
        export_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)  # Case insensitive search
    
    # Set max visible items di dropdown untuk mencegah dropdown terlalu panjang
    export_label_type_combo.setMaxVisibleItems(15)
    export_label_type_combo.setEnabled(False)  # Disabled by default, enabled saat checkbox dicentang
    
    # Set default label dari main interface jika ada
    # Ini membuat dialog otomatis pilih label yang sama dengan yang sedang aktif
    current_session = jis_type_combo.currentText()
    if current_session and current_session != "Select Label . . .":
        index = export_label_type_combo.findText(current_session)  # Cari index label
        if index >= 0:
            export_label_type_combo.setCurrentIndex(index)  # Set sebagai pilihan default
    
    # Toggle enable/disable label combo saat checkbox berubah
    # Jika checkbox dicentang, enable combo box. Jika tidak, disable
    export_label_filter_enabled.toggled.connect(
        lambda checked: export_label_type_combo.setEnabled(checked)
    )
    
    # Tambahkan combo box label ke layout dan layout ke group
    label_type_select_layout.addWidget(export_label_type_combo)
    preset_select_layout.addLayout(label_type_select_layout)
    
    layout.addWidget(preset_select_group) #Tambahkan group box preset & label ke main layout

    # ===== Date Range Selection =====
    # Variabel untuk menyimpan pilihan range export (All, Today, Month, CustomDate)
    export_range_var = "All"

    def set_range(value):
        # Fungsi: Update variabel range export saat radio button berubah | Tujuan: Menyimpan pilihan range yang dipilih user
        # Parameter: value (string) - nilai range yang dipilih (All/Today/Month/CustomDate)
        nonlocal export_range_var  # Akses variabel dari outer scope
        export_range_var = value  # Update nilai range
        dialog.export_range_var = value  # Update dialog.export_range_var agar nilai tersimpan dan bisa diakses dari parent
        toggle_selectors()  # Toggle enable/disable date selector widgets
    
    rb_today = QRadioButton("Tanggal Hari Ini") #Radio button untuk export data hari ini saja
    
    # Radio button untuk export semua data (default checked)
    rb_all = QRadioButton("Semua Data")
    rb_all.setChecked(True)  # Set sebagai pilihan default

    # Connect radio buttons ke fungsi set_range
    # Ketika rb_today dicentang, set range ke "Today"
    rb_today.toggled.connect(lambda: set_range("Today") if rb_today.isChecked() else None)
    # Ketika rb_all dicentang, set range ke "All"
    rb_all.toggled.connect(lambda: set_range("All") if rb_all.isChecked() else None)

    # Tambahkan radio buttons ke main layout
    layout.addWidget(rb_today)
    layout.addWidget(rb_all)

    # ===== Month Selection =====
    # Header label untuk pemilihan bulan
    layout.addWidget(QLabel("Pilih Bulan:", font=QFont("Arial", 12, QFont.Bold), styleSheet="margin-top: 15px;"))
    
    # Frame/widget container untuk month selection
    month_select_frame = QWidget()
    month_select_layout = QHBoxLayout(month_select_frame)
    month_select_layout.setContentsMargins(0, 0, 0, 0)  # Hapus margins

    # Radio button untuk mengaktifkan mode pilih bulan
    rb_month = QRadioButton("Pilih Bulan:")
    rb_month.toggled.connect(lambda: set_range("Month") if rb_month.isChecked() else None)
    month_select_layout.addWidget(rb_month)

    # Ambil tahun saat ini untuk default value
    current_year = datetime.now().year
    selected_month_var = datetime.now().strftime("%B")  # Nama bulan saat ini (e.g., "January")
    selected_year_var = str(current_year)  # Tahun saat ini sebagai string

    years = [str(y) for y in range(current_year, current_year - 5, -1)] #Buat list tahun (tahun ini dan 4 tahun sebelumnya)

    # ComboBox untuk memilih bulan (January-December)
    month_combo = QComboBox()
    month_combo.addItems(MONTHS)  # MONTHS dari config.py
    
    # Set bulan saat ini sebagai default
    current_month_name = datetime.now().strftime("%B")
    if current_month_name in MONTHS:
         month_combo.setCurrentText(current_month_name)  # Set by nama bulan
    else:
         month_combo.setCurrentIndex(datetime.now().month - 1)  # Set by index (0-11)

    month_combo.setDisabled(True)  # Disabled by default, enabled saat rb_month dicentang
    month_select_layout.addWidget(month_combo)

    # ComboBox untuk memilih tahun
    year_combo = QComboBox()
    year_combo.addItems(years)  # List tahun yang sudah dibuat
    year_combo.setCurrentText(str(current_year))  # Set tahun saat ini sebagai default
    year_combo.setDisabled(True)  # Disabled by default
    month_select_layout.addWidget(year_combo)
    
    layout.addWidget(month_select_frame) #Tambahkan month selection frame ke main layout

    # ===== Custom Date Range Selection =====
    # Header label untuk pemilihan tanggal custom
    layout.addWidget(QLabel("Pilih Berdasarkan Tanggal:", font=QFont("Arial", 12, QFont.Bold), styleSheet="margin-top: 15px;"))
    
    # Frame/widget container untuk custom date selection
    custom_date_frame = QWidget()
    custom_date_layout = QHBoxLayout(custom_date_frame)
    custom_date_layout.setContentsMargins(0, 0, 0, 0)  # Hapus margins

    # Radio button untuk mengaktifkan mode pilih tanggal custom
    rb_custom = QRadioButton("Pilih Tanggal:")
    rb_custom.toggled.connect(lambda: set_range("CustomDate") if rb_custom.isChecked() else None)
    custom_date_layout.addWidget(rb_custom)

    # Label dan DateEdit untuk tanggal mulai
    custom_date_layout.addWidget(QLabel("Mulai:"))
    start_date_entry = QDateEdit()
    start_date_entry.setCalendarPopup(True)  # Tampilkan calendar popup saat diklik
    start_date_entry.setDisplayFormat("dd-MM-yyyy")  # Format tanggal: hari-bulan-tahun
    start_date_entry.setDate(QDate.currentDate())  # Set tanggal hari ini sebagai default
    start_date_entry.setDisabled(True)  # Disabled by default, enabled saat rb_custom dicentang
    custom_date_layout.addWidget(start_date_entry)

    # Label dan DateEdit untuk tanggal akhir
    custom_date_layout.addWidget(QLabel("Akhir:"))
    end_date_entry = QDateEdit()
    end_date_entry.setCalendarPopup(True)  # Tampilkan calendar popup saat diklik
    end_date_entry.setDisplayFormat("dd-MM-yyyy")  # Format tanggal: hari-bulan-tahun
    end_date_entry.setDate(QDate.currentDate())  # Set tanggal hari ini sebagai default
    end_date_entry.setDisabled(True)  # Disabled by default
    custom_date_layout.addWidget(end_date_entry)
    
    layout.addWidget(custom_date_frame) #Tambahkan custom date frame ke main layout

    # ===== Toggle Selectors Function =====
    def toggle_selectors():
        # Fungsi: Enable/disable date selector widgets berdasarkan pilihan range yang dipilih
        # Tujuan: Hanya aktifkan widget yang relevan dengan pilihan user
        # Misalnya: jika pilih "Month", aktifkan month/year combo, disable start/end date
        
        # Default semua disabled
        month_state = False
        date_state = False
        
        # Jika pilih "Month", enable month & year selector
        if export_range_var == "Month":
            month_state = True
        # Jika pilih "CustomDate", enable start & end date selector
        elif export_range_var == "CustomDate":
            date_state = True
        
        # Update state month/year combo boxes
        if month_combo:
            month_combo.setEnabled(month_state)
            year_combo.setEnabled(month_state)
        
        # Update state start/end date entries
        if start_date_entry:
            start_date_entry.setEnabled(date_state)
            end_date_entry.setEnabled(date_state)

    # ===== Export Button =====
    # Button untuk trigger proses export ke Excel
    export_btn = QPushButton("EXPORT DATA")
    export_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 30px; margin-top: 20px;")
    layout.addWidget(export_btn)

    # Store dialog components untuk akses dari parent
    # Simpan reference ke semua widget penting agar bisa diakses dari fungsi parent
    dialog.export_preset_combo = export_preset_combo  # Reference preset combo
    dialog.export_label_filter_enabled = export_label_filter_enabled  # Reference checkbox filter label
    dialog.export_label_type_combo = export_label_type_combo  # Reference combo label/sesi
    dialog.export_range_var = export_range_var  # Reference variabel range yang dipilih
    dialog.export_btn = export_btn  # Reference button export
    dialog.month_combo = month_combo  # Reference month selector
    dialog.year_combo = year_combo  # Reference year selector
    dialog.start_date_entry = start_date_entry  # Reference start date picker
    dialog.end_date_entry = end_date_entry  # Reference end date picker
    dialog.toggle_selectors = toggle_selectors  # Reference fungsi toggle selectors

    return dialog #Return dialog object yang sudah dikonfigurasi lengkap