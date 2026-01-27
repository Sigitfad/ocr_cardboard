# Konfigurasi PENGATURAN dan pengaturan aplikasi QC_GS-Battery
# File ini berisi semua PENGATURAN yang digunakan di seluruh aplikasi untuk setup global
# Tujuan: Centralized configuration untuk semua constants dan settings aplikasi
# Fungsi: Menyediakan satu tempat untuk manage semua configuration values

import os  # File system operations | Modul untuk operasi file system (buat folder, cek path, dll)
from PIL import Image  # Image processing library | Library untuk proses image (resize, format, dll)

# === PENGATURAN APLIKASI ===
# Nilai-nilai tetap yang digunakan di seluruh aplikasi untuk konfigurasi umum
# Tujuan: Define basic application information dan window dimensions
APP_NAME = "QC_GS-Battery"  # Nama aplikasi utama | Tampil di window title dan UI
APP_VERSION = "1.0.0"  # Versi aplikasi | Untuk tracking version updates
WINDOW_WIDTH = 1293  # Lebar window utama (pixels) | Ukuran default window saat dibuka
WINDOW_HEIGHT = 720  # Tinggi window utama (pixels) | Ukuran default window saat dibuka
CONTROL_PANEL_WIDTH = 280  # Lebar panel kontrol sebelah kiri (pixels) | Tempat buttons dan controls
RIGHT_PANEL_WIDTH = 280  # Lebar panel kanan untuk data display (pixels) | Tempat tabel dan export

# === DIREKTORI ===
# Path folder untuk menyimpan berbagai jenis file aplikasi
# Tujuan: Define storage locations untuk images, exports, dan database
IMAGE_DIR = "images"  # Direktori untuk menyimpan gambar scan | Folder tempat screenshot kamera disimpan
EXCEL_DIR = "file_excel"  # Direktori untuk menyimpan file Excel export | Folder tempat file export di-save
DB_FILE = "detection.db"  # File database SQLite | Database file untuk menyimpan semua deteksi

# === KAMERA ===
# PENGATURAN untuk konfigurasi kamera dan pengolahan frame
# Tujuan: Define camera resolution dan processing parameters
CAMERA_WIDTH = 1280  # Resolusi lebar kamera (pixels) | Ukuran capture dari kamera
CAMERA_HEIGHT = 720  # Resolusi tinggi kamera (pixels) | Ukuran capture dari kamera
TARGET_WIDTH = 640  # Lebar target untuk tampilan (pixels) | Ukuran display di UI
TARGET_HEIGHT = 640  # Tinggi target untuk tampilan (pixels) | Ukuran display di UI
BUFFER_SIZE = 1  # Ukuran buffer kamera (untuk mengurangi lag) | Jumlah frame yang di-buffer
SCAN_INTERVAL = 2.0  # Interval scan OCR (detik) | Berapa lama tunggu sebelum scan ulang
MAX_CAMERAS = 5  # Maksimal kamera yang dicek | Berapa banyak index kamera yang di-test

# === PENGATURAN RESAMPLING GAMBAR (KOMPATIBILITAS PILLOW) ===
# Pillow adalah library Python yang digunakan untuk mengolah gambar,
# seperti membaca, mengubah ukuran, dan memproses gambar.
# Tujuan: Setup image resampling method yang compatible dengan berbagai versi Pillow
# Fungsi: Fallback mechanism untuk support Pillow versi lama dan baru

# Resampling adalah proses menghitung ulang piksel gambar
# saat gambar diperbesar atau diperkecil agar hasilnya tetap bagus.

# LANCZOS adalah metode resampling dengan kualitas tinggi
# yang menghasilkan gambar lebih halus dan tajam,
# sangat cocok untuk memperkecil atau memperbesar gambar.

# ANTIALIAS adalah metode lama untuk menghaluskan gambar
# agar tidak terlihat pecah, biasanya tersedia di Pillow versi lama.

try:
    # Coba gunakan metode LANCZOS dari Pillow versi terbaru (10.0+)
    # Jika tersedia, ini adalah pilihan terbaik untuk kualitas gambar
    Resampling = Image.Resampling.LANCZOS

except AttributeError:
    try:
        # Jika Pillow versi lama tidak memiliki Image.Resampling
        # Gunakan LANCZOS versi lama agar tetap kompatibel (Pillow 9.x)
        Resampling = Image.LANCZOS

    except AttributeError:
        # Jika Pillow sangat lama dan LANCZOS tidak tersedia (Pillow < 9)
        # Gunakan ANTIALIAS sebagai pilihan terakhir
        Resampling = Image.ANTIALIAS


# === PRESET DAN POLA ===
# Preset format dan pattern OCR untuk deteksi kode yang berbeda
# Tujuan: Define available presets dan regex patterns untuk OCR matching
PRESETS = ["JIS", "DIN"]  # Daftar preset yang tersedia | Format kode yang bisa dideteksi

# Pattern format: [2-4 HURUF] [ANGKA][HURUF OPSIONAL] [HURUF OPSIONAL]
# Contoh DIN: LBN 1, LN0 260A, LN4 776A ISS
# Contoh JIS: 26A17, 28B19L, 50D23R(S)
PATTERNS = {
    "JIS": r"\b\d{2,3}[A-H]\d{2,3}[LR]?(?:\(S\))?\b",  # Pattern untuk format JIS | Untuk match kode JIS
    "DIN": r"^[A-Z]{2,4}\s+\d+[A-Z]?(?:\s+[A-Z]+)?$"   # Pattern untuk format DIN | Untuk match kode DIN
}

# === ALLOWLIST KARAKTER OCR ===
# Karakter yang diizinkan dalam OCR untuk mengurangi false positive detection
# Tujuan: Limit OCR recognition ke karakter yang valid saja
# Fungsi: Improve OCR accuracy dengan membatasi character set
ALLOWLIST_JIS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYLRS()'  # Karakter yang diizinkan untuk JIS OCR | Validasi hasil OCR JIS
ALLOWLIST_DIN = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ '     # Karakter yang diizinkan untuk DIN OCR | Validasi hasil OCR DIN

# === JIS TYPES (LABEL) ===
# Daftar semua tipe/label JIS yang valid untuk dipilih user
# Tujuan: Provide complete list of valid JIS battery codes
# Fungsi: Digunakan untuk validasi input user dan dropdown options
# Format JIS: [Capacity][Type Letter][Size][Terminal (L/R)][Special Marker (S)]
# Contoh: 50D23R(S) = 50 capacity, D type, 23 size, R terminal, (S) special
JIS_TYPES = [
    "Select Label . . .",  # Placeholder - tidak valid untuk scanning
    "26A17", "26A17L", "26A17R", "26A17L(S)", "26A17R(S)",
    "26A19", "26A19L", "26A19R", "26A19L(S)", "26A19R(S)",
    "28A19", "28A19L", "28A19R", "28A19L(S)", "28A19R(S)",
    "32A19", "32A19L", "32A19R", "32A19L(S)", "32A19R(S)",
    "26B17", "26B17L", "26B17R", "26B17L(S)", "26B17R(S)",
    "28B17", "28B17L", "28B17R", "28B17L(S)", "28B17R(S)",
    "28B19", "28B19L", "28B19R", "28B19L(S)", "28B19R(S)",
    "28B20", "28B20L", "28B20R", "28B20L(S)", "28B20R(S)",
    "30C24", "30C24L", "30C24R", "30C24L(S)", "30C24R(S)",
    "32B24", "32B24L", "32B24R", "32B24L(S)", "32B24R(S)",
    "32B20", "32B20L", "32B20R", "32B20L(S)", "32B20R(S)",
    "32C24", "32C24L", "32C24R", "32C24L(S)", "32C24R(S)",
    "34B17", "34B17L", "34B17R", "34B17L(S)", "34B17R(S)",
    "34B19", "34B19L", "34B19R", "34B19L(S)", "34B19R(S)",
    "36B20", "36B20L", "36B20R", "36B20L(S)", "36B20R(S)",
    "38B20", "38B20L", "38B20R", "38B20L(S)", "38B20R(S)",
    "38B19", "38B19L", "38B19R", "38B19L(S)", "38B19R(S)",
    "40B20", "40B20L", "40B20R", "40B20L(S)", "40B20R(S)",
    "42B20", "42B20L", "42B20R", "42B20L(S)", "42B20R(S)",
    "42B26", "42B26L", "42B26R", "42B26L(S)", "42B26R(S)",
    "46B19", "46B19L", "46B19R", "46B19L(S)", "46B19R(S)",
    "46B24", "46B24L", "46B24R", "46B24L(S)", "46B24R(S)",
    "46B26", "46B26L", "46B26R", "46B26L(S)", "46B26R(S)",
    "46D26", "46D26L", "46D26R", "46D26L(S)", "46D26R(S)",
    "48D26", "48D26L", "48D26R", "48D26L(S)", "48D26R(S)",
    "50B24", "50B24L", "50B24R", "50B24L(S)", "50B24R(S)",
    "50D20", "50D20L", "50D20R", "50D20L(S)", "50D20R(S)",
    "50D23", "50D23L", "50D23R", "50D23L(S)", "50D23R(S)",
    "50D26", "50D26L", "50D26R", "50D26L(S)", "50D26R(S)",
    "55D23", "55D23R", "55D23L", "55D23R(S)", "55D23L(S)",
    "55B24", "55B24L", "55B24R", "55B24L(S)", "55B24R(S)",
    "55D26", "55D26L", "55D26R", "55D26L(S)", "55D26R(S)",
    "60D23", "60D23L", "60D23R", "60D23L(S)", "60D23R(S)",
    "60D26", "60D26L", "60D26R", "60D26L(S)", "60D26R(S)",
    "60D31", "60D31L", "60D31R", "60D31L(S)", "60D31R(S)",
    "65B24", "65B24L", "65B24R", "65B24L(S)", "65B24R(S)",
    "65D23", "65D23L", "65D23R", "65D23L(S)", "65D23R(S)",
    "65D26", "65D26L", "65D26R", "65D26L(S)", "65D26R(S)",
    "65D31", "65D31L", "65D31R", "65D31L(S)", "65D31R(S)",
    "70D23", "70D23L", "70D23R", "70D23L(S)", "70D23R(S)",
    "70D31", "70D31L", "70D31R", "70D31L(S)", "70D31R(S)",
    "75D23", "75D23L", "75D23R", "75D23L(S)", "75D23R(S)",
    "75D26", "75D26L", "75D26R", "75D26L(S)", "75D26R(S)",
    "75D31", "75D31L", "75D31R", "75D31L(S)", "75D31R(S)",
    "80D23", "80D23L", "80D23R", "80D23L(S)", "80D23R(S)",
    "80D26", "80D26L", "80D26R", "80D26L(S)", "80D26R(S)",
    "80D31", "80D31L", "80D31R", "80D31L(S)", "80D31R(S)",
    "85D26", "85D26L", "85D26R", "85D26L(S)", "85D26R(S)",
    "85E41", "85E41L", "85E41R", "85E41L(S)", "85E41R(S)",
    "90D26", "90D26L", "90D26R", "90D26L(S)", "90D26R(S)",
    "95D31", "95D31L", "95D31R", "95D31L(S)", "95D31R(S)",
    "95E41", "95E41L", "95E41R", "95E41L(S)", "95E41R(S)",
    "105D31", "105D31L", "105D31R", "105D31L(S)", "105D31R(S)",
    "105E41", "105E41L", "105E41R", "105E41L(S)", "105E41R(S)",
    "105F51", "105F51L", "105F51R", "105F51L(S)", "105F51R(S)",
    "115E41", "115E41L", "115E41R", "115E41L(S)", "115E41R(S)",
    "115F51", "115F51L", "115F51R", "115F51L(S)", "115F51R(S)",
    "125D31", "125D31L", "125D31R", "125D31L(S)", "125D31R(S)",
    "130E41", "130E41L", "130E41R", "130E41L(S)", "130E41R(S)",
    "130F51", "130F51L", "130F51R", "130F51L(S)", "130F51R(S)",
    "135G51", "135G51L", "135G51R", "135G51L(S)", "135G51R(S)",
    "140H52", "140H52L", "140H52R", "140H52L(S)", "140H52R(S)",
    "145F51", "145F51L", "145F51R", "145F51L(S)", "145F51R(S)",
    "145G51", "145G51L", "145G51R", "145G51L(S)", "145G51R(S)",
    "150F51", "150F51L", "150F51R", "150F51L(S)", "150F51R(S)",
    "165G51", "165G51L", "165G51R", "165G51L(S)", "165G51R(S)",
    "170F51", "170F51L", "170F51R", "170F51L(S)", "170F51R(S)",
    "180G51", "180G51L", "180G51R", "180G51L(S)", "180G51R(S)",
    "195G51", "195G51L", "195G51R", "195G51L(S)", "195G51R(S)",
    "190H52", "190H52L", "190H52R", "190H52L(S)", "190H52R(S)",
    "245H52", "245H52L", "245H52R", "245H52L(S)", "245H52R(S)",
]

# === DIN TYPES (LABEL) ===
# Daftar semua tipe/label DIN yang valid untuk dipilih user
# Tujuan: Provide complete list of valid DIN battery codes
# Fungsi: Digunakan untuk validasi input user dan dropdown options
# Format DIN: [Prefix] [Capacity+Letter] [Optional ISS marker]
# Contoh: LN4 776A ISS = LN4 prefix, 776A capacity rating, ISS marker
DIN_TYPES = [
    "Select Label . . .",  # Placeholder - tidak valid untuk scanning
    "LBN 1", "LBN 2", "LBN 3",
    "LN1 450A", "LN1 295A",
    "LN0 260A",
    "LN2 360A", "LN2 345A",
    "LN3 490A",
    "LN4 650A", "LN4 776A ISS"
]

# === MONTHS ===
# Daftar bulan dalam bahasa Indonesia untuk dropdown date selection
# Tujuan: Provide month names untuk date picker dan export dialog
# Fungsi: User-friendly month selection dalam bahasa Indonesia
MONTHS = ["January", "February", "March", "April", "May", "June", 
          "July", "August", "September", "Oktober", "November", "Desember"]

# Mapping nama bulan ke nomor bulan (1-12)
# Tujuan: Convert nama bulan Indonesia ke integer untuk date processing
# Fungsi: Digunakan saat build date range untuk export filtering
MONTH_MAP = {
    "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, 
    "July": 7, "August": 8, "September": 9, "Oktober": 10, "November": 11, "Desember": 12
}