# Entry point aplikasi QC_GS-Battery
# File ini adalah file yang dijalankan untuk memulai aplikasi | Tujuan: Main entry point untuk application startup

import sys  # Module untuk system operations | Modul untuk system-level operations
from PySide6.QtWidgets import QApplication, QMessageBox  # GUI framework widgets | PySide6 UI components
from PySide6.QtCore import QLocale  # Untuk set locale/bahasa | Untuk set language/locale settings
from ui import MainWindow  # Import main window class | Import MainWindow dari ui module


def main():
    # Fungsi main untuk menjalankan aplikasi | Tujuan: Entry point utama aplikasi - setup dan run QApplication
    
    app = QApplication(sys.argv)
    
    try:
        # Cek apakah semua dependency library sudah terinstall - validasi required libraries sebelum jalankan app
        import PIL.Image as Image #PIL (Pillow) digunakan untuk pengolahan gambar (image processing), seperti membuka gambar, resize, crop, rotate, convert format (RGB, grayscale), dll.
        import numpy as np #NumPy digunakan untuk operasi numerik dan array (matriks),untuk mengolah data gambar (pixel).
        import easyocr #EasyOCR adalah library Optical Character Recognition (OCR), fungsinya untuk membaca teks dari gambar secara otomatis.
        import pandas #Pandas digunakan untuk pengolahan dan manajemen data berbentuk tabel (DataFrame), seperti  mengelola kolom/baris, filtering data.
        import openpyxl #OpenPyXL digunakan untuk membaca dan menulis file Excel (.xlsx).
        import xlsxwriter #XlsxWriter digunakan untuk membuat file Excel (.xlsx) dari awal, fokus pada penulisan data dan formatting (warna cell, border, merge cell, dll).
    except ImportError as e:
        # Tampilkan error dialog jika ada library yang missing - user harus install dependencies dulu
        QMessageBox.critical(None, "Dependency Error", f"Library yang dibutuhkan tidak ditemukan. Harap instal:\nPySide6, opencv-python, easyocr, pandas, openpyxl, xlsxwriter, pillow, numpy.\nError: {e}")
        sys.exit(1)
        
    # Set locale ke Indonesian untuk date/time formatting - tampilkan tanggal/waktu dalam bahasa Indonesia
    locale = QLocale(QLocale.Indonesian, QLocale.Indonesia)
    QLocale.setDefault(locale)
    
    # Buat instance MainWindow dan tampilkan - inisialisasi UI utama dan tampilkan ke screen
    window = MainWindow()
    window.show()
    
    # Jalankan event loop aplikasi - start application main loop
    sys.exit(app.exec())


if __name__ == "__main__":
    # Jalankan fungsi main jika file ini dijalankan langsung - standard Python entry point check
    main()
