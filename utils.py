"""
Fungsi utility dan helper yang digunakan di berbagai modul.
"""

import os  # File system operations | Modul untuk operasi file system
import re  # Regular expression operations | Modul untuk pattern matching dan text processing
import numpy as np  # Numerical computing | Library untuk numerical operations
import cv2  # Computer vision library | Library OpenCV untuk image processing
from datetime import datetime  # Date/time operations | Modul untuk date/time handling
from config import Resampling  # Image resampling method dari config | Metode resize image

# === OCR CORRECTION LOGIC ===
# Fungsi-fungsi koreksi untuk mengatasi error OCR yang umum terjadi

def fix_common_ocr_errors_jis(text):
    # Fungsi koreksi OCR khusus untuk format JIS | Tujuan: Perbaiki error OCR dan normalisasi format kode JIS
    # Format: [2-3 digit][1 huruf][2-3 digit][L/R optional][(S) optional]
    # Parameter: text = String kode JIS yang terdeteksi (mungkin ada error)
    # Return: String kode JIS yang sudah dikoreksi dan dinormalisasi
    
    text = text.strip().upper()
    text = re.sub(r'[^A-Z0-9()]', '', text)

    char_to_digit = {
        "O": "0", "Q": "0", "D": "0", "U": "0", "C": "0",
        "I": "1", "L": "1", "J": "1",
        "Z": "2", "E": "3", "A": "4", "H": "4",
        "S": "5", "G": "6", "T": "7", "Y": "7",
        "B": "8", "P": "9", "R": "9"
    }

    digit_to_char = {
        "0": "D", "1": "L", "2": "Z", "3": "B", "4": "A", "5": "S", 
        "6": "G", "7": "T", "8": "B", "9": "R", "D" : "G"
    }

    match = re.search(r'(\d+|[A-Z]+)(\d+|[A-Z])(\d+|[A-Z]+)([L|R|1|0|4|D|I]?)(\(S\)|5\)|S)?$', text)

    if match:
        capacity = match.group(1)
        type_char = match.group(2)
        size = match.group(3)
        terminal = match.group(4)
        option = match.group(5)

        new_capacity = "".join([char_to_digit.get(c, c) for c in capacity])
        
        if type_char.isdigit():
            new_type = digit_to_char.get(type_char, type_char)
        else:
            new_type = type_char
        
        if new_type in ['O', 'Q', 'G', '0', 'U', 'C']: new_type = 'D'
        if new_type in ['8', '3']: new_type = 'B'
        if new_type in ['4']: new_type = 'A'

        new_size = "".join([char_to_digit.get(c, c) for c in size])

        if terminal:
            if terminal in ['1', 'I', 'J', '4']: 
                terminal = 'L'
            elif terminal in ['0', 'Q', 'D', 'O']: 
                terminal = 'R'

        if option:
            option = '(S)'

        text_fixed = f"{new_capacity}{new_type}{new_size}{terminal}{option if option else ''}"
        return text_fixed.strip().upper()

    for char, digit in char_to_digit.items():
        text = text.replace(char, digit)
    
    text = text.replace('5)', '(S)').replace('(5)', '(S)')
    return text.strip().upper()


def fix_common_ocr_errors_din(text):
    # Fungsi koreksi OCR khusus untuk format DIN | Tujuan: Perbaiki error OCR pada format DIN dan normalisasi spasi
    # Format: [KODE_ALPHA] [ANGKA dengan optional huruf] [optional ISS]
    # Contoh: LBN 1, LN0 260A, LN4 776A ISS
    # Parameter: text = String kode DIN yang terdeteksi (mungkin ada error)
    # Return: String kode DIN yang sudah dikoreksi dan dinormalisasi
    
    text = text.strip().upper()
    
    # Mapping OCR errors yang umum
    char_to_digit = {
        'O': '0', 'Q': '0',  # O dan Q sering terbaca sebagai 0
        'I': '1', 'l': '1',  # I dan l sering terbaca sebagai 1
        'Z': '2',            # Z terbaca sebagai 2
        'S': '5',            # S bisa terbaca sebagai 5
        'G': '6',            # G terbaca sebagai 6
        'B': '8',            # B terbaca sebagai 8
    }
    
    digit_to_char = {
        '0': 'O',
        '1': 'I',
    }
    
    # Hilangkan karakter non-alphanumeric kecuali spasi
    text = re.sub(r'[^A-Z0-9\s]', '', text)
    
    # Pisahkan menjadi tokens
    tokens = text.split()
    
    if len(tokens) == 0:
        return text
    
    corrected_tokens = []
    
    for i, token in enumerate(tokens):
        if i == 0:
            # Token pertama: harus prefix huruf (LBN, LN0, LN1, dll)
            # Format: [L][B/N][N/0-4]
            corrected = ""
            for j, char in enumerate(token):
                if j == 0:
                    # Huruf pertama harus 'L'
                    if char in ['I', '1', 'l']:
                        corrected += 'L'
                    else:
                        corrected += char
                elif j == 1:
                    # Huruf kedua: B atau N
                    if char in ['8']:
                        corrected += 'B'
                    elif char in ['H', 'M']:
                        corrected += 'N'
                    else:
                        corrected += char
                elif j == 2:
                    # Bisa huruf (N) atau angka (0-4)
                    if token[:2] == "LB":
                        # LBN format
                        if char in ['H', 'M']:
                            corrected += 'N'
                        else:
                            corrected += char
                    else:
                        # LN[digit] format
                        if char in char_to_digit:
                            corrected += char_to_digit[char]
                        else:
                            corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
            
        elif i == 1:
            # Token kedua: angka dengan optional huruf di akhir (260A, 450A, 1, 2, 3)
            corrected = ""
            for j, char in enumerate(token):
                if char.isdigit():
                    corrected += char
                elif char in char_to_digit:
                    corrected += char_to_digit[char]
                elif j == len(token) - 1 and char.isalpha():
                    # Huruf di akhir (biasanya 'A')
                    if char in ['4', 'H']:
                        corrected += 'A'
                    else:
                        corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
            
        elif i == 2:
            # Token ketiga: biasanya ISS
            corrected = token
            if token in ['I55', 'IS5', 'I5S', '155']:
                corrected = 'ISS'
            elif token.replace('5', 'S').replace('I', 'I') == 'ISS':
                corrected = 'ISS'
            
            corrected_tokens.append(corrected)
    
    result = ' '.join(corrected_tokens)
    
    # Final cleanup: pastikan format yang valid
    # LBN harus diikuti spasi dan satu digit
    result = re.sub(r'(LBN)(\d)', r'\1 \2', result)
    # LN[digit] harus diikuti spasi dan angka
    result = re.sub(r'(LN\d)(\d)', r'\1 \2', result)
    # Pastikan ada spasi sebelum ISS
    result = re.sub(r'([A-Z0-9])(ISS)', r'\1 \2', result)
    # Bersihkan multiple spaces
    result = re.sub(r'\s+', ' ', result)
    
    return result.strip()


def fix_common_ocr_errors(text, preset):
    # Main function yang memilih koreksi berdasarkan preset aktif | Tujuan: Dispatcher function untuk koreksi JIS atau DIN
    # Parameter: text = String kode yang terdeteksi, preset = String "JIS" atau "DIN"
    # Return: String kode yang sudah dikoreksi sesuai preset
    
    if preset == "JIS":
        return fix_common_ocr_errors_jis(text)
    elif preset == "DIN":
        return fix_common_ocr_errors_din(text)
    else:
        return fix_common_ocr_errors_jis(text)


# === FRAME PROCESSING ===
# Fungsi-fungsi untuk processing dan manipulasi frame dari kamera

def convert_frame_to_binary(frame):
    # Fungsi untuk konversi frame BGR ke binary (hitam putih) menggunakan OTSU threshold | Tujuan: Ubah gambar ke binary untuk processing lebih baik
    # Parameter: frame = numpy array BGR frame dari OpenCV
    # Return: numpy array frame binary dalam format BGR (3 channel untuk konsistensi)
    
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    binary_bgr = cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)
    return binary_bgr


def find_external_camera(max_cameras=5):
    # Fungsi mencari kamera eksternal yang berfungsi, prioritas ke kamera eksternal | Tujuan: Auto-detect kamera terbaik yang tersedia
    # Iterasi semua kamera dan prioritas: eksternal (indeks > 0) > internal (indeks 0)
    # Parameter: max_cameras = Maksimal indeks kamera yang akan di-check
    # Return: Integer indeks kamera yang berfungsi (preferensi eksternal)
    
    best_working_index = 0

    for i in range(max_cameras):
        cap = cv2.VideoCapture(i)
        if cap.isOpened():
            w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            cap.release()
            
            if w > 0 and h > 0:
                if i > 0:
                    return i
                else:
                    best_working_index = i
        
    return best_working_index


def create_directories():
    # Fungsi membuat direktori yang diperlukan jika belum ada | Tujuan: Setup folder untuk menyimpan file
    # Membuat folder untuk menyimpan gambar scan dan file Excel jika tidak ada
    
    from config import IMAGE_DIR, EXCEL_DIR
    os.makedirs(IMAGE_DIR, exist_ok=True)
    os.makedirs(EXCEL_DIR, exist_ok=True)


def cleanup_temp_files(temp_files_list):
    # Fungsi hapus file temporary dari sistem | Tujuan: Bersihkan temporary files untuk save disk space
    # Iterasi list file dan hapus satu per satu, handle exception jika file tidak ada
    # Parameter: temp_files_list = List of string paths untuk file yang akan dihapus
    
    for t_path in temp_files_list:
        if os.path.exists(t_path):
            try:
                os.remove(t_path)
            except:
                pass