"""
Fungsi utility dan helper yang digunakan di berbagai modul.
UPDATED: Edge Detection dengan garis putih neon terang pada background hitam pekat
"""

import os
import re
import numpy as np
import cv2
from datetime import datetime
from config import Resampling

# === EDGE DETECTION ===
# Fungsi untuk Edge Detection dengan background hitam pekat dan garis putih neon terang

def apply_edge_detection(frame):
    """
    Fungsi untuk menerapkan Edge Detection pada frame dengan garis putih neon terang
    Tujuan: Deteksi tepi objek dan teks dengan background HITAM PEKAT MURNI dan garis putih neon
    Parameter: frame = numpy array BGR frame dari OpenCV
    Return: numpy array frame edge detection dalam format BGR (3 channel)
    
    Menggunakan algoritma Canny Edge Detection dengan background pure black (0,0,0)
    """
    # Konversi ke grayscale
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    
    # Apply Gaussian Blur untuk mengurangi noise
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    
    # Canny Edge Detection dengan threshold lebih rendah untuk lebih banyak detail
    edges = cv2.Canny(blurred, 30, 100)
    
    # Dilate edges untuk membuat garis lebih tebal dan terang
    kernel = np.ones((2, 2), np.uint8)
    edges_dilated = cv2.dilate(edges, kernel, iterations=1)
    
    # CRITICAL: Buat canvas HITAM MURNI (pure black) terlebih dahulu
    # Ini memastikan background benar-benar hitam pekat tanpa noise abu-abu
    edges_bgr = np.zeros((edges_dilated.shape[0], edges_dilated.shape[1], 3), dtype=np.uint8)
    
    # HANYA gambar garis putih pada pixel yang terdeteksi sebagai edge
    # Background tetap hitam murni (0,0,0), hanya edge yang jadi putih (255,255,255)
    edges_bgr[edges_dilated > 0] = [255, 255, 255]  # BGR: Putih murni hanya di edge
    
    # OPTIONAL: Tingkatkan brightness garis putih jika perlu lebih terang
    # Uncomment baris di bawah untuk garis lebih glowing
    # edges_bgr[edges_dilated > 0] = [255, 255, 255]
    
    return edges_bgr


# === OCR CORRECTION LOGIC ===

def fix_common_ocr_errors_jis(text):
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
    text = text.strip().upper()
    
    char_to_digit = {
        'O': '0', 'Q': '0',
        'I': '1', 'l': '1',
        'Z': '2',
        'S': '5',
        'G': '6',
        'B': '8',
    }
    
    digit_to_char = {
        '0': 'O',
        '1': 'I',
    }
    
    text = re.sub(r'[^A-Z0-9\s]', '', text)
    
    tokens = text.split()
    
    if len(tokens) == 0:
        return text
    
    corrected_tokens = []
    
    for i, token in enumerate(tokens):
        if i == 0:
            corrected = ""
            for j, char in enumerate(token):
                if j == 0:
                    if char in ['I', '1', 'l']:
                        corrected += 'L'
                    else:
                        corrected += char
                elif j == 1:
                    if char in ['8']:
                        corrected += 'B'
                    elif char in ['H', 'M']:
                        corrected += 'N'
                    else:
                        corrected += char
                elif j == 2:
                    if token[:2] == "LB":
                        if char in ['H', 'M']:
                            corrected += 'N'
                        else:
                            corrected += char
                    else:
                        if char in char_to_digit:
                            corrected += char_to_digit[char]
                        else:
                            corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
            
        elif i == 1:
            corrected = ""
            for j, char in enumerate(token):
                if char.isdigit():
                    corrected += char
                elif char in char_to_digit:
                    corrected += char_to_digit[char]
                elif j == len(token) - 1 and char.isalpha():
                    if char in ['4', 'H']:
                        corrected += 'A'
                    else:
                        corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
            
        elif i == 2:
            corrected = token
            if token in ['I55', 'IS5', 'I5S', '155']:
                corrected = 'ISS'
            elif token.replace('5', 'S').replace('I', 'I') == 'ISS':
                corrected = 'ISS'
            
            corrected_tokens.append(corrected)
    
    result = ' '.join(corrected_tokens)
    
    result = re.sub(r'(LBN)(\d)', r'\1 \2', result)
    result = re.sub(r'(LN\d)(\d)', r'\1 \2', result)
    result = re.sub(r'([A-Z0-9])(ISS)', r'\1 \2', result)
    result = re.sub(r'\s+', ' ', result)
    
    return result.strip()


def fix_common_ocr_errors(text, preset):
    if preset == "JIS":
        return fix_common_ocr_errors_jis(text)
    elif preset == "DIN":
        return fix_common_ocr_errors_din(text)
    else:
        return fix_common_ocr_errors_jis(text)


# === FRAME PROCESSING ===

def convert_frame_to_binary(frame):
    """
    UPDATED: Gunakan edge detection dengan putih neon untuk export Excel
    """
    # Gunakan edge detection dengan garis putih neon pada background hitam
    return apply_edge_detection(frame)


def find_external_camera(max_cameras=5):
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
    from config import IMAGE_DIR, EXCEL_DIR
    os.makedirs(IMAGE_DIR, exist_ok=True)
    os.makedirs(EXCEL_DIR, exist_ok=True)


def cleanup_temp_files(temp_files_list):
    for t_path in temp_files_list:
        if os.path.exists(t_path):
            try:
                os.remove(t_path)
            except:
                pass