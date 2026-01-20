# Logika deteksi OCR menggunakan EasyOCR dengan threading untuk live camera dan file scanning
# File ini berisi DetectionLogic class yang menangani OCR, frame processing, dan detection logic
# MODIFIED: 2-Stage Detection dengan STRUCTURAL CORRECTION untuk huruf tengah JIS dan (S) detection
# UPDATED: Binary mode diganti dengan Edge Detection mode

import cv2 #Import OpenCV untuk camera capture dan image processing
import easyocr #Import EasyOCR untuk optical character recognition
import re #Import regex untuk pattern matching dan text manipulation
import os #Import os untuk file/directory operations
import time #Import time untuk timing dan delay operations
import threading #Import threading untuk concurrent processing (camera + OCR)
import atexit #Import atexit untuk cleanup saat aplikasi exit
import numpy as np #Import numpy untuk array operations dan image manipulation
from datetime import datetime #Import datetime untuk timestamp handling
from difflib import SequenceMatcher #Import SequenceMatcher untuk fuzzy string matching
from PIL import Image #Import PIL Image untuk image processing

#Import konfigurasi dari config.py
from config import (
    IMAGE_DIR, EXCEL_DIR, DB_FILE, PATTERNS, ALLOWLIST_JIS, ALLOWLIST_DIN, DIN_TYPES,
    CAMERA_WIDTH, CAMERA_HEIGHT, TARGET_WIDTH, TARGET_HEIGHT, BUFFER_SIZE,
    MAX_CAMERAS, SCAN_INTERVAL, JIS_TYPES
)
#Import utility functions dari utils.py
from utils import (
    fix_common_ocr_errors, convert_frame_to_binary, find_external_camera,
    create_directories, apply_edge_detection
)
#Import database functions dari database.py
from database import (
    setup_database, load_existing_data, insert_detection
)

class DetectionLogic(threading.Thread):
    # Class utama untuk detection logic yang inherit dari Thread
    # Tujuan: Menjalankan camera capture dan OCR detection secara concurrent dengan UI
    
    def __init__(self, update_signal, code_detected_signal, camera_status_signal, data_reset_signal, all_text_signal=None):
        # Constructor untuk inisialisasi DetectionLogic
        # Parameter: berbagai signal untuk komunikasi dengan UI (PySide6 signals)
        # update_signal: untuk update preview frame
        # code_detected_signal: untuk notify saat code terdeteksi
        # camera_status_signal: untuk update status camera
        # data_reset_signal: untuk notify saat daily reset
        # all_text_signal: untuk debug/menampilkan semua text yang terdeteksi OCR
        
        super().__init__() #Call parent constructor (threading.Thread)
        
        # Store semua signal untuk komunikasi dengan UI thread
        self.update_signal = update_signal
        self.code_detected_signal = code_detected_signal
        self.camera_status_signal = camera_status_signal
        self.data_reset_signal = data_reset_signal
        self.all_text_signal = all_text_signal
        
        self.running = False #Flag untuk kontrol thread running state
        self.cap = None #VideoCapture object untuk camera (None saat init)
        self.preset = "JIS" #Preset default (JIS atau DIN)
        self.last_scan_time = 0 #Timestamp terakhir kali scan dilakukan (untuk throttling)
        self.scan_interval = SCAN_INTERVAL #Interval waktu antara scan (dalam detik)
        self.target_label = "" #Target label/sesi yang sedang aktif (untuk validasi OK/Not OK)
        
        create_directories() #Buat direktori untuk simpan gambar dan Excel jika belum ada
        
        self.current_camera_index = 0 #Index camera yang digunakan (0 = built-in, >0 = external)
        self.scan_lock = threading.Lock() #Lock untuk prevent concurrent OCR scan (hanya 1 scan at a time)
        self.temp_files_on_exit = [] #List untuk menyimpan temp files yang perlu dihapus saat exit
        
        # Flag untuk edge detection mode (menggantikan binary_mode)
        self.edge_mode = False  # CHANGED: dari binary_mode ke edge_mode
        self.split_mode = False #Flag untuk split mode (preview top=edge, bottom=original)    
        self.current_date = datetime.now().date() #Current date untuk daily reset check
        
        # Target display size untuk preview window
        self.TARGET_WIDTH = TARGET_WIDTH
        self.TARGET_HEIGHT = TARGET_HEIGHT
        self.patterns = PATTERNS #Regex patterns untuk detection (dari config)
        setup_database() #Setup database dan buat table jika belum ada
        self.detected_codes = load_existing_data(self.current_date) #Load data deteksi yang sudah ada untuk hari ini
        
        # Inisialisasi EasyOCR reader
        # ['en'] = English language, gpu=False untuk CPU processing
        # verbose=False untuk disable logging output
        self.reader = easyocr.Reader(['en'], gpu=True, verbose=False)

        atexit.register(self.cleanup_temp_files) #Register cleanup function untuk dipanggil saat aplikasi exit
    
    def cleanup_temp_files(self):
        # Fungsi untuk cleanup temporary files saat aplikasi exit
        # Tujuan: Hapus semua file temporary yang dibuat selama runtime
        # Loop setiap path dalam list temp files
        for t_path in self.temp_files_on_exit:
            # Check apakah file masih exist
            if os.path.exists(t_path):
                try:
                    os.remove(t_path) #Hapus file
                except:
                    pass #Silent fail jika ada error (file locked, permission denied, dll)
    
    def run(self):
        # Method utama thread yang akan dijalankan saat start()
        # Tujuan: Capture frame dari camera secara continuous dan trigger OCR scan
        # Cari external camera (prioritas eksternal, fallback ke built-in)
        self.current_camera_index = find_external_camera(MAX_CAMERAS)
        
        # Buka camera dengan DirectShow backend (Windows)
        # DirectShow biasanya lebih reliable untuk Windows
        self.cap = cv2.VideoCapture(self.current_camera_index + cv2.CAP_DSHOW)
        
        # Jika gagal dengan DirectShow, coba tanpa backend spesifik
        if not self.cap.isOpened():
            self.cap = cv2.VideoCapture(self.current_camera_index)

        # Jika masih gagal, emit error signal dan return
        if not self.cap.isOpened():
            self.camera_status_signal.emit(f"Error: Kamera Index {self.current_camera_index} Gagal Dibuka.", False)
            self.running = False
            return
        
        # Set camera buffer size untuk reduce latency
        try:
            self.cap.set(cv2.CAP_PROP_BUFFERSIZE, BUFFER_SIZE)
        except:
            pass #Silent fail jika property tidak supported
        
        # Set resolusi camera dan codec
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
        self.cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))  # MJPEG codec
        
        # Buat info string untuk display (External/Internal camera)
        camera_info = f"Camera: {self.current_camera_index} ({'External' if self.current_camera_index != 0 else 'Internal'})"
        self.camera_status_signal.emit(camera_info, True)

        # Main loop: capture frame selama running=True
        while self.running:
            ret, frame = self.cap.read() #Read frame dari camera
            
            # Jika gagal read frame, break loop
            if not ret:
                break

            self._process_and_send_frame(frame, is_static=False) #Process dan kirim frame ke UI untuk preview
            current_time = time.time() #Check apakah sudah waktunya untuk scan OCR
            
            # Jika sudah melewati scan_interval DAN tidak ada scan yang sedang berjalan
            if current_time - self.last_scan_time >= self.scan_interval and not self.scan_lock.locked():
                self.last_scan_time = current_time #Update last scan time
                # Jalankan OCR scan di thread terpisah (non-blocking)
                # daemon=True agar thread otomatis terminate saat main thread exit
                threading.Thread(target=self.scan_frame, 
                                args=(frame.copy(),),  # Copy frame untuk avoid race condition
                                kwargs={'is_static': False, 'original_frame': frame.copy()}, 
                                daemon=True).start()
        
        # Cleanup: release camera saat loop selesai
        if self.cap:
             self.cap.release()
        
        self.camera_status_signal.emit("Camera Off", False) #Emit signal camera off
    
    def _process_and_send_frame(self, frame, is_static):
        # Fungsi internal untuk process frame sebelum dikirim ke UI
        # Tujuan: Apply transformasi (crop, resize, edge detection, split mode) dan convert ke PIL Image
        # Parameter: frame (numpy array BGR), is_static (boolean untuk distinguish camera vs file)
    
        from PIL import Image #Import PIL Image untuk convert ke format yang bisa ditampilkan UI

        frame_display = frame.copy() #Copy frame untuk avoid modifying original

        # Jika bukan static file (live camera)
        if not is_static:
            h, w, _ = frame_display.shape #Get dimensi frame
            
            # Crop frame menjadi square (center crop)
            min_dim = min(h, w)
            start_x = (w - min_dim) // 2
            start_y = (h - min_dim) // 2
            frame_cropped = frame_display[start_y:start_y + min_dim, start_x:start_x + min_dim]

            # UPDATED: Edge Detection mode
            # Jika edge mode aktif, apply edge detection ke frame
            if self.edge_mode:
                frame_cropped = apply_edge_detection(frame_cropped)

            # Jika split mode aktif (preview top=edge, bottom=original)
            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2 #Hitung ukuran masing-masing bagian (setengah dari target height)
                
                # Resize frame ke ukuran konten
                frame_scaled_320 = cv2.resize(frame_cropped, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)

                frame_top_edge = apply_edge_detection(frame_scaled_320.copy()) #Apply edge detection untuk top frame
                frame_bottom_original = frame_scaled_320.copy() #Bottom frame tetap original (tidak di-edge)

                # Buat canvas kosong untuk top dan bottom
                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2 #Hitung offset untuk center horizontal

                # Paste frame ke canvas dengan offset (center)
                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_edge
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original
                frame_combined = np.vstack([canvas_top, canvas_bottom]) #Stack vertical: top di atas, bottom di bawah
                
                # Convert BGR ke RGB untuk PIL Image
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)

            else:
                # Mode normal (tidak split): langsung convert dan resize
                frame_rgb = cv2.cvtColor(frame_cropped, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                
                from config import Resampling #Import resampling method dari config
                
                img = img.resize((self.TARGET_WIDTH, self.TARGET_HEIGHT), Resampling) #Resize ke target size

        else:
            # Jika static file (bukan live camera)
            
            # UPDATED: Edge detection untuk static file
            # Apply edge detection jika mode aktif
            if self.edge_mode or self.split_mode:
                frame_display = apply_edge_detection(frame_display)
                
                # Tambahkan text overlay untuk info mode
                cv2.putText(frame_display, "Edge Detection Mode", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1, cv2.LINE_AA)

            # Convert BGR ke RGB
            frame_rgb = cv2.cvtColor(frame_display, cv2.COLOR_BGR2RGB)
            original_img = Image.fromarray(frame_rgb)
            original_width, original_height = original_img.size #Get dimensi original image
            # Hitung ratio untuk fit dalam target size (maintain aspect ratio)
            ratio = min(self.TARGET_WIDTH / original_width, self.TARGET_HEIGHT / original_height)

            # Hitung new size setelah scaling
            new_width = int(original_width * ratio)
            new_height = int(original_height * ratio)
            
            from config import Resampling #Import resampling method dari config
            
            img_resized = original_img.resize((new_width, new_height), Resampling) #Resize image dengan maintain aspect ratio
            img = Image.new('RGB', (self.TARGET_WIDTH, self.TARGET_HEIGHT), 'black') #Buat canvas hitam dengan target size
            
            # Hitung offset untuk center image di canvas
            x_offset = (self.TARGET_WIDTH - new_width) // 2
            y_offset = (self.TARGET_HEIGHT - new_height) // 2
            
            img.paste(img_resized, (x_offset, y_offset)) #Paste resized image ke center canvas

            # Tambahkan text overlay "STATIC FILE SCAN"
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(img)
            
            # Load font untuk text
            try:
                font = ImageFont.truetype("arial.ttf", 10)
            except IOError:
                font = ImageFont.load_default()
            
            text_to_display = "STATIC FILE SCAN" #Text yang akan ditampilkan

            # Hitung bounding box untuk center text horizontal
            bbox = draw.textbbox((0, 0), text_to_display, font=font)
            text_width = bbox[2] - bbox[0]
            x_center = (self.TARGET_WIDTH - text_width) // 2
            y_top = 12

            draw.text((x_center, y_top), text_to_display, fill=(255, 255, 0), font=font) #Draw text dengan warna kuning

        self.update_signal.emit(img) #Emit signal untuk update preview UI dengan PIL Image
    
    def _normalize_din_code(self, code):
        # Fungsi untuk normalisasi format DIN code
        # Tujuan: Standarisasi spacing dan format DIN code
        # Parameter: code (string) - raw DIN code
        # Return: normalized DIN code dengan spacing yang benar
        
        code = code.strip().upper() #Trim whitespace dan uppercase
        code_no_space = re.sub(r'\s+', '', code) #Hapus semua spasi untuk processing
        
        # Pattern 1: LBN + digit (contoh: LBN3 -> "LBN 3")
        match = re.match(r'^(LBN)(\d)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}"
        
        # Pattern 2: LN + digit + kapasitas + huruf (contoh: LN260A -> "LN2 60A")
        match = re.match(r'^(LN\d)(\d+)([A-Z])$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)}"
        
        # Pattern 3: Pattern 2 + ISS (contoh: LN260AISS -> "LN2 60A ISS")
        match = re.match(r'^(LN\d)(\d+)([A-Z])(ISS)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)} {match.group(4)}"
        
        return re.sub(r'\s+', ' ', code).strip() #Default: normalize multiple spaces menjadi single space
    
    def _find_best_din_match(self, detected_text):
        # Fungsi untuk mencari DIN type terbaik yang match dengan detected text
        # Tujuan: Fuzzy matching antara detected text dengan daftar DIN_TYPES
        # Parameter: detected_text (string) - text hasil OCR
        # Return: tuple (best_match, best_score) atau (None, 0.0) jika tidak ada match
        detected_normalized = self._normalize_din_code(detected_text) #Normalize detected text
        detected_clean = detected_normalized.replace(' ', '').upper() #Hapus spasi untuk comparison
        
        # Variabel untuk menyimpan match terbaik
        best_match = None
        best_score = 0.0
        
        # Loop setiap DIN type di config (skip index 0 yang biasanya placeholder)
        for din_type in DIN_TYPES[1:]:
            target_clean = din_type.replace(' ', '').upper() #Hapus spasi dari target untuk comparison
            
            # Hitung similarity ratio menggunakan SequenceMatcher
            # ratio() return nilai 0.0-1.0 (0=completely different, 1=identical)
            ratio = SequenceMatcher(None, detected_clean, target_clean).ratio()
            
            # Update best match jika ratio > 0.8 dan lebih baik dari score sebelumnya
            if ratio > 0.8 and ratio > best_score:
                best_score = ratio
                best_match = din_type
        
        return best_match, best_score #Return best match dan score
    
    def _correct_jis_structure(self, text):
        """
        KOREKSI STRUKTURAL JIS: Perbaiki huruf tengah yang salah terbaca sebagai angka
        Format JIS: [2-3 digit][1 HURUF A-H][2-3 digit][L/R optional][(S) optional]
        Tujuan: OCR sering salah baca huruf tengah (D/B/A) sebagai angka, fungsi ini koreksi struktur
        Parameter: text (string) - raw text dari OCR
        Return: corrected text dengan struktur JIS yang benar
        """
        # Trim dan uppercase text
        text = text.strip().upper().replace(' ', '')
        
        # Dictionary mapping angka yang sering salah dibaca ke huruf yang benar
        digit_to_letter = {
            '0': 'D', '1': 'I', '2': 'Z', '3': 'B', 
            '4': 'A', '5': 'S', '6': 'G', '8': 'B',
        }
        
        # Koreksi (S) option: berbagai variasi salah baca
        text = re.sub(r'\(5\)', r'(S)', text)  # (5) -> (S)
        text = re.sub(r'5\)', r'(S)', text)  # 5) -> (S)
        text = re.sub(r'\([S5](?!\))', r'(S)', text)  # (S atau (5 tanpa ) -> (S)
        
        # Pattern JIS: capacity(2-3 digit) + type(1 char) + size(2-3 digit) + terminal(L/R) + option((S))
        pattern = r'^(\d{2,3})([A-Z0-9])(\d{2,3})([LR])?(\(S\))?$'
        match = re.match(pattern, text) #Match pattern
        
        # Jika pattern cocok, lakukan koreksi struktural
        if match:
            # Extract setiap group
            capacity = match.group(1)  # 2-3 digit kapasitas
            middle_char = match.group(2)  # 1 karakter tengah (harus huruf A-H)
            size = match.group(3)  # 2-3 digit ukuran
            terminal = match.group(4) or ''  # L atau R (optional)
            option = match.group(5) or ''  # (S) (optional)
            
            # CRITICAL: Jika middle_char adalah digit, koreksi ke huruf
            if middle_char.isdigit():
                corrected_letter = digit_to_letter.get(middle_char, 'D') #Ambil huruf koreksi dari dictionary
                
                # Validasi: hanya gunakan jika huruf valid (A-H)
                if corrected_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
                    middle_char = corrected_letter
            
            # Gabungkan kembali dengan struktur yang benar
            corrected = f"{capacity}{middle_char}{size}{terminal}{option}"
            return corrected
  
        return text  #Jika pattern tidak cocok, return as is
    
    def _find_best_jis_match(self, detected_text):
        """
        TAHAP 2: Mencari match terbaik dari JIS_TYPES dengan (S) preserved
        Tujuan: Fuzzy matching dengan daftar JIS_TYPES, handle (S) option secara khusus
        Parameter: detected_text (string) - text hasil OCR
        Return: tuple (best_match, best_score) atau (None, 0.0) jika tidak ada match
        """
        # TAHAP 1: Koreksi struktural terlebih dahulu
        detected_corrected = self._correct_jis_structure(detected_text)
        detected_clean = detected_corrected.replace(' ', '').upper() #Hapus spasi untuk comparison
        
        # Variabel untuk menyimpan match terbaik
        best_match = None
        best_score = 0.0
        
        # Loop pertama: cari exact match atau high similarity (>0.85)
        for jis_type in JIS_TYPES[1:]:
            target_clean = jis_type.replace(' ', '').upper() #Hapus spasi dari target
            ratio = SequenceMatcher(None, detected_clean, target_clean).ratio() #Hitung similarity ratio
            # Update best match jika ratio > 0.85
            if ratio > 0.85 and ratio > best_score:
                best_score = ratio
                best_match = jis_type
        
        # Jika tidak ada match yang cukup baik ATAU score < 0.90
        # Coba matching tanpa (S) untuk handle case dimana (S) detection tidak reliable
        if not best_match or best_score < 0.90:
            detected_without_s = detected_clean.replace('(S)', '') #Hapus (S) dari detected untuk comparison
            
            # Loop untuk matching tanpa (S)
            for jis_type in JIS_TYPES[1:]:
                target_without_s = jis_type.replace(' ', '').replace('(S)', '').upper() #Hapus (S) dari target juga
                ratio = SequenceMatcher(None, detected_without_s, target_without_s).ratio() #Hitung similarity ratio tanpa (S)
                
                # Jika ratio > 0.90, punya match yang bagus
                if ratio > 0.90:
                    # Jika detected punya (S), tambahkan (S) ke hasil match
                    if '(S)' in detected_clean:
                        base_code = jis_type.replace('(S)', '')
                        candidate_with_s = base_code + '(S)'
                        
                        # Check apakah versi dengan (S) ada di JIS_TYPES
                        if candidate_with_s in JIS_TYPES:
                            best_match = candidate_with_s
                            best_score = ratio
                            break
                    else:
                        # Jika detected tidak punya (S), gunakan versi tanpa (S)
                        if '(S)' not in jis_type and ratio > best_score:
                            best_match = jis_type
                            best_score = ratio
        
        return best_match, best_score #Return best match dan score
    
    def scan_frame(self, frame, is_static=False, original_frame=None):
        """
        TAHAP 1: OCR mentah
        TAHAP 2: Structural correction + Fuzzy matching
        Tujuan: Main function untuk scan frame dan detect battery code
        Parameter: frame (numpy array), is_static (boolean), original_frame (untuk save)
        """
        # Variabel untuk menyimpan hasil match terbaik
        best_match = None
        
        # Frame yang akan disave: gunakan original jika ada, fallback ke frame
        frame_to_save = original_frame if original_frame is not None else frame
        
        # Jika bukan static file (live camera)
        if not is_static:
            # Try acquire lock (non-blocking), return jika sudah ada scan yang berjalan
            if not self.scan_lock.acquire(blocking=False):
                return
            
            # Crop frame menjadi square (same logic as _process_and_send_frame)
            h_orig, w_orig, _ = frame.shape
            min_dim_orig = min(h_orig, w_orig)
            start_x_orig = (w_orig - min_dim_orig) // 2
            start_y_orig = (h_orig - min_dim_orig) // 2
            frame = frame[start_y_orig:start_y_orig + min_dim_orig, start_x_orig:start_x_orig + min_dim_orig]
            
            # UPDATED: Edge detection mode
            # Apply edge detection jika mode aktif
            if self.edge_mode:
                frame = apply_edge_detection(frame)
            
            # Jika split mode aktif
            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2
                frame_scaled_320 = cv2.resize(frame, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)

                # UPDATED: Apply edge detection untuk top frame
                frame_top_edge = apply_edge_detection(frame_scaled_320.copy())
                frame_bottom_original = frame_scaled_320.copy()
                
                # Buat canvas untuk top dan bottom
                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                
                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2
                
                # Paste frame ke canvas
                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_edge
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original
                
                # Stack vertical
                frame_combined = np.vstack([canvas_top, canvas_bottom])
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)

        try:
            # Get dimensi frame
            h, w = frame.shape[:2]
            
            # Resize frame jika terlalu besar (max width 640 untuk speed up OCR)
            if w > 640:
                scale = 640 / w
                new_w, new_h = 640, int(h * scale)
                frame_small = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
            else:
                frame_small = frame
            
            gray = cv2.cvtColor(frame_small, cv2.COLOR_BGR2GRAY) # Convert frame ke grayscale untuk preprocessing
            processing_stages = {} # Dictionary untuk menyimpan berbagai preprocessing stages
            
            # Preprocessing berbeda untuk DIN vs JIS
            if self.preset == "DIN":
                # DIN: Gunakan CLAHE (Contrast Limited Adaptive Histogram Equalization)
                clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
                enhanced = clahe.apply(gray)
                processing_stages['Enhanced'] = enhanced
                
                # Binary threshold dengan Otsu
                _, binary1 = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                processing_stages['Binary_Otsu'] = binary1
                
                # Adaptive threshold Gaussian
                adaptive1 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Gaussian'] = adaptive1
                
                # Adaptive threshold Mean
                adaptive2 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Mean'] = adaptive2
                
                # Inverted versions (untuk handle white text on dark background)
                processing_stages['Binary_Inv'] = cv2.bitwise_not(binary1)
                processing_stages['Adaptive_Inv'] = cv2.bitwise_not(adaptive1)
                
                # Morphological closing untuk connect broken characters
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
                morph = cv2.morphologyEx(binary1, cv2.MORPH_CLOSE, kernel)
                processing_stages['Morphed'] = morph
                
            else:
                # JIS: Preprocessing lebih simple
                # Sharpening kernel untuk enhance edges
                kernel = np.array([[-1,-1,-1], [-1, 9,-1],[-1,-1,-1]])
                processing_stages['Sharpened'] = cv2.filter2D(gray, -1, kernel)
                processing_stages['Grayscale'] = gray #Grayscale original
                processing_stages['Inverted_Gray'] = cv2.bitwise_not(gray) #Inverted grayscale
                # Binary adaptive threshold inverted
                processed_frame_binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                processing_stages['Binary'] = processed_frame_binary

            all_results = [] #List untuk menyimpan semua hasil OCR dari berbagai preprocessing
            
            # Pilih allowlist characters sesuai preset
            if self.preset == "JIS":
                allowlist_chars = ALLOWLIST_JIS  # Hanya karakter valid untuk JIS
            else:
                allowlist_chars = ALLOWLIST_DIN  # Hanya karakter valid untuk DIN

            # Loop setiap preprocessing stage dan jalankan OCR
            for stage_name, processed_frame in processing_stages.items():
                try:
                    # Jalankan EasyOCR readtext
                    # detail=0: return hanya text tanpa bounding box dan confidence
                    # paragraph: True untuk DIN (text bisa multi-line), False untuk JIS (single line)
                    # min_size: minimum size detection box
                    # width_ths: threshold untuk grouping text
                    # allowlist: karakter yang diperbolehkan
                    results = self.reader.readtext(
                        processed_frame, 
                        detail=0,
                        paragraph=False if self.preset == "JIS" else True,
                        min_size=10 if self.preset == "JIS" else 15,
                        width_ths=0.7 if self.preset == "JIS" else 0.5,
                        allowlist=allowlist_chars
                    )
                    # Tambahkan hasil ke all_results
                    all_results.extend(results)
                except Exception as e:
                    # Print error tapi lanjut ke stage berikutnya
                    print(f"OCR error on {stage_name}: {e}")
                    continue

            # Emit semua text yang terdeteksi untuk debugging (jika signal ada)
            if self.all_text_signal:
                # Ambil unique results (hapus duplikat)
                unique_results = list(set(all_results))
                self.all_text_signal.emit(unique_results)
            current_preset = self.preset #Simpan current preset untuk konsistensi
            
            # MATCHING LOGIC: berbeda untuk DIN vs JIS
            if current_preset == "DIN":
                # DIN: Cari match terbaik dari semua OCR results
                best_match_text = None
                best_match_score = 0.0
                
                # Loop setiap text hasil OCR
                for text in all_results:
                    # Apply OCR error correction
                    text_fixed = fix_common_ocr_errors(text, self.preset)
                    
                    # Skip jika text terlalu pendek (< 4 char tanpa spasi)
                    if len(text_fixed.replace(' ', '')) < 4:
                        continue
                    
                    # Cari best DIN match dengan fuzzy matching
                    matched_type, score = self._find_best_din_match(text_fixed)
                    
                    # Update best match jika score lebih baik
                    if matched_type and score > best_match_score:
                        best_match_score = score
                        best_match_text = matched_type
                
                # Jika ada match dengan score > 0.8, gunakan itu
                if best_match_text and best_match_score > 0.8:
                    best_match = best_match_text
                    
            else:
                # JIS: Cari match terbaik dari semua OCR results
                best_match_text = None
                best_match_score = 0.0
                
                # Loop setiap text hasil OCR
                for text in all_results:
                    # Skip jika text terlalu pendek (< 5 char tanpa spasi dan (S))
                    if len(text.replace(' ', '').replace('(S)', '')) < 5:
                        continue
                    
                    # Cari best JIS match dengan structural correction + fuzzy matching
                    matched_type, score = self._find_best_jis_match(text)
                    
                    # Update best match jika score lebih baik
                    if matched_type and score > best_match_score:
                        best_match_score = score
                        best_match_text = matched_type
                
                # Jika ada match dengan score > 0.85, gunakan itu
                if best_match_text and best_match_score > 0.85:
                    best_match = best_match_text
            
            # Jika ada match yang ditemukan
            if best_match:
                detected_code = best_match.strip() # Clean detected code
                
                # Normalize DIN code (add proper spacing)
                if self.preset == "DIN":
                    detected_code = self._normalize_din_code(detected_code)
                else:
                    # JIS: hapus spasi
                    detected_code = detected_code.replace(' ', '')

                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S") #Get timestamp untuk database
                detected_type = self._detect_code_type(detected_code) #Detect tipe code (JIS/DIN) untuk validasi
                # Validasi apakah detected type match dengan preset
                is_valid_type, error_message = self._validate_preset_match(detected_code, detected_type)
                
                # Jika tidak valid, emit error dan return
                if not is_valid_type:
                    self.code_detected_signal.emit(error_message)
                    if not is_static:
                        self.scan_lock.release()
                    return
                
                # Tentukan status OK/Not OK berdasarkan match dengan target_label
                if self.preset == "DIN":
                    # DIN: normalize kedua code untuk comparison
                    target_normalized = self._normalize_din_code(self.target_label)
                    detected_normalized = self._normalize_din_code(detected_code)
                    status = "OK" if detected_normalized.upper() == target_normalized.upper() else "Not OK"
                else:
                    # JIS: exact string match
                    status = "OK" if detected_code == self.target_label else "Not OK"
                
                # Set target_session: gunakan target_label jika ada, fallback ke detected_code
                target_session = self.target_label if self.target_label else detected_code

                # Prevent duplicate detection (dalam 5 detik terakhir)
                if not is_static:
                    # Check apakah code ini sudah terdeteksi dalam 5 detik terakhir
                    if any(rec["Code"] == detected_code and 
                           (datetime.now() - datetime.strptime(rec["Time"], "%Y-%m-%d %H:%M:%S")).total_seconds() < 5
                           for rec in self.detected_codes):
                        return
                
                # Generate filename untuk save image
                img_filename = f"karton_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                img_path = os.path.join(IMAGE_DIR, img_filename)
                frame_binary = convert_frame_to_binary(frame_to_save) #Convert frame ke binary/edge detection untuk save 
                cv2.imwrite(img_path, frame_binary) #Save image ke disk
                #Insert detection ke database
                new_id = insert_detection(timestamp, detected_code, current_preset, img_path, status, target_session)

                # Jika insert berhasil (dapat ID baru)
                if new_id:
                    # Buat record dictionary untuk local list
                    record = {
                        "ID": new_id,
                        "Time": timestamp,
                        "Code": detected_code,
                        "Type": current_preset,
                        "ImagePath": img_path,
                        "Status": status,
                        "TargetSession": target_session
                    }
                    
                    self.detected_codes.append(record) #Append ke local detected_codes list

                self.code_detected_signal.emit(detected_code) #Emit signal code detected ke UI
                
            else:
                # Jika tidak ada match dan ini static file scan
                if is_static:
                    self.code_detected_signal.emit("FAILED") #Emit signal FAILED

        except Exception as e:
            print(f"OCR/Regex error: {e}") #Jika terjadi error di OCR/processing
            # Emit error message untuk static scan
            if is_static:
                self.code_detected_signal.emit(f"ERROR: {e}")
                
        finally:
            # Release lock jika bukan static scan
            if not is_static:
                self.scan_lock.release()
    
    def start_detection(self):
        # Fungsi untuk start detection thread
        # Tujuan: Mulai live camera detection
        # Jika sudah running, return (prevent duplicate start)
        if self.running:
            return 
        self.running = True # Set flag running True
        self.start() # Start thread (akan call method run())

    def stop_detection(self):
        # Fungsi untuk stop detection thread
        # Tujuan: Hentikan live camera detection    
        self.running = False # Set flag running False (akan stop loop di run())
        
        # Release camera jika ada
        if self.cap:
             self.cap.release()
             
    def set_camera_options(self, preset, flip_h, flip_v, edge_mode, split_mode, scan_interval):
        # Fungsi untuk set camera options dari UI
        # Tujuan: Update settings deteksi (preset, flip, mode, interval)
        # Parameter: preset (JIS/DIN), flip_h/v (boolean), edge_mode (boolean), split_mode (boolean), scan_interval (seconds)
        
        self.preset = preset  # Set preset battery type
        self.flip_h = flip_h  # Set horizontal flip (tidak digunakan di code saat ini)
        self.flip_v = flip_v  # Set vertical flip (tidak digunakan di code saat ini)
        self.edge_mode = edge_mode  # CHANGED: dari binary_mode ke edge_mode
        self.split_mode = split_mode  # Set split preview mode
        self.scan_interval = scan_interval  # Set interval scan OCR
    
    def set_target_label(self, label):
        # Fungsi untuk set target label/sesi
        # Tujuan: Set label yang sedang dideteksi untuk validasi OK/Not OK
        # Parameter: label (string) - target label
        self.target_label = label

    def check_daily_reset(self):
        # Fungsi untuk check apakah sudah ganti hari (daily reset)
        # Tujuan: Reset data detected_codes setiap ganti hari
        # Return: Boolean True jika terjadi reset, False jika tidak
        
        # Get current datetime
        now = datetime.now()
        new_date = now.date()
        
        # Compare dengan current_date yang tersimpan
        if new_date > self.current_date:
            self.current_date = new_date # Ganti hari terdeteksi
            self.detected_codes = [] # Clear local detected_codes list
            self.detected_codes = load_existing_data(self.current_date) # Load data untuk tanggal baru
            self.data_reset_signal.emit() # Emit signal daily reset ke UI
            
            return True
        
        return False
        
    def scan_file(self, filepath):
        # Fungsi untuk scan static image file
        # Tujuan: OCR detection dari file gambar (bukan live camera)
        # Parameter: filepath (string) - path ke image file
        # Return: String status code
        
        # Check apakah live camera sedang running
        if self.running: 
            return "STOP_LIVE"  # User harus stop camera dulu
        
        try:
            frame = cv2.imread(filepath) #Load image file dengan OpenCV
            # Validasi image berhasil loaded
            if frame is None or frame.size == 0:
                return "LOAD_ERROR"
            self._process_and_send_frame(frame, is_static=True) #Process dan tampilkan frame ke UI
            
            # Jalankan OCR scan di thread terpisah
            # is_static=True untuk distinguish dari live camera
            threading.Thread(target=self.scan_frame,
                            args=(frame.copy(),),
                            kwargs={'is_static': True, 'original_frame': frame.copy()},
                            daemon=True).start()
            
            # Return status scanning
            return "SCANNING"
            
        except Exception as e:
            # Print error dan return error message
            print(f"File scan error: {e}")
            return f"PROCESS_ERROR: {e}"

    def _detect_code_type(self, code):
        # Fungsi untuk detect tipe code (JIS/DIN) dari string
        # Tujuan: Auto-detect apakah code termasuk JIS atau DIN format
        # Parameter: code (string) - battery code
        # Return: "JIS", "DIN", atau None jika tidak dikenali
        code_normalized = code.replace(' ', '').upper() # Normalize code (hapus spasi, uppercase)
        
        # Check JIS pattern: 2-3 digit + huruf A-H + 2-3 digit + optional L/R + optional (S)
        jis_pattern = r"\b\d{2,3}[A-H]\d{2,3}[LR]?(?:\(S\))?\b"
        if re.match(jis_pattern, code_normalized):
            return "JIS"
        
        # Check exact match dengan DIN_TYPES list
        for din_type in DIN_TYPES[1:]:
            if code_normalized == din_type.replace(' ', '').upper():
                return "DIN"
        
        # Check DIN patterns dengan regex
        din_patterns = [
            r'^LBN\d$', #LBN + 1 digit
            r'^LN\d\d{2,3}A(?:ISS)?$', #LN + digit + kapasitas + A + optional ISS
        ]
        
        # Loop setiap DIN pattern
        for pattern in din_patterns:
            if re.match(pattern, code_normalized):
                return "DIN"
    
        return None # Return None jika tidak match dengan pattern manapun

    def _validate_preset_match(self, detected_code, detected_type):
        # Fungsi untuk validasi apakah detected type match dengan preset
        # Tujuan: Prevent salah deteksi (user pilih JIS tapi scan DIN, atau sebaliknya)
        # Parameter: detected_code (string), detected_type (JIS/DIN/None)
        # Return: tuple (is_valid, error_message)
        
        # Jika type tidak terdeteksi (None)
        if detected_type is None:
            return False, "Format kode tidak valid"
        
        # Jika detected type tidak match dengan preset yang dipilih user
        if detected_type != self.preset:
            # Return error message sesuai preset
            if self.preset == "JIS":
                return False, "Pastikan foto anda adalah Type JIS"
            else:
                return False, "Pastikan foto anda adalah Type DIN"
        
        # Valid: type match dengan preset
        return True, ""

    def delete_codes(self, record_ids):
        # Fungsi untuk delete deteksi berdasarkan ID
        # Tujuan: Hapus record dari database dan local list
        # Parameter: record_ids (list of int) - ID yang akan dihapus
        # Return: Boolean success/failure 
        from database import delete_codes # Import delete_codes function dari database module
        
        # Call database delete function
        if delete_codes(record_ids):
            # Jika berhasil, hapus juga dari local detected_codes list
            # List comprehension: keep hanya record yang ID-nya TIDAK di record_ids
            self.detected_codes = [rec for rec in self.detected_codes if rec['ID'] not in record_ids]
            return True
        
        return False