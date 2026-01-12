# Logika deteksi OCR menggunakan EasyOCR dengan threading untuk live camera dan file scanning
# File ini berisi DetectionLogic class yang menangani OCR, frame processing, dan detection logic

import cv2  # Computer vision | Library untuk video capture dan image processing
import easyocr  # OCR engine | Library untuk optical character recognition
import re  # Regular expressions | Modul untuk pattern matching
import os  # File operations | Modul untuk file system operations
import time  # Time operations | Modul untuk time tracking
import threading  # Threading | Modul untuk multi-threading operations
import atexit  # Exit handler | Modul untuk cleanup saat exit
import numpy as np  # Numerical library | Library untuk numerical operations
from datetime import datetime  # Date/time | Modul untuk date/time operations
from config import (
    IMAGE_DIR, EXCEL_DIR, DB_FILE, PATTERNS, ALLOWLIST_JIS, ALLOWLIST_DIN, DIN_TYPES,
    CAMERA_WIDTH, CAMERA_HEIGHT, TARGET_WIDTH, TARGET_HEIGHT, BUFFER_SIZE,
    MAX_CAMERAS, SCAN_INTERVAL #Mengambil pengaturan dari config.py
)
from utils import (
    fix_common_ocr_errors, convert_frame_to_binary, find_external_camera,
    create_directories #Mengambil fungsi dari utilitas.py
)
from database import (
    setup_database, load_existing_data, insert_detection # Mengambil fungsi database.py
)
from PIL import Image  # Image processing | Library untuk image manipulation atau membuka/menyimpan/resize/crop/rotate/flip/optimasi kebutuhan ocr/excel.


class DetectionLogic(threading.Thread):
    # Class untuk semua logika deteksi | Tujuan: Handle camera capture, OCR, frame processing, dan database operations di thread terpisah
    
    def __init__(self, update_signal, code_detected_signal, camera_status_signal, data_reset_signal, all_text_signal=None):
        # Fungsi inisialisasi | Tujuan: Setup thread dan inisialisasi semua variabel untuk deteksi
        super().__init__()  # Inisialisasi parent class threading.Thread
        
        # Signals untuk komunikasi dengan UI thread
        self.update_signal = update_signal  # Signal untuk update frame display
        self.code_detected_signal = code_detected_signal  # Signal ketika kode terdeteksi
        self.camera_status_signal = camera_status_signal  # Signal untuk status kamera
        self.data_reset_signal = data_reset_signal  # Signal untuk reset data harian
        self.all_text_signal = all_text_signal  # Signal untuk mengirim semua text hasil OCR
        
        # Thread control variables
        self.running = False  # Flag untuk kontrol loop thread
        self.cap = None  # Object VideoCapture untuk kamera
        self.preset = "JIS"  # Default preset untuk OCR (JIS/DIN)
        self.last_scan_time = 0  # Timestamp scan terakhir untuk throttling
        self.scan_interval = SCAN_INTERVAL  # Interval waktu antar scan
        
        self.target_label = ""  # Label target untuk validasi OK/Not OK
        
        create_directories()  # Buat direktori yang diperlukan (IMAGE_DIR, EXCEL_DIR)
        
        self.current_camera_index = 0  # Index kamera yang digunakan
        self.scan_lock = threading.Lock()  # Lock untuk mencegah scan bersamaan
        self.temp_files_on_exit = []  # List file temporary untuk cleanup
        
        self.binary_mode = False  # Mode binary/grayscale | Flag untuk toggle binary color mode
        self.split_mode = False  # Mode split screen | Flag untuk toggle split screen mode
        
        self.current_date = datetime.now().date()  # Tanggal saat ini untuk reset harian
        
        # Ukuran target display
        self.TARGET_WIDTH = TARGET_WIDTH  # Lebar target display
        self.TARGET_HEIGHT = TARGET_HEIGHT  # Tinggi target display
        
        self.patterns = PATTERNS  # Regex patterns untuk validasi kode
        
        setup_database()  # Setup database dan buat tabel jika belum ada
        self.detected_codes = load_existing_data(self.current_date)  # Load data yang sudah ada untuk hari ini
        
        # Inisialisasi EasyOCR reader dengan bahasa English, tanpa GPU
        self.reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        
        atexit.register(self.cleanup_temp_files)  # Register fungsi cleanup saat aplikasi ditutup
    
    def cleanup_temp_files(self):
        # Hapus file temporary saat aplikasi ditutup | Tujuan: Cleanup temporary files on exit
        for t_path in self.temp_files_on_exit:  # Loop semua file temporary
            if os.path.exists(t_path):  # Cek apakah file masih ada
                try:
                    os.remove(t_path)  # Hapus file
                except:
                    pass  # Ignore error jika gagal hapus
    
    def run(self):
        # Fungsi utama thread untuk Live Camera Loop | Tujuan: Main thread loop - capture frame dan process terus-menerus
        self.current_camera_index = find_external_camera(MAX_CAMERAS)  # Cari external camera, fallback ke 0 jika tidak ada
        
        # Coba buka kamera dengan DirectShow backend (Windows)
        self.cap = cv2.VideoCapture(self.current_camera_index + cv2.CAP_DSHOW)
        
        # Jika gagal dengan DirectShow, coba tanpa backend
        if not self.cap.isOpened():
            self.cap = cv2.VideoCapture(self.current_camera_index)

        # Jika masih gagal, emit error signal dan keluar
        if not self.cap.isOpened():
            self.camera_status_signal.emit(f"Error: Kamera Index {self.current_camera_index} Gagal Dibuka.", False)
            self.running = False
            return
        
        # Set buffer size untuk mengurangi latency
        try:
            self.cap.set(cv2.CAP_PROP_BUFFERSIZE, BUFFER_SIZE)
        except:
            pass  # Ignore jika gagal set buffer
            
        # Set resolusi kamera
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
        self.cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))  # Set codec MJPG untuk kualitas lebih baik
        
        # Emit camera status dengan info index dan tipe (External/Internal)
        camera_info = f"Camera: {self.current_camera_index} ({'External' if self.current_camera_index != 0 else 'Internal'})"
        self.camera_status_signal.emit(camera_info, True)

        # Main loop - capture dan process frame terus-menerus
        while self.running:
            ret, frame = self.cap.read()  # Baca frame dari kamera
            if not ret:  # Jika gagal baca frame
                break  # Keluar dari loop

            # Process dan kirim frame ke UI untuk display
            self._process_and_send_frame(frame, is_static=False)

            # Throttling - scan OCR hanya jika sudah lewat interval waktu
            current_time = time.time()
            if current_time - self.last_scan_time >= self.scan_interval and not self.scan_lock.locked():
                self.last_scan_time = current_time  # Update timestamp scan terakhir
                # Jalankan scan di thread terpisah agar tidak blocking main loop
                threading.Thread(target=self.scan_frame, 
                                args=(frame.copy(),),  # Copy frame untuk avoid race condition
                                kwargs={'is_static': False, 'original_frame': frame.copy()}, 
                                daemon=True).start()
        
        # Cleanup - release kamera saat loop berhenti
        if self.cap:
             self.cap.release()
        self.camera_status_signal.emit("Camera Off", False)  # Emit status kamera off
    
    def _process_and_send_frame(self, frame, is_static):
        # Fungsi process frame dengan berbagai filter | Tujuan: Apply flip/binary/split processing dan send ke UI
        from PIL import Image
        
        frame_display = frame.copy()  # Copy frame untuk processing
        
        # Processing untuk live camera feed
        if not is_static:
            h, w, _ = frame_display.shape  # Ambil dimensi frame
            min_dim = min(h, w)  # Ambil dimensi terkecil untuk crop square
            start_x = (w - min_dim) // 2  # Hitung offset x untuk center crop
            start_y = (h - min_dim) // 2  # Hitung offset y untuk center crop
            frame_cropped = frame_display[start_y:start_y + min_dim, start_x:start_x + min_dim]  # Crop ke square
            
            # Apply binary mode jika aktif
            if self.binary_mode:
                gray_frame = cv2.cvtColor(frame_cropped, cv2.COLOR_BGR2GRAY)  # Convert ke grayscale
                _, frame_cropped = cv2.threshold(gray_frame, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  # Apply Otsu threshold
                frame_cropped = cv2.cvtColor(frame_cropped, cv2.COLOR_GRAY2BGR)  # Convert kembali ke BGR

            # Apply split mode jika aktif (atas binary, bawah original)
            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2  # Ukuran setengah untuk split
                frame_scaled_320 = cv2.resize(frame_cropped, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)  # Resize frame

                # Bagian atas - binary version
                gray_top = cv2.cvtColor(frame_scaled_320.copy(), cv2.COLOR_BGR2GRAY)
                _, frame_top_binary = cv2.threshold(gray_top, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                frame_top_binary = cv2.cvtColor(frame_top_binary, cv2.COLOR_GRAY2BGR)
                
                # Bagian bawah - original version
                frame_bottom_original = frame_scaled_320.copy()
                
                # Buat canvas hitam untuk masing-masing bagian
                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3))
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3))
                
                # Center frame dalam canvas
                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2
                
                # Paste frame ke canvas
                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_binary
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original
                            
                # Stack vertikal atas dan bawah
                frame_combined = np.vstack([canvas_top, canvas_bottom])
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)  # Convert ke RGB untuk PIL
                img = Image.fromarray(frame_rgb)  # Convert ke PIL Image

            else:  # Mode normal (tidak split)
                frame_rgb = cv2.cvtColor(frame_cropped, cv2.COLOR_BGR2RGB)  # Convert ke RGB
                img = Image.fromarray(frame_rgb)  # Convert ke PIL Image
                from config import Resampling
                img = img.resize((self.TARGET_WIDTH, self.TARGET_HEIGHT), Resampling)  # Resize ke target size

        # Processing untuk static file
        else:
            # Apply binary mode jika aktif untuk static file
            if self.binary_mode or self.split_mode:
                gray_frame = cv2.cvtColor(frame_display, cv2.COLOR_BGR2GRAY)  # Convert ke grayscale
                _, frame_display = cv2.threshold(gray_frame, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  # Apply Otsu threshold
                frame_display = cv2.cvtColor(frame_display, cv2.COLOR_GRAY2BGR)  # Convert kembali ke BGR
                cv2.putText(frame_display, "Binary Mode", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1, cv2.LINE_AA)  # Add text overlay
            
            # Convert frame ke PIL Image
            frame_rgb = cv2.cvtColor(frame_display, cv2.COLOR_BGR2RGB)
            original_img = Image.fromarray(frame_rgb)
            
            # Maintain aspect ratio saat resize
            original_width, original_height = original_img.size
            ratio = min(self.TARGET_WIDTH / original_width, self.TARGET_HEIGHT / original_height)  # Hitung ratio untuk maintain aspect
            
            # Hitung dimensi baru dengan aspect ratio
            new_width = int(original_width * ratio)
            new_height = int(original_height * ratio)
            
            from config import Resampling
            img_resized = original_img.resize((new_width, new_height), Resampling)  # Resize dengan aspect ratio
            
            # Buat canvas hitam dan center image
            img = Image.new('RGB', (self.TARGET_WIDTH, self.TARGET_HEIGHT), 'black')
            x_offset = (self.TARGET_WIDTH - new_width) // 2  # Hitung offset x untuk center
            y_offset = (self.TARGET_HEIGHT - new_height) // 2  # Hitung offset y untuk center
            img.paste(img_resized, (x_offset, y_offset))  # Paste image di center canvas
            
            # Add "STATIC FILE SCAN" text overlay
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("arial.ttf", 10)  # Load Arial font
            except IOError:
                font = ImageFont.load_default()  # Fallback ke default font
                
            text_to_display = "STATIC FILE SCAN"
            bbox = draw.textbbox((0, 0), text_to_display, font=font)  # Get text bounding box
            text_width = bbox[2] - bbox[0]  # Hitung lebar text
            x_center = (self.TARGET_WIDTH - text_width) // 2  # Center text horizontal
            y_top = 12  # Posisi vertikal dari atas
            draw.text((x_center, y_top), text_to_display, fill=(255, 255, 0), font=font)  # Draw text kuning

        self.update_signal.emit(img)  # Emit signal untuk update display di UI
    
    def _normalize_din_code(self, code):
        # Normalisasi format DIN code untuk konsistensi | Tujuan: Normalize DIN code format
        code = code.strip().upper()  # Trim whitespace dan uppercase
        code_no_space = re.sub(r'\s+', '', code)  # Hapus semua whitespace
        
        # Coba match dengan berbagai pattern DIN
        # Pattern 1: LBN + angka (e.g., LBN1, LBN2, LBN3)
        match = re.match(r'^(LBN)(\d)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}"  # Format: "LBN 1"
        
        # Pattern 2: LN + digit + spasi + angka + A (e.g., LN1450A, LN0260A)
        match = re.match(r'^(LN\d)(\d+)([A-Z])$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)}"  # Format: "LN1 450A"
        
        # Pattern 3: LN + digit + spasi + angka + A + spasi + ISS (e.g., LN4776AISS)
        match = re.match(r'^(LN\d)(\d+)([A-Z])(ISS)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)} {match.group(4)}"  # Format: "LN4 776A ISS"
        
        # Jika sudah dalam format yang benar, kembalikan dengan spasi dinormalisasi
        return re.sub(r'\s+', ' ', code).strip()
    
    def _find_best_din_match(self, detected_text):
        # Mencari match terbaik dari DIN_TYPES berdasarkan similarity | Tujuan: Find best match in DIN_TYPES using fuzzy matching
        from difflib import SequenceMatcher
        
        detected_normalized = self._normalize_din_code(detected_text)  # Normalize detected text
        detected_clean = detected_normalized.replace(' ', '').upper()  # Remove space dan uppercase
        
        best_match = None  # Variable untuk menyimpan match terbaik
        best_score = 0.0  # Variable untuk menyimpan score terbaik
        
        # Loop semua DIN types (skip index 0 yang berisi "Select Label . . .")
        for din_type in DIN_TYPES[1:]:
            target_clean = din_type.replace(' ', '').upper()  # Clean target untuk comparison
            
            # Calculate similarity ratio menggunakan SequenceMatcher
            ratio = SequenceMatcher(None, detected_clean, target_clean).ratio()
            
            # Jika similarity > 80% dan lebih baik dari best_score, update best_match
            if ratio > 0.8 and ratio > best_score:
                best_score = ratio
                best_match = din_type
        
        return best_match, best_score  # Return match terbaik dan scorenya
    
    def _is_valid_din_format(self, text):
        # Validasi apakah text memiliki format DIN yang valid | Tujuan: Validate DIN format
        text_clean = text.replace(' ', '').upper()  # Remove space dan uppercase
        
        # Pattern untuk berbagai format DIN
        patterns = [
            r'^LBN\d$',                    # LBN1, LBN2, LBN3
            r'^LN\d\d{3}A$',               # LN1450A, LN0260A, LN1295A
            r'^LN\d\d{3}A$',               # LN2360A, LN2345A
            r'^LN\d\d{3}A$',               # LN3490A
            r'^LN\d\d{3}A$',               # LN4650A
            r'^LN\d\d{3}AISS$',            # LN4776AISS
        ]
        
        # Cek apakah text match dengan salah satu pattern
        for pattern in patterns:
            if re.match(pattern, text_clean):
                return True  # Valid format
        
        return False  # Invalid format
    
    def scan_frame(self, frame, is_static=False, original_frame=None):
        # Fungsi OCR terpisah yang menggunakan regex dan koreksi | Tujuan: Perform OCR pada frame dan detect kode yang valid
        best_match = None  # Variable untuk menyimpan match terbaik
        frame_to_save = original_frame if original_frame is not None else frame  # Frame untuk disimpan ke file
        
        # Handling untuk live camera - cek scan lock
        if not is_static:
            if not self.scan_lock.acquire(blocking=False):  # Try acquire lock tanpa blocking
                return  # Jika lock sudah digunakan, skip scan ini
            
            # Crop frame ke square seperti di _process_and_send_frame
            h_orig, w_orig, _ = frame.shape
            min_dim_orig = min(h_orig, w_orig)
            start_x_orig = (w_orig - min_dim_orig) // 2
            start_y_orig = (h_orig - min_dim_orig) // 2
            frame = frame[start_y_orig:start_y_orig + min_dim_orig, start_x_orig:start_x_orig + min_dim_orig]
            
            # Apply binary mode jika aktif
            if self.binary_mode:
                gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                _, frame = cv2.threshold(gray_frame, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                frame = cv2.cvtColor(frame, cv2.COLOR_GRAY2BGR)
            
            # Apply split mode jika aktif
            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2
                frame_scaled_320 = cv2.resize(frame, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)

                # Bagian atas - binary version
                gray_top = cv2.cvtColor(frame_scaled_320.copy(), cv2.COLOR_BGR2GRAY)
                _, frame_top_binary = cv2.threshold(gray_top, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                frame_top_binary = cv2.cvtColor(frame_top_binary, cv2.COLOR_GRAY2BGR)
                frame_bottom_original = frame_scaled_320.copy()
                
                # Buat canvas untuk split
                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3))
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3))
                
                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2
                
                # Paste frame ke canvas
                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_binary
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original
                            
                # Stack vertikal
                frame_combined = np.vstack([canvas_top, canvas_bottom])
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)

            else:  # Mode normal
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                from config import Resampling
                img = img.resize((self.TARGET_WIDTH, self.TARGET_HEIGHT), Resampling)

        try:
            h, w = frame.shape[:2]  # Get dimensi frame
            
            # Resize frame jika terlalu besar untuk performa OCR
            if w > 640:
                scale = 640 / w  # Hitung scale factor
                new_w, new_h = 640, int(h * scale)  # Hitung dimensi baru
                frame_small = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
            else:
                frame_small = frame
                
            gray = cv2.cvtColor(frame_small, cv2.COLOR_BGR2GRAY)  # Convert ke grayscale untuk preprocessing
            
            processing_stages = {}  # Dictionary untuk menyimpan berbagai preprocessing result
            
            # Preprocessing untuk DIN - lebih agresif
            if self.preset == "DIN":
                # 1. Contrast enhancement menggunakan CLAHE
                clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
                enhanced = clahe.apply(gray)
                processing_stages['Enhanced'] = enhanced
                
                # 2. Multiple threshold methods untuk cover berbagai kondisi pencahayaan
                _, binary1 = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                processing_stages['Binary_Otsu'] = binary1
                
                adaptive1 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Gaussian'] = adaptive1
                
                adaptive2 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Mean'] = adaptive2
                
                # 3. Inverted versions untuk text putih di background hitam
                processing_stages['Binary_Inv'] = cv2.bitwise_not(binary1)
                processing_stages['Adaptive_Inv'] = cv2.bitwise_not(adaptive1)
                
                # 4. Morphological operations untuk cleanup noise
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
                morph = cv2.morphologyEx(binary1, cv2.MORPH_CLOSE, kernel)
                processing_stages['Morphed'] = morph
                
            else:  # JIS - preprocessing lebih simple
                # Sharpen image untuk edge enhancement
                kernel = np.array([[-1,-1,-1], [-1, 9,-1],[-1,-1,-1]])
                processing_stages['Sharpened'] = cv2.filter2D(gray, -1, kernel)
                processing_stages['Grayscale'] = gray  # Original grayscale
                processing_stages['Inverted_Gray'] = cv2.bitwise_not(gray)  # Inverted grayscale
                # Adaptive threshold dengan inverted untuk JIS
                processed_frame_binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                processing_stages['Binary'] = processed_frame_binary

            all_results = []  # List untuk menyimpan semua hasil OCR
            
            # Set allowlist characters berdasarkan preset
            if self.preset == "JIS":
                allowlist_chars = ALLOWLIST_JIS
            else:
                allowlist_chars = ALLOWLIST_DIN

            # OCR dengan berbagai preprocessing - jalankan OCR untuk setiap preprocessing stage
            for stage_name, processed_frame in processing_stages.items():
                try:
                    results = self.reader.readtext(
                        processed_frame, 
                        detail=0,  # Return text only, tanpa bounding box
                        paragraph=False if self.preset == "JIS" else True,  # Paragraph mode untuk DIN (multi-line)
                        min_size=10 if self.preset == "JIS" else 15,  # Minimum size text untuk detect
                        width_ths=0.7 if self.preset == "JIS" else 0.5,  # Width threshold untuk grouping
                        allowlist=allowlist_chars  # Character whitelist
                    )
                    all_results.extend(results)  # Tambahkan hasil ke list
                except Exception as e:
                    print(f"OCR error on {stage_name}: {e}")
                    continue  # Skip stage ini jika error

            # Emit semua text hasil OCR untuk debugging/display
            if self.all_text_signal:
                unique_results = list(set(all_results))  # Remove duplicate
                self.all_text_signal.emit(unique_results)

            current_preset = self.preset  # Save current preset
            
            # Processing hasil OCR berdasarkan preset
            if current_preset == "DIN":
                # Untuk DIN, gunakan pendekatan matching dengan DIN_TYPES
                best_match_text = None
                best_match_score = 0.0
                
                # Loop semua hasil OCR
                for text in all_results:
                    # Apply OCR correction untuk fix common errors
                    text_fixed = fix_common_ocr_errors(text, self.preset)
                    
                    # Skip jika terlalu pendek (bukan kode valid)
                    if len(text_fixed.replace(' ', '')) < 4:
                        continue
                    
                    # Cari match terbaik dari DIN_TYPES menggunakan fuzzy matching
                    matched_type, score = self._find_best_din_match(text_fixed)
                    
                    # Update best match jika score lebih baik
                    if matched_type and score > best_match_score:
                        best_match_score = score
                        best_match_text = matched_type
                
                # Set best_match jika score cukup tinggi (>80%)
                if best_match_text and best_match_score > 0.8:
                    best_match = best_match_text
                    
            else:  # JIS - gunakan regex pattern matching
                pattern = self.patterns.get(current_preset, r".*")  # Get pattern untuk preset
                best_match_length = -1  # Track panjang match terbaik
                
                # Loop semua hasil OCR
                for text in all_results:
                    text_fixed = fix_common_ocr_errors(text, self.preset)  # Apply OCR correction
                    
                    detected_type = self._detect_code_type(text_fixed)  # Detect tipe kode (JIS/DIN)
                    is_valid, message = self._validate_preset_match(text_fixed, detected_type)  # Validate match dengan preset
                    
                    # Skip jika tidak valid
                    if not is_valid:
                        continue
                    
                    match = re.search(pattern, text_fixed, re.IGNORECASE)  # Try match dengan regex pattern
                    
                    # Jika match dan lebih panjang dari match sebelumnya
                    if match:
                        detected = text_fixed.strip()
                        if len(detected) > best_match_length:
                            best_match = detected  # Update best match
                            best_match_length = len(detected)
                        
            # Jika ada match yang ditemukan, process dan save
            if best_match:
                detected_code = best_match.strip()
                
                # Final formatting berdasarkan preset
                if self.preset == "DIN":
                    detected_code = self._normalize_din_code(detected_code)  # Normalize DIN format
                else:
                    detected_code = detected_code.replace(' ', '')  # Remove space untuk JIS
                
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Generate timestamp untuk record
                
                # Validate tipe kode dengan preset
                detected_type = self._detect_code_type(detected_code)
                is_valid_type, error_message = self._validate_preset_match(detected_code, detected_type)
                
                # Jika tipe tidak valid, emit error dan return
                if not is_valid_type:
                    self.code_detected_signal.emit(error_message)
                    if not is_static:
                        self.scan_lock.release()  # Release lock sebelum return
                    return
                
                # Compare dengan target label untuk status OK/Not OK
                if self.preset == "DIN":
                    target_normalized = self._normalize_din_code(self.target_label)  # Normalize target
                    detected_normalized = self._normalize_din_code(detected_code)  # Normalize detected
                    status = "OK" if detected_normalized.upper() == target_normalized.upper() else "Not OK"
                else:
                    status = "OK" if detected_code == self.target_label else "Not OK"
                
                # Set target session - gunakan target_label jika ada, kalau tidak gunakan detected_code
                target_session = self.target_label if self.target_label else detected_code

                # Untuk live camera, cek duplicate detection dalam 5 detik terakhir
                if not is_static:
                    if any(rec["Code"] == detected_code and 
                           (datetime.now() - datetime.strptime(rec["Time"], "%Y-%m-%d %H:%M:%S")).total_seconds() < 5
                           for rec in self.detected_codes):
                        return  # Skip jika duplicate
                
                # Generate filename untuk image
                img_filename = f"karton_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                img_path = os.path.join(IMAGE_DIR, img_filename)  # Full path untuk save image
                
                # Convert frame ke binary dan save
                frame_binary = convert_frame_to_binary(frame_to_save)
                cv2.imwrite(img_path, frame_binary)

                # Insert ke database
                new_id = insert_detection(timestamp, detected_code, current_preset, img_path, status, target_session)

                # Jika insert berhasil, tambahkan ke detected_codes list
                if new_id:
                    record = {
                        "ID": new_id,
                        "Time": timestamp,
                        "Code": detected_code,
                        "Type": current_preset,
                        "ImagePath": img_path,
                        "Status": status,
                        "TargetSession": target_session
                    }
                    self.detected_codes.append(record)  # Tambahkan ke list
                
                self.code_detected_signal.emit(detected_code)  # Emit signal kode terdeteksi
                
            else:  # Jika tidak ada match
                if is_static:  # Untuk static file, emit FAILED
                    self.code_detected_signal.emit("FAILED")

        except Exception as e:
            print(f"OCR/Regex error: {e}")  # Print error untuk debugging
            if is_static:
                self.code_detected_signal.emit(f"ERROR: {e}")  # Emit error untuk static file
                
        finally:
            # Cleanup - release lock untuk live camera
            if not is_static:
                self.scan_lock.release()
    
    def start_detection(self):
        # Mulai deteksi kamera | Tujuan: Start camera detection
        if self.running:  # Jika sudah running, skip
            return
        self.running = True  # Set flag running
        self.start()  # Start thread (akan call run())

    def stop_detection(self):
        # Hentikan deteksi kamera | Tujuan: Stop camera detection
        self.running = False  # Set flag running ke False untuk stop loop
        if self.cap:  # Jika kamera masih terbuka
             self.cap.release()  # Release kamera
             
    def set_camera_options(self, preset, flip_h, flip_v, binary_mode, split_mode, scan_interval):
        # Set opsi kamera | Tujuan: Set camera options
        self.preset = preset  # Set preset (JIS/DIN)
        self.flip_h = flip_h  # Set horizontal flip (tidak digunakan di kode saat ini)
        self.flip_v = flip_v  # Set vertical flip (tidak digunakan di kode saat ini)
        self.binary_mode = binary_mode  # Set binary mode flag
        self.split_mode = split_mode  # Set split mode flag
        self.scan_interval = scan_interval  # Set scan interval
    
    def set_target_label(self, label):
        # Set target label yang akan divalidasi | Tujuan: Set target label for validation
        self.target_label = label  # Set target label untuk comparison OK/Not OK

    def check_daily_reset(self):
        # Dipanggil setiap detik oleh timer UI (Clock) | Tujuan: Check daily reset
        now = datetime.now()  # Get waktu sekarang
        new_date = now.date()  # Get tanggal sekarang
        if new_date > self.current_date:  # Jika tanggal berubah (lewat tengah malam)
            self.current_date = new_date  # Update current date
            self.detected_codes = []  # Clear detected codes list
            self.detected_codes = load_existing_data(self.current_date)  # Load data untuk hari baru
            self.data_reset_signal.emit()  # Emit signal reset data
            return True  # Return True jika terjadi reset
        return False  # Return False jika tidak ada reset
        
    def scan_file(self, filepath):
        # Memproses file statis di Logic, dipanggil dari thread terpisah di UI | Tujuan: Process static file in Logic
        if self.running: return "STOP_LIVE"  # Jika live camera aktif, return error
        
        try:
            frame = cv2.imread(filepath)  # Baca image file
            
            # Validasi frame berhasil di-load
            if frame is None or frame.size == 0:
                return "LOAD_ERROR"  # Return error jika gagal load
            
            # Process dan send frame ke UI untuk display
            self._process_and_send_frame(frame, is_static=True)
            
            # Jalankan scan di thread terpisah agar tidak blocking
            threading.Thread(target=self.scan_frame,
                            args=(frame.copy(),),  # Copy frame untuk avoid race condition
                            kwargs={'is_static': True, 'original_frame': frame.copy()},
                            daemon=True).start()
            
            return "SCANNING"  # Return status scanning
            
        except Exception as e:
            print(f"File scan error: {e}")  # Print error untuk debugging
            return f"PROCESS_ERROR: {e}"  # Return error message

    def _detect_code_type(self, code):
        # Mendeteksi apakah kode yang terdeteksi adalah JIS atau DIN | Tujuan: Detect code type as JIS or DIN
        code_normalized = code.replace(' ', '').upper()  # Remove space dan uppercase
        
        # Pattern JIS: format angka-huruf-angka tanpa space
        jis_pattern = r"\b\d{2,3}[A-H]\d{2,3}[LR]?(?:\(S\))?\b"
        if re.match(jis_pattern, code_normalized):
            return "JIS"  # Return JIS jika match pattern
        
        # Pattern DIN: cek apakah ada di DIN_TYPES
        for din_type in DIN_TYPES[1:]:  # Skip index 0 yang berisi "Select Label . . ."
            if code_normalized == din_type.replace(' ', '').upper():
                return "DIN"  # Return DIN jika match dengan salah satu type
        
        # Pattern DIN general - untuk menangkap format DIN yang belum terdaftar
        din_patterns = [
            r'^LBN\d$',
            # Pattern LBN + 1 digit
            r'^LN\d\d{2,3}A(?:ISS)?$',
            # Pattern LN + digit + 2-3 angka + A + optional ISS
        ]
        
        # Cek apakah match dengan salah satu pattern DIN general
        for pattern in din_patterns:
            if re.match(pattern, code_normalized):
                return "DIN"  # Return DIN jika match
        
        return None  # Return None jika tidak match dengan JIS maupun DIN

    def _validate_preset_match(self, detected_code, detected_type):
        # Memvalidasi apakah tipe kode yang terdeteksi sesuai dengan preset yang dipilih | Tujuan: Validate preset match
        if detected_type is None:  # Jika tipe tidak terdeteksi
            return False, "Format kode tidak valid"  # Return False dengan error message
        
        if detected_type != self.preset:  # Jika tipe tidak sesuai dengan preset
            if self.preset == "JIS":  # Jika preset JIS tapi detected DIN
                return False, "Pastikan foto anda adalah Type JIS"
            else:  # Jika preset DIN tapi detected JIS
                return False, "Pastikan foto anda adalah Type DIN"
        
        return True, ""  # Return True jika valid

    def delete_codes(self, record_ids):
        # Wrapper untuk delete_codes dari database | Tujuan: Delete codes from database
        from database import delete_codes
        
        # Call fungsi delete dari database module
        if delete_codes(record_ids):
            # Jika berhasil, remove dari detected_codes list
            self.detected_codes = [rec for rec in self.detected_codes if rec['ID'] not in record_ids]
            return True  # Return True jika berhasil
        return False  # Return False jika gagal%