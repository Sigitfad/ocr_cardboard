# Logika deteksi OCR menggunakan EasyOCR dengan threading untuk live camera dan file scanning
# File ini berisi DetectionLogic class yang menangani OCR, frame processing, dan detection logic
# MODIFIED: 2-Stage Detection dengan STRUCTURAL CORRECTION untuk huruf tengah JIS dan (S) detection
# UPDATED: Binary mode diganti dengan Edge Detection mode

import cv2
import easyocr
import re
import os
import time
import threading
import atexit
import numpy as np
from datetime import datetime
from difflib import SequenceMatcher
from config import (
    IMAGE_DIR, EXCEL_DIR, DB_FILE, PATTERNS, ALLOWLIST_JIS, ALLOWLIST_DIN, DIN_TYPES,
    CAMERA_WIDTH, CAMERA_HEIGHT, TARGET_WIDTH, TARGET_HEIGHT, BUFFER_SIZE,
    MAX_CAMERAS, SCAN_INTERVAL, JIS_TYPES
)
from utils import (
    fix_common_ocr_errors, convert_frame_to_binary, find_external_camera,
    create_directories, apply_edge_detection
)
from database import (
    setup_database, load_existing_data, insert_detection
)
from PIL import Image


class DetectionLogic(threading.Thread):
    
    def __init__(self, update_signal, code_detected_signal, camera_status_signal, data_reset_signal, all_text_signal=None):
        super().__init__()
        
        self.update_signal = update_signal
        self.code_detected_signal = code_detected_signal
        self.camera_status_signal = camera_status_signal
        self.data_reset_signal = data_reset_signal
        self.all_text_signal = all_text_signal
        
        self.running = False
        self.cap = None
        self.preset = "JIS"
        self.last_scan_time = 0
        self.scan_interval = SCAN_INTERVAL
        
        self.target_label = ""
        
        create_directories()
        
        self.current_camera_index = 0
        self.scan_lock = threading.Lock()
        self.temp_files_on_exit = []
        
        self.edge_mode = False  # CHANGED: dari binary_mode ke edge_mode
        self.split_mode = False
        
        self.current_date = datetime.now().date()
        
        self.TARGET_WIDTH = TARGET_WIDTH
        self.TARGET_HEIGHT = TARGET_HEIGHT
        
        self.patterns = PATTERNS
        
        setup_database()
        self.detected_codes = load_existing_data(self.current_date)
        
        self.reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        
        atexit.register(self.cleanup_temp_files)
    
    def cleanup_temp_files(self):
        for t_path in self.temp_files_on_exit:
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    pass
    
    def run(self):
        self.current_camera_index = find_external_camera(MAX_CAMERAS)
        
        self.cap = cv2.VideoCapture(self.current_camera_index + cv2.CAP_DSHOW)
        
        if not self.cap.isOpened():
            self.cap = cv2.VideoCapture(self.current_camera_index)

        if not self.cap.isOpened():
            self.camera_status_signal.emit(f"Error: Kamera Index {self.current_camera_index} Gagal Dibuka.", False)
            self.running = False
            return
        
        try:
            self.cap.set(cv2.CAP_PROP_BUFFERSIZE, BUFFER_SIZE)
        except:
            pass
            
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
        self.cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))
        
        camera_info = f"Camera: {self.current_camera_index} ({'External' if self.current_camera_index != 0 else 'Internal'})"
        self.camera_status_signal.emit(camera_info, True)

        while self.running:
            ret, frame = self.cap.read()
            if not ret:
                break

            self._process_and_send_frame(frame, is_static=False)

            current_time = time.time()
            if current_time - self.last_scan_time >= self.scan_interval and not self.scan_lock.locked():
                self.last_scan_time = current_time
                threading.Thread(target=self.scan_frame, 
                                args=(frame.copy(),),
                                kwargs={'is_static': False, 'original_frame': frame.copy()}, 
                                daemon=True).start()
        
        if self.cap:
             self.cap.release()
        self.camera_status_signal.emit("Camera Off", False)
    
    def _process_and_send_frame(self, frame, is_static):
        from PIL import Image

        frame_display = frame.copy()

        if not is_static:
            h, w, _ = frame_display.shape
            min_dim = min(h, w)
            start_x = (w - min_dim) // 2
            start_y = (h - min_dim) // 2
            frame_cropped = frame_display[start_y:start_y + min_dim, start_x:start_x + min_dim]

            # UPDATED: Edge Detection mode
            if self.edge_mode:
                frame_cropped = apply_edge_detection(frame_cropped)

            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2
                frame_scaled_320 = cv2.resize(frame_cropped, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)

                # UPDATED: Apply edge detection untuk top frame
                frame_top_edge = apply_edge_detection(frame_scaled_320.copy())
                frame_bottom_original = frame_scaled_320.copy()

                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)

                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2

                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_edge
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original

                frame_combined = np.vstack([canvas_top, canvas_bottom])
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)

            else:
                frame_rgb = cv2.cvtColor(frame_cropped, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                from config import Resampling
                img = img.resize((self.TARGET_WIDTH, self.TARGET_HEIGHT), Resampling)

        else:
            # UPDATED: Edge detection untuk static file
            if self.edge_mode or self.split_mode:
                frame_display = apply_edge_detection(frame_display)
                cv2.putText(frame_display, "Edge Detection Mode", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1, cv2.LINE_AA)

            frame_rgb = cv2.cvtColor(frame_display, cv2.COLOR_BGR2RGB)
            original_img = Image.fromarray(frame_rgb)

            original_width, original_height = original_img.size
            ratio = min(self.TARGET_WIDTH / original_width, self.TARGET_HEIGHT / original_height)

            new_width = int(original_width * ratio)
            new_height = int(original_height * ratio)

            from config import Resampling
            img_resized = original_img.resize((new_width, new_height), Resampling)

            img = Image.new('RGB', (self.TARGET_WIDTH, self.TARGET_HEIGHT), 'black')
            x_offset = (self.TARGET_WIDTH - new_width) // 2
            y_offset = (self.TARGET_HEIGHT - new_height) // 2
            img.paste(img_resized, (x_offset, y_offset))

            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("arial.ttf", 10)
            except IOError:
                font = ImageFont.load_default()

            text_to_display = "STATIC FILE SCAN"
            bbox = draw.textbbox((0, 0), text_to_display, font=font)
            text_width = bbox[2] - bbox[0]
            x_center = (self.TARGET_WIDTH - text_width) // 2
            y_top = 12
            draw.text((x_center, y_top), text_to_display, fill=(255, 255, 0), font=font)

        self.update_signal.emit(img)
    
    def _normalize_din_code(self, code):
        code = code.strip().upper()
        code_no_space = re.sub(r'\s+', '', code)
        
        match = re.match(r'^(LBN)(\d)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}"
        
        match = re.match(r'^(LN\d)(\d+)([A-Z])$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)}"
        
        match = re.match(r'^(LN\d)(\d+)([A-Z])(ISS)$', code_no_space)
        if match:
            return f"{match.group(1)} {match.group(2)}{match.group(3)} {match.group(4)}"
        
        return re.sub(r'\s+', ' ', code).strip()
    
    def _find_best_din_match(self, detected_text):
        detected_normalized = self._normalize_din_code(detected_text)
        detected_clean = detected_normalized.replace(' ', '').upper()
        
        best_match = None
        best_score = 0.0
        
        for din_type in DIN_TYPES[1:]:
            target_clean = din_type.replace(' ', '').upper()
            
            ratio = SequenceMatcher(None, detected_clean, target_clean).ratio()
            
            if ratio > 0.8 and ratio > best_score:
                best_score = ratio
                best_match = din_type
        
        return best_match, best_score
    
    def _correct_jis_structure(self, text):
        """
        KOREKSI STRUKTURAL JIS: Perbaiki huruf tengah yang salah terbaca sebagai angka
        Format JIS: [2-3 digit][1 HURUF A-H][2-3 digit][L/R optional][(S) optional]
        """
        text = text.strip().upper().replace(' ', '')
        
        digit_to_letter = {
            '0': 'D', '1': 'I', '2': 'Z', '3': 'B', 
            '4': 'A', '5': 'S', '6': 'G', '8': 'B',
        }
        
        text = re.sub(r'\(5\)', r'(S)', text)
        text = re.sub(r'5\)', r'(S)', text)
        text = re.sub(r'\([S5](?!\))', r'(S)', text)
        
        pattern = r'^(\d{2,3})([A-Z0-9])(\d{2,3})([LR])?(\(S\))?$'
        
        match = re.match(pattern, text)
        
        if match:
            capacity = match.group(1)
            middle_char = match.group(2)
            size = match.group(3)
            terminal = match.group(4) or ''
            option = match.group(5) or ''
            
            if middle_char.isdigit():
                corrected_letter = digit_to_letter.get(middle_char, 'D')
                
                if corrected_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
                    middle_char = corrected_letter
            
            corrected = f"{capacity}{middle_char}{size}{terminal}{option}"
            return corrected
        
        return text
    
    def _find_best_jis_match(self, detected_text):
        """
        TAHAP 2: Mencari match terbaik dari JIS_TYPES dengan (S) preserved
        """
        detected_corrected = self._correct_jis_structure(detected_text)
        
        detected_clean = detected_corrected.replace(' ', '').upper()
        
        best_match = None
        best_score = 0.0
        
        for jis_type in JIS_TYPES[1:]:
            target_clean = jis_type.replace(' ', '').upper()
            
            ratio = SequenceMatcher(None, detected_clean, target_clean).ratio()
            
            if ratio > 0.85 and ratio > best_score:
                best_score = ratio
                best_match = jis_type
        
        if not best_match or best_score < 0.90:
            detected_without_s = detected_clean.replace('(S)', '')
            
            for jis_type in JIS_TYPES[1:]:
                target_without_s = jis_type.replace(' ', '').replace('(S)', '').upper()
                
                ratio = SequenceMatcher(None, detected_without_s, target_without_s).ratio()
                
                if ratio > 0.90:
                    if '(S)' in detected_clean:
                        base_code = jis_type.replace('(S)', '')
                        candidate_with_s = base_code + '(S)'
                        
                        if candidate_with_s in JIS_TYPES:
                            best_match = candidate_with_s
                            best_score = ratio
                            break
                    else:
                        if '(S)' not in jis_type and ratio > best_score:
                            best_match = jis_type
                            best_score = ratio
        
        return best_match, best_score
    
    def scan_frame(self, frame, is_static=False, original_frame=None):
        """
        TAHAP 1: OCR mentah
        TAHAP 2: Structural correction + Fuzzy matching
        """
        best_match = None
        frame_to_save = original_frame if original_frame is not None else frame
        
        if not is_static:
            if not self.scan_lock.acquire(blocking=False):
                return
            
            h_orig, w_orig, _ = frame.shape
            min_dim_orig = min(h_orig, w_orig)
            start_x_orig = (w_orig - min_dim_orig) // 2
            start_y_orig = (h_orig - min_dim_orig) // 2
            frame = frame[start_y_orig:start_y_orig + min_dim_orig, start_x_orig:start_x_orig + min_dim_orig]
            
            # UPDATED: Edge detection mode
            if self.edge_mode:
                frame = apply_edge_detection(frame)
            
            if self.split_mode:
                TARGET_CONTENT_SIZE = self.TARGET_HEIGHT // 2
                frame_scaled_320 = cv2.resize(frame, (TARGET_CONTENT_SIZE, TARGET_CONTENT_SIZE), interpolation=cv2.INTER_AREA)

                # UPDATED: Apply edge detection untuk top frame
                frame_top_edge = apply_edge_detection(frame_scaled_320.copy())
                frame_bottom_original = frame_scaled_320.copy()
                
                canvas_top = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                canvas_bottom = np.zeros((TARGET_CONTENT_SIZE, self.TARGET_WIDTH, 3), dtype=np.uint8)
                
                x_offset = (self.TARGET_WIDTH - TARGET_CONTENT_SIZE) // 2
                
                canvas_top[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_top_edge
                canvas_bottom[:, x_offset:x_offset + TARGET_CONTENT_SIZE] = frame_bottom_original
                            
                frame_combined = np.vstack([canvas_top, canvas_bottom])
                frame_rgb = cv2.cvtColor(frame_combined, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)

        try:
            h, w = frame.shape[:2]
            
            if w > 640:
                scale = 640 / w
                new_w, new_h = 640, int(h * scale)
                frame_small = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
            else:
                frame_small = frame
                
            gray = cv2.cvtColor(frame_small, cv2.COLOR_BGR2GRAY)
            
            processing_stages = {}
            
            if self.preset == "DIN":
                clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
                enhanced = clahe.apply(gray)
                processing_stages['Enhanced'] = enhanced
                
                _, binary1 = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                processing_stages['Binary_Otsu'] = binary1
                
                adaptive1 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Gaussian'] = adaptive1
                
                adaptive2 = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2)
                processing_stages['Adaptive_Mean'] = adaptive2
                
                processing_stages['Binary_Inv'] = cv2.bitwise_not(binary1)
                processing_stages['Adaptive_Inv'] = cv2.bitwise_not(adaptive1)
                
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
                morph = cv2.morphologyEx(binary1, cv2.MORPH_CLOSE, kernel)
                processing_stages['Morphed'] = morph
                
            else:
                kernel = np.array([[-1,-1,-1], [-1, 9,-1],[-1,-1,-1]])
                processing_stages['Sharpened'] = cv2.filter2D(gray, -1, kernel)
                processing_stages['Grayscale'] = gray
                processing_stages['Inverted_Gray'] = cv2.bitwise_not(gray)
                processed_frame_binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                processing_stages['Binary'] = processed_frame_binary

            all_results = []
            
            if self.preset == "JIS":
                allowlist_chars = ALLOWLIST_JIS
            else:
                allowlist_chars = ALLOWLIST_DIN

            for stage_name, processed_frame in processing_stages.items():
                try:
                    results = self.reader.readtext(
                        processed_frame, 
                        detail=0,
                        paragraph=False if self.preset == "JIS" else True,
                        min_size=10 if self.preset == "JIS" else 15,
                        width_ths=0.7 if self.preset == "JIS" else 0.5,
                        allowlist=allowlist_chars
                    )
                    all_results.extend(results)
                except Exception as e:
                    print(f"OCR error on {stage_name}: {e}")
                    continue

            if self.all_text_signal:
                unique_results = list(set(all_results))
                self.all_text_signal.emit(unique_results)

            current_preset = self.preset
            
            if current_preset == "DIN":
                best_match_text = None
                best_match_score = 0.0
                
                for text in all_results:
                    text_fixed = fix_common_ocr_errors(text, self.preset)
                    
                    if len(text_fixed.replace(' ', '')) < 4:
                        continue
                    
                    matched_type, score = self._find_best_din_match(text_fixed)
                    
                    if matched_type and score > best_match_score:
                        best_match_score = score
                        best_match_text = matched_type
                
                if best_match_text and best_match_score > 0.8:
                    best_match = best_match_text
                    
            else:
                best_match_text = None
                best_match_score = 0.0
                
                for text in all_results:
                    if len(text.replace(' ', '').replace('(S)', '')) < 5:
                        continue
                    
                    matched_type, score = self._find_best_jis_match(text)
                    
                    if matched_type and score > best_match_score:
                        best_match_score = score
                        best_match_text = matched_type
                
                if best_match_text and best_match_score > 0.85:
                    best_match = best_match_text
                        
            if best_match:
                detected_code = best_match.strip()
                
                if self.preset == "DIN":
                    detected_code = self._normalize_din_code(detected_code)
                else:
                    detected_code = detected_code.replace(' ', '')
                
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                detected_type = self._detect_code_type(detected_code)
                is_valid_type, error_message = self._validate_preset_match(detected_code, detected_type)
                
                if not is_valid_type:
                    self.code_detected_signal.emit(error_message)
                    if not is_static:
                        self.scan_lock.release()
                    return
                
                if self.preset == "DIN":
                    target_normalized = self._normalize_din_code(self.target_label)
                    detected_normalized = self._normalize_din_code(detected_code)
                    status = "OK" if detected_normalized.upper() == target_normalized.upper() else "Not OK"
                else:
                    status = "OK" if detected_code == self.target_label else "Not OK"
                
                target_session = self.target_label if self.target_label else detected_code

                if not is_static:
                    if any(rec["Code"] == detected_code and 
                           (datetime.now() - datetime.strptime(rec["Time"], "%Y-%m-%d %H:%M:%S")).total_seconds() < 5
                           for rec in self.detected_codes):
                        return
                
                img_filename = f"karton_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                img_path = os.path.join(IMAGE_DIR, img_filename)
                
                frame_binary = convert_frame_to_binary(frame_to_save)
                cv2.imwrite(img_path, frame_binary)

                new_id = insert_detection(timestamp, detected_code, current_preset, img_path, status, target_session)

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
                    self.detected_codes.append(record)
                
                self.code_detected_signal.emit(detected_code)
                
            else:
                if is_static:
                    self.code_detected_signal.emit("FAILED")

        except Exception as e:
            print(f"OCR/Regex error: {e}")
            if is_static:
                self.code_detected_signal.emit(f"ERROR: {e}")
                
        finally:
            if not is_static:
                self.scan_lock.release()
    
    def start_detection(self):
        if self.running:
            return
        self.running = True
        self.start()

    def stop_detection(self):
        self.running = False
        if self.cap:
             self.cap.release()
             
    def set_camera_options(self, preset, flip_h, flip_v, edge_mode, split_mode, scan_interval):
        self.preset = preset
        self.flip_h = flip_h
        self.flip_v = flip_v
        self.edge_mode = edge_mode  # CHANGED: dari binary_mode ke edge_mode
        self.split_mode = split_mode
        self.scan_interval = scan_interval
    
    def set_target_label(self, label):
        self.target_label = label

    def check_daily_reset(self):
        now = datetime.now()
        new_date = now.date()
        if new_date > self.current_date:
            self.current_date = new_date
            self.detected_codes = []
            self.detected_codes = load_existing_data(self.current_date)
            self.data_reset_signal.emit()
            return True
        return False
        
    def scan_file(self, filepath):
        if self.running: 
            return "STOP_LIVE"
        
        try:
            frame = cv2.imread(filepath)
            
            if frame is None or frame.size == 0:
                return "LOAD_ERROR"
            
            self._process_and_send_frame(frame, is_static=True)
            
            threading.Thread(target=self.scan_frame,
                            args=(frame.copy(),),
                            kwargs={'is_static': True, 'original_frame': frame.copy()},
                            daemon=True).start()
            
            return "SCANNING"
            
        except Exception as e:
            print(f"File scan error: {e}")
            return f"PROCESS_ERROR: {e}"

    def _detect_code_type(self, code):
        code_normalized = code.replace(' ', '').upper()
        
        jis_pattern = r"\b\d{2,3}[A-H]\d{2,3}[LR]?(?:\(S\))?\b"
        if re.match(jis_pattern, code_normalized):
            return "JIS"
        
        for din_type in DIN_TYPES[1:]:
            if code_normalized == din_type.replace(' ', '').upper():
                return "DIN"
        
        din_patterns = [
            r'^LBN\d$',
            r'^LN\d\d{2,3}A(?:ISS)?$',
        ]
        
        for pattern in din_patterns:
            if re.match(pattern, code_normalized):
                return "DIN"
        
        return None

    def _validate_preset_match(self, detected_code, detected_type):
        if detected_type is None:
            return False, "Format kode tidak valid"
        
        if detected_type != self.preset:
            if self.preset == "JIS":
                return False, "Pastikan foto anda adalah Type JIS"
            else:
                return False, "Pastikan foto anda adalah Type DIN"
        
        return True, ""

    def delete_codes(self, record_ids):
        from database import delete_codes
        
        if delete_codes(record_ids):
            self.detected_codes = [rec for rec in self.detected_codes if rec['ID'] not in record_ids]
            return True
        return False