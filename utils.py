import os # Digunakan untuk mengakses dan mengelola file/folder di sistem operasi
import re # Digunakan untuk pencocokan pola teks (regex), misalnya validasi kode atau format tertentu
import numpy as np # Digunakan untuk pengolahan array dan perhitungan numerik, terutama data gambar
import cv2 # Library OpenCV untuk pengolahan citra dan video (resize, threshold, edge detection, dll)
from datetime import datetime # Digunakan untuk mengambil dan mengelola tanggal serta waktu saat ini
from config import Resampling # Digunakan untuk metode resampling gambar (misalnya LANCZOS / ANTIALIAS)

# Suppress OpenCV error messages untuk camera index out of range
# Ini mencegah spam error message saat aplikasi scan untuk kamera
os.environ['OPENCV_LOG_LEVEL'] = 'ERROR'
os.environ['OPENCV_VIDEOIO_DEBUG'] = '0'
cv2.setLogLevel(0)  # 0 = Silent, 1 = Error, 2 = Warning, 3 = Info, 4 = Debug

# LANCZOS / ANTIALIAS adalah metode resampling gambar yang digunakan saat resize
# Fungsinya untuk menjaga kualitas dan ketajaman gambar, terutama teks, agar hasil OCR lebih akurat dan tidak pecah

def apply_edge_detection(frame):
    """
    Fungsi untuk menerapkan Edge Detection pada frame dengan garis putih neon terang
    Tujuan: Deteksi tepi objek dan teks dengan background HITAM PEKAT MURNI dan garis putih neon
    Parameter: frame = numpy array BGR frame dari OpenCV
    Return: numpy array frame edge detection dalam format BGR (3 channel)
    
    Menggunakan algoritma Canny Edge Detection dengan background pure black (0,0,0)
    """
    # Konversi frame BGR ke grayscale untuk memudahkan deteksi tepi
    # Grayscale diperlukan karena algoritma Canny bekerja pada single channel
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    
    # Apply Gaussian Blur untuk mengurangi noise
    # Kernel size (3,3) digunakan untuk menghaluskan gambar sebelum deteksi tepi
    # Sigma=0 berarti OpenCV akan menghitung nilai sigma otomatis
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    
    # Canny Edge Detection dengan threshold lebih rendah untuk lebih banyak detail
    # Threshold 30-100: nilai rendah untuk mendeteksi lebih banyak edge/tepi
    # Threshold rendah = lebih sensitif, menangkap detail lebih banyak
    edges = cv2.Canny(blurred, 30, 100)
    
    # Dilate edges untuk membuat garis lebih tebal dan terang
    # Kernel 2x2 digunakan untuk memperlebar/menebalkan garis edge
    # Iterations=1 berarti dilasi dilakukan 1 kali untuk membuat garis lebih terlihat
    kernel = np.ones((2, 2), np.uint8)
    edges_dilated = cv2.dilate(edges, kernel, iterations=1)
    
    # CRITICAL: Buat canvas HITAM MURNI (pure black) terlebih dahulu
    # Ini memastikan background benar-benar hitam pekat tanpa noise abu-abu
    # np.zeros membuat array dengan semua nilai 0 (hitam murni RGB 0,0,0)
    # Shape: tinggi x lebar x 3 channel (BGR) dengan tipe data uint8 (0-255)
    edges_bgr = np.zeros((edges_dilated.shape[0], edges_dilated.shape[1], 3), dtype=np.uint8)
    
    # HANYA gambar garis putih pada pixel yang terdeteksi sebagai edge
    # Background tetap hitam murni (0,0,0), hanya edge yang jadi putih (255,255,255)
    # Indexing boolean: pixel dengan nilai > 0 dari edges_dilated diisi putih [255,255,255]
    # Hasilnya: garis putih neon terang hanya di tepi objek, background tetap hitam
    edges_bgr[edges_dilated > 0] = [255, 255, 255]  # BGR: Putih murni hanya di edge
    
    # OPTIONAL: Tingkatkan brightness garis putih jika perlu lebih terang
    # Uncomment baris di bawah untuk garis lebih glowing
    # edges_bgr[edges_dilated > 0] = [255, 255, 255]
    
    return edges_bgr #Return frame edge detection dalam format BGR 3 channel


# === OCR CORRECTION LOGIC ===
def fix_common_ocr_errors_jis(text):
    # Bersihkan text: trim whitespace dan konversi ke uppercase
    # Strip menghapus spasi di awal/akhir, upper() untuk konsistensi huruf kapital
    text = text.strip().upper()
    
    # Hapus semua karakter selain huruf A-Z, angka 0-9, dan tanda kurung ()
    # Regex [^A-Z0-9()] = ambil semua KECUALI yang disebutkan, lalu replace dengan ''
    text = re.sub(r'[^A-Z0-9()]', '', text)

    # Dictionary mapping karakter yang sering salah dibaca OCR menjadi angka yang benar
    # Misalnya: huruf O sering terbaca sebagai angka 0, I sebagai 1, dst
    char_to_digit = {
        "O": "0", "Q": "0", "D": "0", "U": "0", "C": "0",  # Mirip angka 0
        "I": "1", "L": "1", "J": "1",  # Mirip angka 1
        "Z": "2", "E": "3", "A": "4", "H": "4",  # Mirip angka 2,3,4
        "S": "5", "G": "6", "T": "7", "Y": "7",  # Mirip angka 5,6,7
        "B": "8", "P": "9", "R": "9"  # Mirip angka 8,9
    }

    # Dictionary mapping angka yang sering salah dibaca menjadi huruf yang benar
    # Untuk koreksi posisi yang seharusnya huruf tapi terbaca angka
    digit_to_char = {
        "0": "D", "1": "L", "2": "Z", "3": "B", "4": "A", "5": "S", 
        "6": "G", "7": "T", "8": "B", "9": "R", "D" : "G"
    }

    # Regex pattern untuk mendeteksi struktur kode JIS battery
    # Pattern: (capacity)(type)(size)(terminal)(option)
    # Contoh: 55D23L atau 75D23R(S)
    match = re.search(r'(\d+|[A-Z]+)(\d+|[A-Z])(\d+|[A-Z]+)([L|R|1|0|4|D|I]?)(\(S\)|5\)|S)?$', text)

    # Jika pattern JIS terdeteksi, lakukan koreksi per-bagian
    if match:
        # Extract setiap komponen dari pattern yang terdeteksi
        capacity = match.group(1)  # Kapasitas baterai (biasanya angka)
        type_char = match.group(2)  # Tipe baterai (huruf)
        size = match.group(3)  # Ukuran baterai (angka)
        terminal = match.group(4)  # Terminal orientation (L/R)
        option = match.group(5)  # Option seperti (S)

        # Koreksi capacity: ubah huruf yang salah dibaca jadi angka
        # Join = gabungkan karakter hasil mapping kembali jadi string
        new_capacity = "".join([char_to_digit.get(c, c) for c in capacity])
        
        # Koreksi type character: jika terbaca angka, ubah ke huruf yang benar
        if type_char.isdigit():
            new_type = digit_to_char.get(type_char, type_char)
        else:
            new_type = type_char
        
        # Koreksi spesifik untuk tipe D, B, A yang sering salah dibaca
        if new_type in ['O', 'Q', 'G', '0', 'U', 'C']: new_type = 'D'
        if new_type in ['8', '3']: new_type = 'B'
        if new_type in ['4']: new_type = 'A'

        # Koreksi size: ubah huruf yang salah dibaca jadi angka
        new_size = "".join([char_to_digit.get(c, c) for c in size])

        # Koreksi terminal orientation (L atau R)
        # L = Left terminal, R = Right terminal
        if terminal:
            if terminal in ['1', 'I', 'J', '4']: 
                terminal = 'L'  # Karakter mirip L
            elif terminal in ['0', 'Q', 'D', 'O']: 
                terminal = 'R'  # Karakter mirip R

        # Koreksi option: pastikan format (S) yang benar
        if option:
            option = '(S)'

        # Gabungkan semua komponen yang sudah dikoreksi
        text_fixed = f"{new_capacity}{new_type}{new_size}{terminal}{option if option else ''}"
        return text_fixed.strip().upper()

    # Jika pattern tidak cocok, lakukan koreksi sederhana
    # Replace semua karakter sesuai mapping char_to_digit
    for char, digit in char_to_digit.items():
        text = text.replace(char, digit)
    
    # Koreksi khusus untuk option (S)
    text = text.replace('5)', '(S)').replace('(5)', '(S)')
    return text.strip().upper()


def fix_common_ocr_errors_din(text):
    # Bersihkan dan uppercase text untuk konsistensi
    text = text.strip().upper()
    
    # Dictionary untuk mapping karakter yang mirip angka
    # OCR sering salah membaca karakter mirip ini
    char_to_digit = {
        'O': '0', 'Q': '0',  # Huruf O/Q mirip angka 0
        'I': '1', 'l': '1',  # Huruf I/l mirip angka 1
        'Z': '2',  # Huruf Z mirip angka 2
        'S': '5',  # Huruf S mirip angka 5
        'G': '6',  # Huruf G mirip angka 6
        'B': '8',  # Huruf B mirip angka 8
    }
    
    # Dictionary untuk mapping angka yang mirip huruf
    # Untuk koreksi posisi yang seharusnya huruf
    digit_to_char = {
        '0': 'O',
        '1': 'I',
    }
    
    # Hapus karakter selain A-Z, 0-9, dan spasi
    # DIN format menggunakan spasi untuk memisahkan komponen
    text = re.sub(r'[^A-Z0-9\s]', '', text)
    
    # Split text berdasarkan spasi untuk memproses per-token
    # DIN format: "LBN 60 ISS" (3 token)
    tokens = text.split()
    
    # Return jika tidak ada token
    if len(tokens) == 0:
        return text
    
    corrected_tokens = []
    
    # Proses setiap token berdasarkan posisinya
    for i, token in enumerate(tokens):
        # Token pertama: kode tipe (contoh: LBN, LN2)
        if i == 0:
            corrected = ""
            for j, char in enumerate(token):
                # Karakter pertama: harus L (Left/Lead)
                if j == 0:
                    if char in ['I', '1', 'l']:
                        corrected += 'L'
                    else:
                        corrected += char
                # Karakter kedua: biasanya B atau N
                elif j == 1:
                    if char in ['8']:
                        corrected += 'B'  # Angka 8 mirip B
                    elif char in ['H', 'M']:
                        corrected += 'N'  # H/M mirip N
                    else:
                        corrected += char
                # Karakter ketiga: tergantung prefix
                elif j == 2:
                    if token[:2] == "LB":
                        if char in ['H', 'M']:
                            corrected += 'N'
                        else:
                            corrected += char
                    else:
                        # Untuk prefix lain, ubah huruf jadi angka
                        if char in char_to_digit:
                            corrected += char_to_digit[char]
                        else:
                            corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
        
        # Token kedua: kapasitas dalam angka (contoh: 60, 100)
        elif i == 1:
            corrected = ""
            for j, char in enumerate(token):
                if char.isdigit():
                    corrected += char
                elif char in char_to_digit:
                    # Ubah huruf mirip angka jadi angka
                    corrected += char_to_digit[char]
                elif j == len(token) - 1 and char.isalpha():
                    # Karakter terakhir bisa jadi suffix huruf (misal: A)
                    if char in ['4', 'H']:
                        corrected += 'A'
                    else:
                        corrected += char
                else:
                    corrected += char
            
            corrected_tokens.append(corrected)
        
        # Token ketiga: standard (ISS, EFB, AGM, dll)
        elif i == 2:
            corrected = token
            # Koreksi khusus untuk ISS yang sering salah dibaca
            if token in ['I55', 'IS5', 'I5S', '155']:
                corrected = 'ISS'
            elif token.replace('5', 'S').replace('I', 'I') == 'ISS':
                corrected = 'ISS'
            
            corrected_tokens.append(corrected)
    
    # Gabungkan kembali semua token dengan spasi
    result = ' '.join(corrected_tokens)
    
    # Tambahkan spasi jika tidak ada spasi antara komponen
    # Regex untuk memastikan format "LBN 60" atau "LN2 100"
    result = re.sub(r'(LBN)(\d)', r'\1 \2', result)
    result = re.sub(r'(LN\d)(\d)', r'\1 \2', result)
    result = re.sub(r'([A-Z0-9])(ISS)', r'\1 \2', result)
    
    # Bersihkan multiple spasi jadi single spasi
    result = re.sub(r'\s+', ' ', result)
    
    return result.strip()


def fix_common_ocr_errors(text, preset):
    # Fungsi dispatcher: pilih fungsi koreksi sesuai preset baterai
    # JIS = Japanese Industrial Standard (format: 55D23L)
    # DIN = Deutsche Industrie Norm (format: LBN 60 ISS)
    if preset == "JIS":
        return fix_common_ocr_errors_jis(text)
    elif preset == "DIN":
        return fix_common_ocr_errors_din(text)
    else:
        # Default ke JIS jika preset tidak dikenali
        return fix_common_ocr_errors_jis(text)


# === FRAME PROCESSING ===
def convert_frame_to_binary(frame):
    """
    UPDATED: Gunakan edge detection dengan putih neon untuk export Excel
    """
    # Konversi frame menjadi binary edge detection
    # Digunakan saat export ke Excel untuk gambar lebih jelas
    # Return: frame dengan background hitam dan garis putih neon
    return apply_edge_detection(frame)

def find_external_camera(max_cameras=5):
    # Fungsi untuk mencari kamera eksternal (bukan built-in laptop)
    # Scan hingga max_cameras perangkat untuk menemukan kamera terbaik
    # Priority: kamera eksternal (index > 0), fallback ke built-in (index 0)
    
    best_working_index = 0 #Default ke kamera pertama (biasanya built-in)

    # Loop cek semua kamera dari index 0 hingga max_cameras-1
    for i in range(max_cameras):
        cap = None
        try:
            # Coba buka kamera pada index i dengan backend default
            # Gunakan CAP_ANY untuk kompatibilitas maksimal
            cap = cv2.VideoCapture(i, cv2.CAP_ANY)
            
            # Timeout untuk mencegah hang - set max waktu tunggu
            cap.set(cv2.CAP_PROP_OPEN_TIMEOUT_MSEC, 1000)  # 1 detik timeout
            
            # Cek apakah kamera berhasil dibuka
            if cap.isOpened():
                # Coba baca 1 frame untuk validasi kamera benar-benar berfungsi
                ret, test_frame = cap.read()
                
                if ret and test_frame is not None:
                    # Ambil resolusi kamera (width x height)
                    w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
                    h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
                    
                    # Validasi kamera memiliki resolusi valid (w>0 dan h>0)
                    if w > 0 and h > 0:
                        # Jika index > 0, ini kamera eksternal - langsung return
                        if i > 0:
                            cap.release()
                            return i
                        else:
                            # Index 0 adalah built-in, simpan sebagai fallback
                            best_working_index = i
        
        except Exception as e:
            # Tangkap semua error OpenCV dan abaikan
            # Ini mencegah crash saat scan kamera yang tidak kompatibel
            pass
        
        finally:
            # Pastikan kamera selalu ditutup, bahkan jika terjadi error
            if cap is not None:
                try:
                    cap.release()
                except:
                    pass
    
    return best_working_index #Return kamera terbaik yang ditemukan (prioritas eksternal, fallback built-in)


def create_directories():
    # Fungsi untuk membuat folder-folder yang dibutuhkan aplikasi
    # Import config untuk ambil path folder IMAGE_DIR dan EXCEL_DIR
    from config import IMAGE_DIR, EXCEL_DIR
    
    # Buat folder untuk menyimpan gambar hasil capture
    # exist_ok=True berarti tidak error jika folder sudah ada
    os.makedirs(IMAGE_DIR, exist_ok=True)
    
    # Buat folder untuk menyimpan file Excel hasil export
    os.makedirs(EXCEL_DIR, exist_ok=True)


def cleanup_temp_files(temp_files_list):
    # Fungsi untuk membersihkan/menghapus file-file temporary
    # Parameter: list berisi path file yang akan dihapus
    
    # Loop setiap path file dalam list
    for t_path in temp_files_list:
        # Cek apakah file benar-benar ada
        if os.path.exists(t_path):
            try:
                # Coba hapus file
                os.remove(t_path)
            except:
                # Jika gagal hapus (file locked/permission), abaikan error
                # Pass = tidak melakukan apa-apa, lanjut ke file berikutnya
                pass