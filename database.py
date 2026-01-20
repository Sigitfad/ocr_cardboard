# Operasi database SQLite untuk menyimpan dan mengambil data deteksi
# File ini berisi semua fungsi database untuk CRUD operations dan migrasi schema
import sqlite3  #Import library SQLite untuk database operations
from datetime import datetime  #Modul untuk date/time handling
from config import DB_FILE  #Import path database file dari config.py


def setup_database():
    # Fungsi setup database dan buat table jika belum ada | Tujuan: Inisialisasi database dengan schema yang benar
    # Juga melakukan migration untuk menambah kolom baru jika diperlukan
    # Return: None (void function, hanya modifikasi database)

    conn = sqlite3.connect(DB_FILE)  # Koneksi ke database SQLite
    
    cursor = conn.cursor() #Buat cursor object untuk execute SQL commands
    
    # Check apakah table 'detected_codes' sudah ada di database
    # Query ke sqlite_master (system table yang menyimpan metadata schema)
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='detected_codes'")
    table_exists = cursor.fetchone() is not None  # fetchone() return None jika table tidak ada
    
    # Jika table belum ada, buat table baru dengan schema lengkap
    if not table_exists:
        # Buat table baru jika tidak ada dengan schema yang lengkap
        # Schema: id (auto increment PK), timestamp, code, preset, image_path, status, target_session
        cursor.execute('''CREATE TABLE detected_codes (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            timestamp TEXT,
                            code TEXT,
                            preset TEXT,
                            image_path TEXT,
                            status TEXT,
                            target_session TEXT
                          )''')
    else:
        # Jika table sudah ada, check dan tambah kolom yang mungkin hilang
        # Ini untuk backward compatibility dengan database lama yang mungkin belum punya kolom baru
        
        # Get informasi kolom yang ada di table
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]  # column[1] adalah nama kolom
        
        # Tambah kolom 'status' jika belum ada (untuk tracking OK/Not OK)
        # Kolom 'status' digunakan untuk validasi apakah deteksi sesuai dengan target_session
        if 'status' not in columns:
            try:
                # ALTER TABLE untuk tambah kolom baru dengan default value 'OK'
                cursor.execute("ALTER TABLE detected_codes ADD COLUMN status TEXT DEFAULT 'OK'")
                
                # Update semua existing records yang NULL menjadi 'OK'
                # Ini untuk ensure semua data lama punya status default
                cursor.execute("UPDATE detected_codes SET status = 'OK' WHERE status IS NULL")
            except Exception as e:
                # Silent fail jika ada error (mungkin kolom sudah ada di DB lain)
                pass
        
        # Tambah kolom 'target_session' jika belum ada (untuk tracking session/label target)
        # Kolom 'target_session' menyimpan label/sesi yang sedang aktif saat deteksi
        if 'target_session' not in columns:
            try:
                # ALTER TABLE untuk tambah kolom target_session
                cursor.execute("ALTER TABLE detected_codes ADD COLUMN target_session TEXT")
                
                # Update existing records: gunakan 'code' sebagai default target_session
                # Asumsi: data lama dianggap match dengan code yang terdeteksi
                cursor.execute("UPDATE detected_codes SET target_session = code WHERE target_session IS NULL")
            except Exception as e:
                # Silent fail jika ada error
                pass
    
    conn.commit() #Commit semua perubahan ke database (save changes)
    
    conn.close() #Tutup koneksi database


def load_existing_data(current_date):
    # Fungsi memuat data dari database berdasarkan tanggal | Tujuan: Load semua deteksi yang sesuai dengan tanggal hari ini
    # Parameter: current_date = datetime.date object untuk tanggal yang ingin diload
    # Return: List of dict berisi data deteksi dengan keys: ID, Time, Code, Type, ImagePath, Status, TargetSession
    
    detected_codes = [] #Inisialisasi list kosong untuk menyimpan hasil query
    
    # Format current_date menjadi string YYYY-MM-DD untuk LIKE query
    today_date_str = current_date.strftime("%Y-%m-%d")
    
    try:
        # Buka koneksi ke database
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Check kolom apa saja yang ada di table untuk backward compatibility
        # Ini penting karena database lama mungkin belum punya kolom status/target_session
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]
        
        # Flag untuk cek apakah kolom status dan target_session ada
        has_status = 'status' in columns
        has_target_session = 'target_session' in columns
        
        # Load data dengan schema yang tersedia
        # Build query berbeda tergantung kolom yang ada
        
        if has_status and has_target_session:
            # Full schema dengan status dan target_session
            # Query untuk load semua data hari ini (LIKE '%' untuk match YYYY-MM-DD HH:MM:SS)
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path, status, target_session FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            
            # Iterate setiap row hasil query dan convert ke dictionary
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0],  # id dari database
                    'Time': row[1],  # timestamp (YYYY-MM-DD HH:MM:SS)
                    'Code': row[2],  # code yang terdeteksi (e.g., 55D23L)
                    'Type': row[3],  # preset (JIS/DIN)
                    'ImagePath': row[4],  # path ke file gambar
                    'Status': row[5] if row[5] else 'OK',  # status (default 'OK' jika NULL)
                    'TargetSession': row[6] if row[6] else row[2]  # target session (default code jika NULL)
                })
        elif has_status:
            # Schema tanpa target_session (database agak lama)
            # Query hanya SELECT kolom yang ada
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path, status FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            
            # Iterate dan append dengan target_session = code sebagai fallback
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0],
                    'Time': row[1],
                    'Code': row[2],
                    'Type': row[3],
                    'ImagePath': row[4],
                    'Status': row[5] if row[5] else 'OK',
                    'TargetSession': row[2]  # Gunakan code sebagai target_session
                })
        else:
            # Schema minimal tanpa status dan target_session (database sangat lama)
            # Query hanya kolom dasar
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            
            # Iterate dan append dengan default values untuk status dan target_session
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0],
                    'Time': row[1],
                    'Code': row[2],
                    'Type': row[3],
                    'ImagePath': row[4],
                    'Status': 'OK',  # Default semua OK
                    'TargetSession': row[2]  # Gunakan code sebagai target_session
                })
        
        conn.close() #Tutup koneksi database
        
        return detected_codes #Return list of dictionaries berisi data deteksi
        
    except Exception as e:
        # Jika ada error, print error message dan return list kosong
        print(f"Error loading data: {e}")
        return detected_codes


def delete_codes(record_ids):
    # Fungsi menghapus data berdasarkan daftar ID | Tujuan: Hapus data dan file gambar terkait dari database dan disk
    # Parameter: record_ids = List of integer IDs yang akan dihapus
    # Return: Boolean True jika berhasil, False jika gagal
    
    # Validasi: jika list kosong, return False
    if not record_ids:
        return False

    try:
        # Buka koneksi database
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Buat placeholders untuk SQL IN clause
        # Jika record_ids = [1, 2, 3], placeholders = '?,?,?'
        placeholders = ','.join('?' for _ in record_ids)
        
        # Ambil image paths sebelum delete (untuk dihapus dari disk)
        # Penting ambil dulu sebelum DELETE karena setelah delete data hilang
        cursor.execute(f"SELECT image_path FROM detected_codes WHERE id IN ({placeholders})", record_ids)
        image_paths = cursor.fetchall()  # List of tuples [(path1,), (path2,), ...]
        
        # Delete records dari database
        # WHERE id IN (?, ?, ?) dengan values dari record_ids
        cursor.execute(f"DELETE FROM detected_codes WHERE id IN ({placeholders})", record_ids)
        
        conn.commit() #Commit perubahan ke database (save delete operation)
        
        conn.close() #Tutup koneksi database

        import os #Import os untuk file operations
        
        # Iterate setiap path dan hapus file jika exist
        for path_tuple in image_paths:
            image_path = path_tuple[0]  # Extract string path dari tuple
            
            # Check apakah path valid dan file exist
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path) #Hapus file dari disk
                except Exception as file_e:
                    # Print warning jika gagal hapus file (file locked, permission denied, dll)
                    # Tapi proses delete tetap lanjut untuk file lain
                    print(f"Warning: Gagal menghapus file gambar {image_path}: {file_e}")

        return True #Return True jika semua proses berhasil

    except Exception as e:
        # Jika ada error di database operation, print error dan return False
        print(f"Error deleting data: {e}")
        return False


def insert_detection(timestamp, code, preset, image_path, status, target_session):
    # Fungsi insert deteksi baru ke database | Tujuan: Simpan informasi lengkap deteksi ke database
    # Parameter: timestamp (format YYYY-MM-DD HH:MM:SS), code, preset (JIS/DIN), image_path, status (OK/Not OK), target_session
    # Return: Integer ID baru dari inserted record, atau None jika gagal
    
    try:
        # Buka koneksi database
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Execute INSERT statement dengan parameterized query (untuk SQL injection protection)
        # ? adalah placeholder yang akan diganti dengan values dari tuple parameter
        cursor.execute("INSERT INTO detected_codes (timestamp, code, preset, image_path, status, target_session) VALUES (?, ?, ?, ?, ?, ?)",
                      (timestamp, code, preset, image_path, status, target_session))
        
        # Ambil ID dari row yang baru saja di-insert
        # lastrowid adalah auto-generated ID dari AUTOINCREMENT column
        new_id = cursor.lastrowid
        
        conn.commit() #Commit perubahan ke database (save insert operation)
        
        conn.close() #Tutup koneksi database
        
        return new_id #Return ID baru untuk reference (bisa digunakan untuk update UI atau operasi lanjutan)
        
    except Exception as e:
        # Jika ada error saat insert, print error message dan return None
        print(f"Error inserting detection: {e}")
        return None


def get_detection_count(db_file=None):
    # Fungsi dapatkan jumlah total deteksi di database | Tujuan: Hitung total records untuk validasi sebelum export
    # Parameter: db_file = String path ke database file (default: DB_FILE dari config)
    # Return: Integer total jumlah deteksi
    
    # Jika db_file tidak diberikan, gunakan DB_FILE default dari config
    if db_file is None:
        db_file = DB_FILE
        
    try:
        # Buka koneksi ke database
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # Execute COUNT query untuk hitung total records di table
        # COUNT(*) menghitung semua rows termasuk yang punya NULL values
        cursor.execute("SELECT COUNT(*) FROM detected_codes")
        
        # Ambil hasil query (COUNT selalu return 1 row dengan 1 column)
        count = cursor.fetchone()[0]  # fetchone() return tuple, [0] untuk ambil value pertama
        
        conn.close() #Tutup koneksi database
        
        return count #Return count sebagai integer
        
    except Exception as e:
        # Jika ada error (misalnya table tidak ada), print error dan return 0
        print(f"Error getting count: {e}")
        return 0