# Operasi database SQLite untuk menyimpan dan mengambil data deteksi
# File ini berisi semua fungsi database untuk CRUD operations dan migrasi schema

import sqlite3  # SQLite database library | Library untuk SQL database operations
from datetime import datetime  # Date/time operations | Modul untuk date/time handling
from config import DB_FILE  # Import database file path dari config | Path database file


def setup_database():
    # Fungsi setup database dan buat table jika belum ada | Tujuan: Inisialisasi database dengan schema yang benar
    # Juga melakukan migration untuk menambah kolom baru jika diperlukan
    
    conn = sqlite3.connect(DB_FILE)  # Koneksi ke database SQLite
    cursor = conn.cursor()
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='detected_codes'")
    table_exists = cursor.fetchone() is not None
    
    if not table_exists:
        # Buat table baru jika tidak ada dengan schema yang lengkap
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
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]
        
        # Tambah kolom 'status' jika belum ada (untuk tracking OK/Not OK)
        if 'status' not in columns:
            try:
                cursor.execute("ALTER TABLE detected_codes ADD COLUMN status TEXT DEFAULT 'OK'")
                cursor.execute("UPDATE detected_codes SET status = 'OK' WHERE status IS NULL")
            except Exception as e:
                pass
        
        # Tambah kolom 'target_session' jika belum ada (untuk tracking session/label target)
        if 'target_session' not in columns:
            try:
                cursor.execute("ALTER TABLE detected_codes ADD COLUMN target_session TEXT")
                cursor.execute("UPDATE detected_codes SET target_session = code WHERE target_session IS NULL")
            except Exception as e:
                pass
                        
    conn.commit()
    conn.close()


def load_existing_data(current_date):
    # Fungsi memuat data dari database berdasarkan tanggal | Tujuan: Load semua deteksi yang sesuai dengan tanggal hari ini
    # Parameter: current_date = datetime.date object untuk tanggal yang ingin diload
    # Return: List of dict berisi data deteksi dengan keys: ID, Time, Code, Type, ImagePath, Status, TargetSession
    
    detected_codes = []
    today_date_str = current_date.strftime("%Y-%m-%d")
    
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]
        has_status = 'status' in columns
        has_target_session = 'target_session' in columns
        
        # Load data dengan schema yang tersedia
        if has_status and has_target_session:
            # Full schema dengan status dan target_session
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path, status, target_session FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0], 
                    'Time': row[1], 
                    'Code': row[2], 
                    'Type': row[3],
                    'ImagePath': row[4],
                    'Status': row[5] if row[5] else 'OK',
                    'TargetSession': row[6] if row[6] else row[2]
                })
        elif has_status:
            # Schema tanpa target_session
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path, status FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0], 
                    'Time': row[1], 
                    'Code': row[2], 
                    'Type': row[3],
                    'ImagePath': row[4],
                    'Status': row[5] if row[5] else 'OK',
                    'TargetSession': row[2]
                })
        else:
            # Schema minimal tanpa status dan target_session
            cursor.execute(f"SELECT id, timestamp, code, preset, image_path FROM detected_codes WHERE timestamp LIKE '{today_date_str}%' ORDER BY timestamp ASC")
            for row in cursor.fetchall():
                detected_codes.append({
                    'ID': row[0], 
                    'Time': row[1], 
                    'Code': row[2], 
                    'Type': row[3],
                    'ImagePath': row[4],
                    'Status': 'OK',
                    'TargetSession': row[2]
                })
        
        conn.close()
        return detected_codes
    except Exception as e:
        print(f"Error loading data: {e}")
        return detected_codes


def delete_codes(record_ids):
    # Fungsi menghapus data berdasarkan daftar ID | Tujuan: Hapus data dan file gambar terkait dari database dan disk
    # Parameter: record_ids = List of integer IDs yang akan dihapus
    # Return: Boolean True jika berhasil, False jika gagal
    
    if not record_ids:
        return False

    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        placeholders = ','.join('?' for _ in record_ids)
        
        # Ambil image paths sebelum delete (untuk dihapus dari disk)
        cursor.execute(f"SELECT image_path FROM detected_codes WHERE id IN ({placeholders})", record_ids)
        image_paths = cursor.fetchall()
        
        # Delete records dari database
        cursor.execute(f"DELETE FROM detected_codes WHERE id IN ({placeholders})", record_ids)
        conn.commit()
        conn.close()

        # Hapus file gambar dari disk
        import os
        for path_tuple in image_paths:
            image_path = path_tuple[0]
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                except Exception as file_e:
                    print(f"Warning: Gagal menghapus file gambar {image_path}: {file_e}")

        return True

    except Exception as e:
        print(f"Error deleting data: {e}")
        return False


def insert_detection(timestamp, code, preset, image_path, status, target_session):
    # Fungsi insert deteksi baru ke database | Tujuan: Simpan informasi lengkap deteksi ke database
    # Parameter: timestamp (format YYYY-MM-DD HH:MM:SS), code, preset (JIS/DIN), image_path, status (OK/Not OK), target_session
    # Return: Integer ID baru dari inserted record, atau None jika gagal
    
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        cursor.execute("INSERT INTO detected_codes (timestamp, code, preset, image_path, status, target_session) VALUES (?, ?, ?, ?, ?, ?)",
                      (timestamp, code, preset, image_path, status, target_session))
        
        new_id = cursor.lastrowid
        
        conn.commit()
        conn.close()
        
        return new_id
    except Exception as e:
        print(f"Error inserting detection: {e}")
        return None


def get_detection_count(db_file=None):
    # Fungsi dapatkan jumlah total deteksi di database | Tujuan: Hitung total records untuk validasi sebelum export
    # Parameter: db_file = String path ke database file (default: DB_FILE dari config)
    # Return: Integer total jumlah deteksi
    
    if db_file is None:
        db_file = DB_FILE
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM detected_codes")
        count = cursor.fetchone()[0]
        
        conn.close()
        return count
    except Exception as e:
        print(f"Error getting count: {e}")
        return 0