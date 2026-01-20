import os  #Import modul untuk operasi file dan direktori
import sqlite3  #Import modul untuk koneksi dan query database SQLite
import pandas as pd  #Import pandas untuk manipulasi data dalam bentuk DataFrame
import xlsxwriter  #Import xlsxwriter untuk membuat dan memformat file Excel
import tempfile  #Import tempfile untuk membuat file temporary saat processing image
from datetime import datetime  #Import datetime untuk tanggal dan waktu
from PIL import Image, ImageDraw, ImageFont  #Import PIL (Pillow) untuk image processing (load, resize, draw text pada image)
from config import DB_FILE, Resampling  #Import konfigurasi database file dan resampling method dari config.py


def execute_export(sql_filter="", date_range_desc="", export_label="", current_preset=""):
    # Fungsi utama untuk mengeksekusi proses export ke Excel dengan filter | Tujuan: Create Excel file dari database dengan styling dan gambar
    # Parameter: sql_filter (WHERE clause), date_range_desc (deskripsi range), export_label (label filter), current_preset (JIS/DIN)
    # Return: String path ke file Excel yang dibuat, atau error message jika gagal
    
    # Generate nama file Excel dengan timestamp agar unik
    # Format: Karton_Report_YYYYMMDD_HHMMSS.xlsx
    excel_filename = f"Karton_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    from config import EXCEL_DIR #Import direktori output Excel dari config
    
    output_path = os.path.join(EXCEL_DIR, excel_filename) #Gabungkan path direktori dengan nama file untuk mendapat full path

    # List untuk menyimpan path file temporary yang perlu dibersihkan nanti
    # File temporary dibuat saat resize image untuk Excel
    temp_files_to_clean = []

    try:
        conn = sqlite3.connect(DB_FILE) #Buka koneksi ke database SQLite
        cursor = conn.cursor() #Buat cursor untuk execute query
        
        # Check kolom apa saja yang ada di table
        # PRAGMA table_info mengembalikan informasi struktur tabel
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]  # Ambil nama kolom (index 1)
        
        # Cek apakah kolom 'status' dan 'target_session' ada di tabel
        # Ini untuk backward compatibility dengan database lama yang mungkin belum punya kolom ini
        has_status = 'status' in columns
        has_target_session = 'target_session' in columns
        
        # Build query berdasarkan schema yang ada
        # Jika kolom status dan target_session ada, gunakan kolom asli
        if has_status and has_target_session:
            query = f"SELECT timestamp, code, preset, image_path, status, target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        # Jika hanya status yang ada, gunakan 'code' sebagai fallback untuk target_session
        elif has_status:
            query = f"SELECT timestamp, code, preset, image_path, status, code as target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        # Jika tidak ada status dan target_session, gunakan default 'OK' dan 'code'
        else:
            query = f"SELECT timestamp, code, preset, image_path, 'OK' as status, code as target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        
        df = pd.read_sql_query(query, conn) #Load data ke pandas DataFrame
        
        conn.close() #Tutup koneksi database setelah selesai query
        
        # Jika tidak ada data, return early dengan kode "NO_DATA"
        # Parent function akan handle message ini untuk tampilkan ke user
        if df.empty:
            return "NO_DATA"

        # Gunakan current_preset dari parameter
        # Preset menentukan standard battery (JIS/DIN)
        export_preset = current_preset if current_preset else "Mixed"
        
        # Jika tidak ada preset yang diberikan, deteksi dari data
        if not current_preset:
            # Cek apakah kolom 'preset' ada di DataFrame
            if 'preset' in df.columns and not df['preset'].empty:
                unique_presets = df['preset'].unique() #Ambil semua unique preset dari data
                # Jika hanya 1 preset, gunakan itu
                if len(unique_presets) == 1:
                    export_preset = unique_presets[0]
                else:
                    # Jika mixed preset, gunakan yang paling sering muncul (mode)
                    export_preset = df['preset'].mode()[0] if not df['preset'].mode().empty else "Mixed"

        # Tentukan label untuk display di Excel header
        # Jika ada filter label spesifik, tampilkan nama label
        # Jika "All Label", tampilkan "All Labels"
        if export_label and export_label != "All Label":
            label_display = export_label
        else:
            label_display = "All Labels"

        # Hitung statistik untuk ditampilkan di header Excel
        qty_actual = len(df)  # Total semua data
        qty_ok = len(df[df['status'] == 'OK'])  # Jumlah data dengan status OK
        qty_not_ok = len(df[df['status'] == 'Not OK'])  # Jumlah data dengan status Not OK
        
        # FIXED: Start data di row 7 (karena info header 1-6)
        # Row 1-6 digunakan untuk info header (Date, Type, Label, OK, Not OK, QTY)
        # Row 7 adalah header tabel (No, Image, Label, Date/Time, dll)
        # Row 8 dst adalah data
        START_ROW_DATA = 7

        # Prepare data untuk Excel
        # Konversi kolom timestamp ke datetime object untuk formatting yang benar
        df['timestamp'] = pd.to_datetime(df['timestamp'])
        
        # Insert kolom 'No' di posisi pertama dengan nomor urut 1, 2, 3, ...
        df.insert(0, 'No', range(1, 1 + len(df)))
        
        # Insert kolom 'Image' placeholder (akan diisi dengan gambar actual nanti)
        df['Image'] = ""  # Placeholder untuk image column
        
        # Rename columns ke nama yang user-friendly
        # Nama internal database -> Nama untuk display di Excel
        df.rename(columns={
            'timestamp': 'Date/Time',  # timestamp -> Date/Time
            'code': 'Label',  # code -> Label
            'preset': 'Standard',  # preset -> Standard
            'image_path': 'Image Path',  # image_path -> Image Path
            'status': 'Status',  # status -> Status
            'target_session': 'Target Session'  # target_session -> Target Session
        }, inplace=True)
        
        # Reorder columns agar sesuai urutan yang diinginkan di Excel
        # Urutan: No, Image, Label, Date/Time, Standard, Status, Image Path (hidden), Target Session (hidden)
        df = df[['No', 'Image', 'Label', 'Date/Time', 'Standard', 'Status', 'Image Path', 'Target Session']]

        # Buat Excel file dengan xlsxwriter engine
        # xlsxwriter memberikan kontrol penuh atas formatting, styling, dan insert image
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        
        sheet_name = datetime.now().strftime("%Y-%m-%d") #Nama sheet menggunakan tanggal hari ini (YYYY-MM-DD)
        
        # Write DataFrame ke Excel tanpa header dan index
        # header=False karena kita akan write header manual dengan format custom
        # startrow=START_ROW_DATA (row 7) karena row 1-6 untuk info header
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=START_ROW_DATA)
        
        # Ambil workbook dan worksheet object untuk formatting
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Define format untuk header tabel (row 7)
        # Bold, centered, abu-abu background
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3'})
        
        # Define format untuk info rows (1-6)
        # Bold, aligned left, font size 11
        info_merge_format = workbook.add_format({
            'bold': True, 'align': 'left', 'valign': 'vleft', 'font_size': 11
        })

        # Define format untuk data cells
        # Centered dengan border
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        # Format khusus untuk datetime cells
        # Format: yyyy-mm-dd hh:mm:ss, centered, border
        datetime_center_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        # Define format untuk "Not OK" rows (red background)
        # Background merah, text putih untuk highlight error
        not_ok_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
        
        # Format datetime untuk Not OK rows
        not_ok_datetime_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
        
        # Row 1: Date Range
        # Merge cell A1:B1 dan tulis info tanggal export
        date_text = f"Date : {date_range_desc}"
        worksheet.merge_range('A1:B1', date_text, info_merge_format)
        
        # Row 2: Type
        # Merge cell A2:B2 dan tulis tipe battery (JIS/DIN/Mixed)
        type_text = f"Type : {export_preset}"
        worksheet.merge_range('A2:B2', type_text, info_merge_format)
        
        # Row 3: Label
        # Merge cell A3:B3 dan tulis label/sesi yang diexport
        label_text = f"Label : {label_display}"
        worksheet.merge_range('A3:B3', label_text, info_merge_format)
        
        # Row 4: Status OK
        # Merge cell A4:B4 dan tulis jumlah data OK
        status_ok_text = f"OK : {qty_ok}"
        worksheet.merge_range('A4:B4', status_ok_text, info_merge_format)
        
        # Row 5: Status Not OK
        # Merge cell A5:B5 dan tulis jumlah data Not OK
        status_not_ok_text = f"Not OK : {qty_not_ok}"
        worksheet.merge_range('A5:B5', status_not_ok_text, info_merge_format)
        
        # Row 6: QTY Actual
        # Merge cell A6:B6 dan tulis total quantity semua data
        qty_text = f"QTY Actual : {qty_actual}"
        worksheet.merge_range('A6:B6', qty_text, info_merge_format)
        
        # Row 7: Table Headers
        # Write header tabel dengan format bold dan background abu-abu
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(START_ROW_DATA - 1, col_num, value, header_format)
        
        # Set column widths untuk setiap kolom
        # Lebar dalam satuan karakter (approx)
        worksheet.set_column('A:A', 5)  # No - narrow column
        worksheet.set_column('B:B', 30)  # Image - wide untuk accommodate gambar
        worksheet.set_column('C:C', 20)  # Label - medium
        worksheet.set_column('D:D', 25)  # Date/Time - cukup untuk format datetime
        worksheet.set_column('E:E', 10)  # Standard - narrow
        worksheet.set_column('F:F', 10)  # Status - narrow
        worksheet.set_column('G:G', 0, options={'hidden': True})  # Hide path column - tidak perlu ditampilkan
        worksheet.set_column('H:H', 0, options={'hidden': True})  # Hide target session column - tidak perlu ditampilkan

        # Iterate setiap row data dan tulis ke Excel
        # iterrows() mengembalikan (index, row_data) untuk setiap baris
        for row_num, row_data in df.iterrows():
            excel_row = row_num + START_ROW_DATA #Hitung posisi row di Excel (row_num dari DataFrame + START_ROW_DATA)
            
            # Ambil image path dan status dari data row
            image_path = row_data['Image Path']
            status = row_data['Status']
            
            # Gunakan format berbeda untuk Not OK rows
            # Jika Not OK, gunakan format merah. Jika OK, gunakan format normal
            cell_format = not_ok_format if status == 'Not OK' else center_format
            datetime_format = not_ok_datetime_format if status == 'Not OK' else datetime_center_format

            # Write kolom 'No' (nomor urut)
            try:
                worksheet.write(excel_row, 0, row_data['No'], cell_format)
            except Exception:
                # Fallback jika ada error, gunakan row_num + 1
                worksheet.write(excel_row, 0, row_num + 1, cell_format)
            
            worksheet.write(excel_row, 1, '', cell_format) #Write kolom 'Image' (placeholder kosong, gambar akan di-insert terpisah)
            
            # Insert image jika ada
            # Cek apakah file image benar-benar exist di path yang tersimpan
            if os.path.exists(image_path):
                temp_dir = tempfile.gettempdir() #Gunakan temp directory untuk simpan thumbnail hasil resize

                # Generate nama file thumbnail yang unik (include PID dan row_num)
                thumbnail_filename = f"app_temp_thumb_{os.getpid()}_{row_num}.png"
                thumbnail_path = os.path.join(temp_dir, thumbnail_filename)
                
                temp_files_to_clean.append(thumbnail_path) #Tambahkan ke list cleanup (akan dihapus setelah selesai)
                
                try:
                    # Hitung max width untuk kolom B dalam pixel
                    # 30 characters * 7 pixels per character (approx)
                    max_col_b_px = int(30 * 7)
                    
                    target_row_max_height = 150 #Target max height untuk row (dalam pixel)

                    # Load gambar dan konversi ke RGB
                    # Convert RGB diperlukan untuk ensure 3 channels (tanpa alpha)
                    img = Image.open(image_path).convert("RGB")

                    # Draw detected label pada gambar
                    # ImageDraw untuk menggambar text dan shape di image
                    draw = ImageDraw.Draw(img)
                    
                    # Load font untuk text (arial 30pt)
                    try:
                        font = ImageFont.truetype("arial.ttf", 30)
                    except IOError:
                        # Fallback ke default font jika arial.ttf tidak ada
                        font = ImageFont.load_default()

                    text_display = f"Detected: {row_data['Label']}" #Text yang akan ditampilkan di gambar

                    # Hitung bounding box untuk text (untuk background rectangle)
                    # Position text di bottom-left image (x=10, y=height-50)
                    bbox = draw.textbbox((10, img.height - 50), text_display, font=font)
                    
                    # Draw semi-transparent black rectangle sebagai background text
                    # Padding 5px dari text bbox
                    draw.rectangle([bbox[0]-5, bbox[1]-5, bbox[2]+5, bbox[3]+5], fill=(0, 0, 0, 100))
                    
                    # Draw text berwarna kuning di atas background hitam
                    draw.text((15, img.height - 50), text_display, fill=(255, 255, 0), font=font)

                    # Calculate scaling untuk fit dalam Excel cell
                    # Prioritas: fit height terlebih dahulu
                    width_percent = (target_row_max_height / float(img.size[1]))
                    target_width = int(float(img.size[0]) * width_percent)
                    target_height = target_row_max_height

                    # Jika setelah scale height, width melebihi max column width
                    # Scale ulang berdasarkan width
                    if target_width > max_col_b_px:
                        scale = max_col_b_px / float(img.size[0])
                        target_width = max_col_b_px
                        target_height = int(float(img.size[1]) * scale)

                    # Set row height untuk accommodate image
                    # Row height dalam points (1 point ≈ 1.33 pixels)
                    worksheet.set_row(excel_row, target_height)
                    
                    # Resize dan save thumbnail
                    # Resampling method dari config (biasanya LANCZOS untuk quality terbaik)
                    img_resized = img.resize((target_width, target_height), Resampling)
                    img_resized.save(thumbnail_path, format='PNG')

                    # Calculate offset untuk center image di cell
                    # Offset X: center horizontal dalam kolom B
                    x_offset = max(0, (max_col_b_px - target_width) // 2 + 5)
                    # Offset Y: center vertical dalam row
                    y_offset = max(0, (target_row_max_height - target_height) // 2)

                    # Insert image ke Excel di kolom B (index 1) pada row yang sesuai
                    # x_scale dan y_scale = 1 berarti no additional scaling
                    # x_offset dan y_offset untuk positioning dalam cell
                    worksheet.insert_image(excel_row, 1, thumbnail_path, {'x_scale': 1, 'y_scale': 1, 'x_offset': x_offset, 'y_offset': y_offset})
                
                except Exception as img_e:
                    # Jika ada error saat process/insert image, print warning
                    # Tapi proses export tetap lanjut (tidak critical error)
                    print(f"Warning: Gagal memproses atau menyisipkan gambar untuk baris {row_num}: {img_e}")
                
            # Write data columns lainnya
            # Kolom C: Label (kode battery yang terdeteksi)
            worksheet.write(excel_row, 2, row_data['Label'], cell_format)
            
            # Kolom D: Date/Time (timestamp dengan format datetime)
            worksheet.write_datetime(excel_row, 3, row_data['Date/Time'], datetime_format)
            
            # Kolom E: Standard (JIS/DIN)
            worksheet.write(excel_row, 4, row_data['Standard'], cell_format)
            
            # Kolom F: Status (OK/Not OK)
            worksheet.write(excel_row, 5, row_data['Status'], cell_format)
            
            # Kolom G: Image Path (hidden column - untuk reference jika perlu)
            worksheet.write(excel_row, 6, row_data['Image Path'], cell_format)
            
            # Kolom H: Target Session (hidden column - untuk reference jika perlu)
            worksheet.write(excel_row, 7, row_data['Target Session'], cell_format)

        writer.close() # Close Excel writer

        # Cleanup temporary files
        # Hapus semua file thumbnail yang dibuat selama proses
        for t_path in temp_files_to_clean:
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    # Jika gagal hapus (file locked), skip saja
                    pass

        return output_path #Return path file Excel yang berhasil dibuat

    except Exception as e:
        # Jika terjadi error di proses manapun, print error message
        print(f"Export error: {e}")
        
        # Cleanup temp files jika terjadi error
        # Penting untuk cleanup walau error untuk hindari file sampah
        for t_path in temp_files_to_clean:
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    pass
        
        return f"EXPORT_ERROR: {e}" #Return error message ke parent untuk ditampilkan ke user