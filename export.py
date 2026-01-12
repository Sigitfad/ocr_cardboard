"""
Fungsi export data ke Excel dengan format dan styling yang sesuai.
Menangani konversi data dari database SQLite ke file Excel dengan formatting lengkap.
"""

import os  # File system operations | Modul untuk operasi file
import sqlite3  # Database operations | Library untuk database queries
import pandas as pd  # Data manipulation | Library untuk manipulasi data/dataframe
import xlsxwriter  # Excel file writing | Library untuk membuat dan format Excel file
import tempfile  # Temporary file handling | Modul untuk buat temporary files
from datetime import datetime  # Date/time operations | Modul untuk date/time
from PIL import Image, ImageDraw, ImageFont  # Image processing | Library untuk image manipulation
from config import DB_FILE, Resampling  # Database file dan resampling method | Import dari config


def execute_export(sql_filter="", date_range_desc="", export_label="", current_preset=""):
    # Fungsi utama untuk mengeksekusi proses export ke Excel dengan filter | Tujuan: Create Excel file dari database dengan styling dan gambar
    # Parameter: sql_filter (WHERE clause), date_range_desc (deskripsi range), export_label (label filter), current_preset (JIS/DIN)
    # Return: String path ke file Excel yang dibuat, atau error message jika gagal
    
    excel_filename = f"Karton_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    from config import EXCEL_DIR
    output_path = os.path.join(EXCEL_DIR, excel_filename)

    temp_files_to_clean = []

    try:
        conn = sqlite3.connect(DB_FILE)
        
        cursor = conn.cursor()
        # Check kolom apa saja yang ada di table
        cursor.execute("PRAGMA table_info(detected_codes)")
        columns = [column[1] for column in cursor.fetchall()]
        has_status = 'status' in columns
        has_target_session = 'target_session' in columns
        
        # Build query berdasarkan schema yang ada
        if has_status and has_target_session:
            query = f"SELECT timestamp, code, preset, image_path, status, target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        elif has_status:
            query = f"SELECT timestamp, code, preset, image_path, status, code as target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        else:
            query = f"SELECT timestamp, code, preset, image_path, 'OK' as status, code as target_session FROM detected_codes {sql_filter} ORDER BY timestamp ASC"
        
        # Load data ke pandas DataFrame
        df = pd.read_sql_query(query, conn)
        conn.close()
        
        # Jika tidak ada data, return early
        if df.empty:
            return "NO_DATA"

        # Gunakan current_preset dari parameter
        export_preset = current_preset if current_preset else "Mixed"
        
        # Jika tidak ada preset yang diberikan, deteksi dari data
        if not current_preset:
            if 'preset' in df.columns and not df['preset'].empty:
                unique_presets = df['preset'].unique()
                if len(unique_presets) == 1:
                    export_preset = unique_presets[0]
                else:
                    export_preset = df['preset'].mode()[0] if not df['preset'].mode().empty else "Mixed"

        # Tentukan label untuk display
        if export_label and export_label != "All Label":
            label_display = export_label
        else:
            label_display = "All Labels"

        # Hitung statistik
        qty_actual = len(df)
        qty_ok = len(df[df['status'] == 'OK'])
        qty_not_ok = len(df[df['status'] == 'Not OK'])
        
        # FIXED: Start data di row 7 (karena info header 1-6)
        START_ROW_DATA = 7

        # Prepare data untuk Excel
        df['timestamp'] = pd.to_datetime(df['timestamp'])
        df.insert(0, 'No', range(1, 1 + len(df)))
        df['Image'] = ""  # Placeholder untuk image column
        
        # Rename columns ke nama yang user-friendly
        df.rename(columns={
            'timestamp': 'Date/Time',
            'code': 'Label',
            'preset': 'Standard',
            'image_path': 'Image Path',
            'status': 'Status',
            'target_session': 'Target Session'
        }, inplace=True)
        
        # Reorder columns
        df = df[['No', 'Image', 'Label', 'Date/Time', 'Standard', 'Status', 'Image Path', 'Target Session']]

        # Buat Excel file dengan xlsxwriter
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        sheet_name = datetime.now().strftime("%Y-%m-%d")
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=START_ROW_DATA)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Define format untuk header
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3'})
        
        # Define format untuk info rows (1-6)
        info_merge_format = workbook.add_format({
            'bold': True, 'align': 'left', 'valign': 'vleft', 'font_size': 11
        })

        # Define format untuk data cells
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        datetime_center_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        # Define format untuk "Not OK" rows (red background)
        not_ok_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
        not_ok_datetime_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
        
        # Row 1: Date Range
        date_text = f"Date : {date_range_desc}"
        worksheet.merge_range('A1:B1', date_text, info_merge_format)
        
        # Row 2: Type
        type_text = f"Type : {export_preset}"
        worksheet.merge_range('A2:B2', type_text, info_merge_format)
        
        # Row 3: Label
        label_text = f"Label : {label_display}"
        worksheet.merge_range('A3:B3', label_text, info_merge_format)
        
        # Row 4: Status OK
        status_ok_text = f"OK : {qty_ok}"
        worksheet.merge_range('A4:B4', status_ok_text, info_merge_format)
        
        # Row 5: Status Not OK
        status_not_ok_text = f"Not OK : {qty_not_ok}"
        worksheet.merge_range('A5:B5', status_not_ok_text, info_merge_format)
        
        # Row 6: QTY Actual
        qty_text = f"QTY Actual : {qty_actual}"
        worksheet.merge_range('A6:B6', qty_text, info_merge_format)
        
        # Row 7: Table Headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(START_ROW_DATA - 1, col_num, value, header_format)
        
        # Set column widths
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 10)
        worksheet.set_column('F:F', 10)
        worksheet.set_column('G:G', 0, options={'hidden': True})  # Hide path column
        worksheet.set_column('H:H', 0, options={'hidden': True})  # Hide target session column

        # Iterate setiap row data dan tulis ke Excel
        for row_num, row_data in df.iterrows():
            excel_row = row_num + START_ROW_DATA
            image_path = row_data['Image Path']
            status = row_data['Status']
            
            # Gunakan format berbeda untuk Not OK rows
            cell_format = not_ok_format if status == 'Not OK' else center_format
            datetime_format = not_ok_datetime_format if status == 'Not OK' else datetime_center_format

            try:
                worksheet.write(excel_row, 0, row_data['No'], cell_format)
            except Exception:
                worksheet.write(excel_row, 0, row_num + 1, cell_format)
            worksheet.write(excel_row, 1, '', cell_format)
            
            # Insert image jika ada
            if os.path.exists(image_path):
                temp_dir = tempfile.gettempdir()
                thumbnail_filename = f"app_temp_thumb_{os.getpid()}_{row_num}.png"
                thumbnail_path = os.path.join(temp_dir, thumbnail_filename)
                temp_files_to_clean.append(thumbnail_path)
                
                try:
                    max_col_b_px = int(30 * 7)
                    target_row_max_height = 150

                    # Load gambar dan resize
                    img = Image.open(image_path).convert("RGB")

                    # Draw detected label pada gambar
                    draw = ImageDraw.Draw(img)
                    try:
                        font = ImageFont.truetype("arial.ttf", 30)
                    except IOError:
                        font = ImageFont.load_default()

                    text_display = f"Detected: {row_data['Label']}"

                    bbox = draw.textbbox((10, img.height - 50), text_display, font=font)
                    draw.rectangle([bbox[0]-5, bbox[1]-5, bbox[2]+5, bbox[3]+5], fill=(0, 0, 0, 100))
                    draw.text((15, img.height - 50), text_display, fill=(255, 255, 0), font=font)

                    # Calculate scaling untuk fit dalam Excel cell
                    width_percent = (target_row_max_height / float(img.size[1]))
                    target_width = int(float(img.size[0]) * width_percent)
                    target_height = target_row_max_height

                    if target_width > max_col_b_px:
                        scale = max_col_b_px / float(img.size[0])
                        target_width = max_col_b_px
                        target_height = int(float(img.size[1]) * scale)

                    # Set row height untuk accommodate image
                    worksheet.set_row(excel_row, target_height)
                    
                    # Resize dan save thumbnail
                    img_resized = img.resize((target_width, target_height), Resampling)
                    img_resized.save(thumbnail_path, format='PNG')

                    # Calculate offset untuk center image di cell
                    x_offset = max(0, (max_col_b_px - target_width) // 2 + 5)
                    y_offset = max(0, (target_row_max_height - target_height) // 2)

                    # Insert image ke Excel
                    worksheet.insert_image(excel_row, 1, thumbnail_path, {'x_scale': 1, 'y_scale': 1, 'x_offset': x_offset, 'y_offset': y_offset})
                
                except Exception as img_e:
                    print(f"Warning: Gagal memproses atau menyisipkan gambar untuk baris {row_num}: {img_e}")
                
            # Write data columns
            worksheet.write(excel_row, 2, row_data['Label'], cell_format)
            worksheet.write_datetime(excel_row, 3, row_data['Date/Time'], datetime_format)
            worksheet.write(excel_row, 4, row_data['Standard'], cell_format)
            worksheet.write(excel_row, 5, row_data['Status'], cell_format)
            worksheet.write(excel_row, 6, row_data['Image Path'], cell_format)
            worksheet.write(excel_row, 7, row_data['Target Session'], cell_format)

        # Close Excel writer
        writer.close()

        # Cleanup temporary files
        for t_path in temp_files_to_clean:
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    pass

        return output_path

    except Exception as e:
        print(f"Export error: {e}")
        # Cleanup temp files jika terjadi error
        for t_path in temp_files_to_clean:
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    pass
        return f"EXPORT_ERROR: {e}"