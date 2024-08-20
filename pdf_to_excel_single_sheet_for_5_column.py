import pdfplumber
import pandas as pd
from openpyxl.utils import get_column_letter
import os

# Path ke file PDF dan direktori penyimpanan Excel
pdf_path = "D:/pdf/TCB.pdf"
excel_dir = "D:/excel/"
excel_filename = "output17.xlsx"

print(f"Processing file: {pdf_path}")

# Fungsi untuk memisahkan baris menjadi kolom berdasarkan posisi teks
def extract_columns_from_line(line):
    parts = line.split()
    if len(parts) < 5:
        return [' '.join(parts), None, None, None, None]
    else:
        # Gabungkan semua bagian sebelum dua bagian terakhir sebagai kategori
        category = ' '.join(parts[:-4]).strip()
        amount1 = parts[-4].strip()
        amount2 = parts[-3].strip()
        amount3 = parts[-2].strip()
        amount4 = parts[-1].strip()
        return [category, amount1, amount2, amount3, amount4]

# Fungsi untuk mendapatkan nama file yang unik
def get_unique_filename(directory, filename):
    base, ext = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):
        unique_filename = f"{base}({counter}){ext}"
        counter += 1
    return os.path.join(directory, unique_filename)

# Buka file PDF dan ekstrak teks dari setiap halaman
rows = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.split('\n')
            for line in lines:
                if line.strip():  # Abaikan baris kosong
                    columns = extract_columns_from_line(line)
                    rows.append(columns)
            rows.append([None, None, None, None, None])  # Tambahkan baris kosong saat berganti halaman

# Buat DataFrame dari data yang diproses
df = pd.DataFrame(rows, columns=['Category', 'Amount1', 'Amount2', 'Amount3', 'Amount4'])

# Dapatkan nama file yang unik
unique_excel_path = get_unique_filename(excel_dir, excel_filename)

# Simpan DataFrame ke file Excel
df.to_excel(unique_excel_path, sheet_name="AllData", index=False)

# Mengatur lebar kolom menggunakan openpyxl
with pd.ExcelWriter(unique_excel_path, engine='openpyxl', mode='a') as writer:
    workbook = writer.book
    worksheet = workbook["AllData"]
    
    # Mengatur lebar kolom
    column_width = 45  # Lebar dalam karakter, setara dengan 450px
    for col_num in range(1, df.shape[1] + 1):
        column_letter = get_column_letter(col_num)
        worksheet.column_dimensions[column_letter].width = column_width

print(f"File PDF telah berhasil dikonversi ke {unique_excel_path} dengan lebar kolom 450px")