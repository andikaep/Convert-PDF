import pdfplumber
import pandas as pd
import os

# Path ke file PDF dan direktori penyimpanan Excel
pdf_path = "D:/pdf/12.pdf"
excel_dir = "D:/excel/"
excel_filename = "output.xlsx"

# Nama kolom yang akan digunakan
columns_names = [
    'No', 'Description', 'Date Acquired', 'Date Sold', 'Cost/ Basis',
    'Business Pct', '179 Bonus', 'Special Depr.', 'Prior 179 Bonus/ SP Depr.',
    'Prior Dec. Bal Depr.', 'Salvag/Basis Reduct', 'Depr. Basis',
    'Prior Depr.', 'Method', 'Life', 'Rate', 'Current Depr.'
]

# Fungsi untuk memproses setiap baris teks dari PDF
def extract_columns_from_line(line):
    parts = line.split()

    # Inisialisasi kolom dengan nilai default kosong
    columns = [''] * 17

    try:
        columns[0] = parts[0]  # No
        columns[1] = ' '.join(parts[1:parts.index("**") + 1])  # Description
        columns[2] = parts[parts.index("**") + 1]  # Date Acquired
        
        # Selanjutnya adalah data numerik, dipindahkan sesuai urutan yang benar
        numeric_data = parts[parts.index("**") + 2:]
        
        if len(numeric_data) >= 1:
            columns[4] = numeric_data[0]  # Cost/ Basis
        if len(numeric_data) >= 2:
            columns[8] = numeric_data[1]  # Prior 179 Bonus/ SP Depr.
        if len(numeric_data) >= 3:
            columns[9] = numeric_data[2]  # Prior Dec. Bal Depr.
        if len(numeric_data) >= 4:
            columns[10] = numeric_data[3]  # Salvag/Basis Reduct
        if len(numeric_data) >= 5:
            columns[11] = numeric_data[4]  # Depr. Basis
        if len(numeric_data) >= 6:
            columns[12] = numeric_data[5]  # Prior Depr.
        if len(numeric_data) >= 7:
            columns[13] = numeric_data[6]  # Method
        if len(numeric_data) >= 8:
            columns[14] = numeric_data[7]  # Life
        if len(numeric_data) >= 9:
            columns[15] = numeric_data[8]  # Rate
        if len(numeric_data) >= 10:
            columns[16] = numeric_data[9]  # Current Depr.
    except (ValueError, IndexError):
        pass

    return columns

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

# Buat DataFrame dari data yang diproses
df = pd.DataFrame(rows, columns=columns_names)

# Dapatkan nama file yang unik
unique_excel_path = get_unique_filename(excel_dir, excel_filename)

# Simpan DataFrame ke file Excel
df.to_excel(unique_excel_path, sheet_name="AllData", index=False)

print(f"File PDF telah berhasil dikonversi ke {unique_excel_path}")
