import camelot
import pandas as pd
import os
from openpyxl import load_workbook

# Fungsi untuk menghasilkan nama file baru jika file sudah ada
def get_available_filename(directory, filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base}({counter}){extension}"
        counter += 1
    return os.path.join(directory, new_filename)

# Tentukan path PDF dan direktori serta nama file output
pdf_path = 'D:/pdf/FS.pdf'
directory = 'D:/excel/'
base_filename = 'hasil_convert'

# Ekstrak nama file PDF tanpa ekstensi
pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

# Buat nama file Excel baru dengan menambahkan nama file PDF
filename = f"{base_filename}_{pdf_name}.xlsx"

# Tampilkan pesan bahwa proses konversi dimulai
print(f"Sedang proses convert file = {pdf_path}")

# Dapatkan nama file yang tersedia
output_path = get_available_filename(directory, filename)

# Baca tabel dari PDF menggunakan stream mode pada semua halaman
tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

# Cek berapa banyak tabel yang ditemukan
print(f"Total tabel yang ditemukan: {len(tables)}")

# Gabungkan semua tabel menjadi satu DataFrame
df_list = [table.df for table in tables]  # List of DataFrames
combined_df = pd.concat(df_list, ignore_index=True)  # Gabungkan semua DataFrame

# Simpan ke file Excel, pertahankan kolom kosong dengan menggantinya dengan spasi atau string kosong
combined_df.to_excel(output_path, index=False, na_rep='')

# Mengatur lebar kolom menjadi 170px
workbook = load_workbook(output_path)
worksheet = workbook.active

# Iterasi melalui semua kolom dan atur lebarnya
for col in worksheet.columns:
    max_length = 400 / 7  # Mengonversi pixel ke lebar kolom Excel (1 karakter ~ 7 piksel)
    col_letter = col[0].column_letter
    worksheet.column_dimensions[col_letter].width = max_length

# Simpan workbook yang telah diubah
workbook.save(output_path)

# Tampilkan pesan bahwa proses selesai
print(f"File berhasil dikonversi dan disimpan sebagai: {output_path} dengan lebar kolom 400px")