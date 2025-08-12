import pandas as pd
import qrcode
import os
import re
import shutil  # untuk membuat zip

file_path = "qr code guru madarul.xls"
output_folder = "QR PTK DARUL MUJAHIDIN"

def clean_filename(text):
    text = str(text).strip()
    text = re.sub(r'[<>:"/\\|?*]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text

os.makedirs(output_folder, exist_ok=True)

df = pd.read_excel(file_path, header=2, engine="xlrd")
df.columns = df.columns.str.strip()

nama_col = "Nama Peserta"
kode_col = "QR-Code"
kelas_col = "Kelas"

if not all(col in df.columns for col in [nama_col, kode_col, kelas_col]):
    raise ValueError(f"Kolom '{nama_col}', '{kode_col}', atau '{kelas_col}' tidak ditemukan di file Excel.")

for _, row in df.iterrows():
    nama = clean_filename(row[nama_col])
    kode = str(row[kode_col]).strip()
    kelas = clean_filename(row[kelas_col])

    if not kode:
        continue

    kelas_folder = os.path.join(output_folder, kelas)
    os.makedirs(kelas_folder, exist_ok=True)

    img = qrcode.make(kode)
    save_path = os.path.join(kelas_folder, f"{nama}.png")
    img.save(save_path)

# Membuat file ZIP dari folder output
zip_name = f"{output_folder}.zip"
shutil.make_archive(output_folder, 'zip', output_folder)

print(f"✅ QR Code berhasil dibuat, sudah dipisah per kelas, dan dikompres menjadi '{zip_name}'!")
