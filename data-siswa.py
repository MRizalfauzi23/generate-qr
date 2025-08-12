import pandas as pd
from rapidfuzz import process, fuzz

# =========================
# KONFIGURASI
# =========================
file_excel = "data_siswa.xlsx"   # Satu file dengan dua sheet
sheet_lama = "Lama"              # Nama sheet data lama
sheet_baru = "Baru"              # Nama sheet data baru

kolom_nama = "Nama"
kolom_kelas = "Kelas"
kolom_nis = "NIS"
kolom_nisn = "NISN"

ambang_batas_otomatis = 90   # Skor minimal fuzzy match untuk dianggap cocok otomatis
# =========================


# Fungsi untuk membersihkan nama
def clean_name(name):
    if pd.isna(name):
        return ""
    return " ".join(str(name).strip().upper().split())


# Baca file lama & baru
old_df = pd.read_excel(file_excel, sheet_name=sheet_lama)
new_df = pd.read_excel(file_excel, sheet_name=sheet_baru)

# Standarisasi nama
old_df["NAMA_BERSIH"] = old_df[kolom_nama].apply(clean_name)
new_df["NAMA_BERSIH"] = new_df[kolom_nama].apply(clean_name)

# List untuk menyimpan hasil match
match_nama = []
match_score = []

# Loop setiap nama di data lama
for old_name in old_df["NAMA_BERSIH"]:
    match = process.extractOne(
        old_name,
        new_df["NAMA_BERSIH"],
        scorer=fuzz.token_sort_ratio
    )
    if match:
        matched_name, score, _ = match
        match_nama.append(matched_name)
        match_score.append(score)
    else:
        match_nama.append(None)
        match_score.append(0)

old_df["MATCH_NAMA"] = match_nama
old_df["MATCH_SCORE"] = match_score

# Gabungkan data lama dan data baru berdasarkan nama match
merged_df = pd.merge(
    old_df,
    new_df,
    left_on="MATCH_NAMA",
    right_on="NAMA_BERSIH",
    suffixes=("_LAMA", "_BARU"),
    how="left"
)

# Ambil data terbaru jika ada, kalau kosong ambil dari lama
merged_df["Kelas_Final"] = merged_df[f"{kolom_kelas}_BARU"].combine_first(merged_df[f"{kolom_kelas}_LAMA"])
merged_df["NIS_Final"] = merged_df[f"{kolom_nis}_BARU"].combine_first(merged_df[f"{kolom_nis}_LAMA"])
merged_df["NISN_Final"] = merged_df[f"{kolom_nisn}_BARU"].combine_first(merged_df[f"{kolom_nisn}_LAMA"])

# Pisahkan yang match otomatis dan yang perlu dicek manual
otomatis_df = merged_df[merged_df["MATCH_SCORE"] >= ambang_batas_otomatis]
manual_df = merged_df[merged_df["MATCH_SCORE"] < ambang_batas_otomatis]

# Simpan ke Excel
otomatis_df.to_excel("data_terupdate.xlsx", index=False)
manual_df.to_excel("perlu_cek_manual.xlsx", index=False)

print(f"✅ Selesai! Data terupdate disimpan di 'data_terupdate.xlsx', yang perlu dicek manual di 'perlu_cek_manual.xlsx'.")
