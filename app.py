import os
import re
import uuid
import shutil
import zipfile
import subprocess
from pathlib import Path
from flask import Flask, request, render_template, send_file, abort, redirect, url_for, flash

import pandas as pd
import qrcode
from werkzeug.utils import secure_filename

# OPTIONAL: rar support
try:
    import rarfile
    RAR_AVAILABLE = True
except Exception:
    RAR_AVAILABLE = False

app = Flask(__name__)
app.secret_key = "ganti-dengan-secret-random"  # untuk flash pesan
BASE_OUTPUT = Path("outputs")
BASE_OUTPUT.mkdir(exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

def clean_filename(text):
    text = str(text).strip()
    text = re.sub(r'[<>:"/\\|?*]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text

def read_excel_detect(path):
    ext = path.suffix.lower()
    if ext == ".xls":
        # xlrd needed
        return pd.read_excel(path, header=2, engine="xlrd")
    else:
        # xlsx or others
        return pd.read_excel(path, header=2, engine="openpyxl")

def zip_folder(folder_path: Path, zip_path: Path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for f in files:
                fp = Path(root) / f
                zipf.write(fp, fp.relative_to(folder_path))

def rar_with_winrar(folder_path: Path, rar_path: Path):
    """
    Try to create rar by calling rar (WinRAR) command line.
    Requires 'rar' or 'WinRAR' available in PATH.
    """
    # Build command: rar a -r output.rar folder_path\*
    cmd = ["rar", "a", "-r", str(rar_path), str(folder_path) + os.sep + "*"]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode == 0:
            return True
        else:
            app.logger.warning("RAR command failed: %s", proc.stderr)
            return False
    except FileNotFoundError:
        app.logger.warning("rar executable not found in PATH.")
        return False

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # file
        f = request.files.get("excel_file")
        if not f:
            flash("Silakan pilih file Excel terlebih dahulu.", "danger")
            return redirect(request.url)

        filename = secure_filename(f.filename)
        ext = Path(filename).suffix.lower()
        if ext not in ALLOWED_EXT:
            flash("Tipe file tidak didukung. Gunakan .xls atau .xlsx", "danger")
            return redirect(request.url)

        # user options
        output_name = request.form.get("output_name") or "QR PTK DARUL MUJAHIDIN"
        output_name = clean_filename(output_name)
        compress = request.form.get("compress")  # 'zip' or 'rar'

        # create workspace for this request
        req_id = uuid.uuid4().hex
        req_folder = BASE_OUTPUT / req_id
        req_folder.mkdir(parents=True, exist_ok=True)
        upload_path = req_folder / filename
        f.save(upload_path)

        # read excel
        try:
            df = read_excel_detect(upload_path)
        except Exception as e:
            flash(f"Gagal membaca file Excel: {e}", "danger")
            return redirect(request.url)

        # normalize columns
        df.columns = df.columns.str.strip()
        nama_col = "Nama Peserta"
        kode_col = "QR-Code"
        kelas_col = "Kelas"

        if not all(col in df.columns for col in [nama_col, kode_col, kelas_col]):
            flash(f"Kolom '{nama_col}', '{kode_col}', atau '{kelas_col}' tidak ditemukan.", "danger")
            return redirect(request.url)

        # output folder where QR images will be stored
        out_root = req_folder / output_name
        out_root.mkdir(exist_ok=True)

        # generate QR
        for _, row in df.iterrows():
            nama = clean_filename(row[nama_col])
            kode = str(row[kode_col]).strip()
            kelas = clean_filename(row[kelas_col])

            if not kode or pd.isna(kode) or len(kode.strip()) == 0:
                continue

            kelas_folder = out_root / kelas
            kelas_folder.mkdir(parents=True, exist_ok=True)

            img = qrcode.make(kode)
            save_path = kelas_folder / f"{nama}.png"
            img.save(save_path)

        # compress
        archive_name = req_folder / f"{output_name}"
        archive_path = None
        if compress == "zip":
            archive_path = req_folder / f"{output_name}.zip"
            zip_folder(out_root, archive_path)
        elif compress == "rar":
            # try rar via system exe
            rar_created = rar_with_winrar(out_root, req_folder / f"{output_name}.rar")
            if rar_created:
                archive_path = req_folder / f"{output_name}.rar"
            else:
                # fallback: if rarfile + system rar available maybe rarfile can handle
                if RAR_AVAILABLE:
                    try:
                        rf_path = req_folder / f"{output_name}.rar"
                        with rarfile.RarFile(rf_path, 'w') as rf:
                            for root, dirs, files in os.walk(out_root):
                                for file in files:
                                    file_path = Path(root) / file
                                    rf.write(file_path, arcname=file_path.relative_to(out_root))
                        archive_path = rf_path
                    except Exception as e:
                        app.logger.warning("rarfile failed: %s", e)
                        archive_path = None
                else:
                    archive_path = None

        # if compress failed or not chosen, fallback to zip
        if archive_path is None:
            archive_path = req_folder / f"{output_name}.zip"
            zip_folder(out_root, archive_path)
            flash("Pembuatan RAR gagal atau tidak tersedia, dibuat ZIP sebagai fallback.", "warning")

        # provide download
        return send_file(str(archive_path), as_attachment=True)

    # GET
    return render_template("index.html", rar_available=RAR_AVAILABLE)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
