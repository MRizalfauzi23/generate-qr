import os
import re
import uuid
import shutil
import zipfile
import subprocess
from pathlib import Path
from flask import Flask, request, render_template_string, send_file, redirect, flash

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
        return pd.read_excel(path, header=2, engine="xlrd")
    else:
        return pd.read_excel(path, header=2, engine="openpyxl")

def zip_folder(folder_path: Path, zip_path: Path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for f in files:
                fp = Path(root) / f
                zipf.write(fp, fp.relative_to(folder_path))

def rar_with_winrar(folder_path: Path, rar_path: Path):
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

# HTML template langsung di sini, menggunakan Jinja2 syntax
HTML_TEMPLATE = """
<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>QR Batch Generator</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 40px;
      background: #f0f3f7;
      color: #333;
    }
    .card {
      background: #fff;
      padding: 32px 28px;
      border-radius: 12px;
      max-width: 720px;
      margin: 0 auto;
      box-shadow: 0 8px 24px rgba(0,0,0,0.08);
      transition: box-shadow 0.3s ease;
    }
    .card:hover {
      box-shadow: 0 12px 36px rgba(0,0,0,0.12);
    }
    h1 {
      margin-top: 0;
      margin-bottom: 16px;
      font-weight: 700;
      font-size: 2rem;
      color: #0b74de;
      letter-spacing: 0.03em;
    }
    .note {
      font-size: 0.875rem;
      color: #6b7280;
      margin-top: 12px;
      line-height: 1.5;
    }
    label {
      display: block;
      margin-top: 24px;
      font-weight: 600;
      font-size: 1rem;
      color: #374151;
    }
    input[type="text"],
    input[type="file"],
    select {
      width: 100%;
      padding: 14px 16px;
      margin-top: 8px;
      border: 1.5px solid #d1d5db;
      border-radius: 10px;
      font-size: 1rem;
      transition: border-color 0.25s ease;
      box-sizing: border-box;
    }
    input[type="text"]:focus,
    input[type="file"]:focus,
    select:focus {
      outline: none;
      border-color: #0b74de;
      box-shadow: 0 0 6px rgba(11,116,222,0.3);
    }
    button {
      margin-top: 32px;
      width: 100%;
      padding: 14px 0;
      border-radius: 12px;
      border: none;
      background: #0b74de;
      color: white;
      font-weight: 700;
      font-size: 1.125rem;
      cursor: pointer;
      box-shadow: 0 4px 14px rgba(11,116,222,0.4);
      transition: background-color 0.3s ease, box-shadow 0.3s ease;
      user-select: none;
    }
    button:hover {
      background-color: #065bb5;
      box-shadow: 0 6px 20px rgba(6,91,181,0.6);
    }
    button:active {
      background-color: #044a8c;
      box-shadow: 0 2px 8px rgba(4,74,140,0.8);
    }
    .alert {
      padding: 14px 20px;
      margin-top: 20px;
      border-radius: 10px;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 12px;
      font-size: 0.95rem;
      user-select: none;
    }
    .alert-icon {
      font-size: 1.3rem;
    }
    .alert.danger {
      background: #fdecea;
      color: #b91c1c;
    }
    .alert.warning {
      background: #fef3c7;
      color: #78350f;
    }
    .alert.success {
      background: #d1fae5;
      color: #065f46;
    }
  </style>
</head>
<body>
  <div class="card">
    <h1>QR Generator Edulink</h1>
    <p class="note">
      Upload file Excel (.xls / .xlsx). Script akan baca header pada baris ke-3 (header=2). <strong>Pastikan Format Excel 97-2003 workbook ya gess!!</strong> kolom Harus ada: <strong>Nama Peserta</strong>, <strong>QR-Code</strong>, <strong>Kelas</strong>.
    </p>

    {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
      {% for cat, msg in messages %}
        <div class="alert {{ cat }}">
          <span class="alert-icon">
            {% if cat == 'danger' %}&#9888;{% elif cat == 'warning' %}&#x26A0;{% else %}&#10003;{% endif %}
          </span>
          <div class="alert-content">
            <strong>
              {% if cat == 'danger' %}Gagal:{% elif cat == 'warning' %}Perhatian:{% else %}Sukses:{% endif %}
            </strong>
            {{ msg }}
          </div>
        </div>
      {% endfor %}
    {% endif %}
    {% endwith %}

    <form method="post" enctype="multipart/form-data">
      <label for="excel_file">File Excel</label>
      <input type="file" id="excel_file" name="excel_file" accept=".xls,.xlsx" required>

      <label for="output_name">Nama folder output (opsional)</label>
      <input type="text" id="output_name" name="output_name" placeholder="QR CODE FOLDER">

      <label for="compress">Format kompresi</label>
      <select id="compress" name="compress">
        <option value="zip" selected>ZIP (direkomendasikan)</option>
        <option value="rar">RAR {% if not rar_available %}(tidak tersedia di server){% endif %}</option>
      </select>

      <button type="submit">Generate &amp; Download</button>
    </form>

    <p class="note">
      Catatan: Jika memilih RAR, server harus punya WinRAR/rar.exe atau support <code>rarfile</code>. Kalau tidak, sistem akan fallback ke ZIP.
    </p>
  </div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("excel_file")
        if not f:
            flash("Silakan pilih file Excel terlebih dahulu.", "danger")
            return redirect(request.url)

        filename = secure_filename(f.filename)
        ext = Path(filename).suffix.lower()
        if ext not in ALLOWED_EXT:
            flash("Tipe file tidak didukung. Gunakan .xls atau .xlsx", "danger")
            return redirect(request.url)

        output_name = request.form.get("output_name") or "QR PTK DARUL MUJAHIDIN"
        output_name = clean_filename(output_name)
        compress = request.form.get("compress")

        req_id = uuid.uuid4().hex
        req_folder = BASE_OUTPUT / req_id
        req_folder.mkdir(parents=True, exist_ok=True)
        upload_path = req_folder / filename
        f.save(upload_path)

        try:
            df = read_excel_detect(upload_path)
        except Exception as e:
            flash(f"Gagal membaca file Excel: {e}", "danger")
            return redirect(request.url)

        df.columns = df.columns.str.strip()
        nama_col = "Nama Peserta"
        kode_col = "QR-Code"
        kelas_col = "Kelas"

        if not all(col in df.columns for col in [nama_col, kode_col, kelas_col]):
            flash(f"Kolom '{nama_col}', '{kode_col}', atau '{kelas_col}' tidak ditemukan.", "danger")
            return redirect(request.url)

        out_root = req_folder / output_name
        out_root.mkdir(exist_ok=True)

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

        archive_path = None
        if compress == "zip":
            archive_path = req_folder / f"{output_name}.zip"
            zip_folder(out_root, archive_path)
        elif compress == "rar":
            rar_created = rar_with_winrar(out_root, req_folder / f"{output_name}.rar")
            if rar_created:
                archive_path = req_folder / f"{output_name}.rar"
            else:
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

        if archive_path is None:
            archive_path = req_folder / f"{output_name}.zip"
            zip_folder(out_root, archive_path)
            flash("Pembuatan RAR gagal atau tidak tersedia, dibuat ZIP sebagai fallback.", "warning")

        return send_file(str(archive_path), as_attachment=True)

    # GET request
    return render_template_string(HTML_TEMPLATE, rar_available=RAR_AVAILABLE)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
