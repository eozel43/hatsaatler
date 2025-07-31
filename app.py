from flask import Flask, request, send_file, render_template
import pandas as pd
from datetime import datetime, timedelta, time
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def generate_schedule(excel_path, output_path, year):
    df_saatler = pd.read_excel(excel_path)
    unique_hat_no = df_saatler["Hat No"].unique()
    unique_ay = df_saatler["Ay"].unique()
    if len(unique_hat_no) != 1 or len(unique_ay) != 1:
        raise ValueError("Excel dosyası birden fazla 'Hat No' veya 'Ay' içeriyor. Lütfen dosyayı kontrol edin.")
    hat_no = unique_hat_no[0]
    month = unique_ay[0]
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    date_list = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
    new_data = []
    for date in date_list:
        if date.weekday() < 5:
            gun_tipi = "Hafta İçi"
        elif date.weekday() == 5:
            gun_tipi = "Cumartesi"
        else:
            gun_tipi = "Pazar"
        for yon in ["G", "D"]:
            df_filtered = df_saatler[
                (df_saatler["Hat No"] == hat_no) & 
                (df_saatler["Ay"] == month) & 
                (df_saatler["Gün Tipi"] == gun_tipi) &
                (df_saatler["Yön"] == yon)
            ]
            if not df_filtered.empty:
                saatler = df_filtered.iloc[:, 4:].values.flatten()
                saatler = [s.time() if isinstance(s, datetime) else s for s in saatler if pd.notna(s)]
            else:
                saatler = []
            max_saat_sutun = max(len(df_saatler.columns[4:]), len(saatler))
            new_data.append([hat_no, date.strftime('%Y-%m-%d'), yon] + saatler + [""] * (max_saat_sutun - len(saatler)))
    column_names = ["Hat No", "Tarih", "Yön"] + [f"Saat{i+1}" for i in range(len(new_data[0]) - 3)]
    df_new = pd.DataFrame(new_data, columns=column_names)
    df_new.to_excel(output_path, index=False)

import zipfile
import io

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('file')
        year = int(request.form.get('year', datetime.now().year))
        
        if not uploaded_files or all(f.filename == '' for f in uploaded_files):
            return 'Dosya seçilmedi!', 400

        output_files = []
        uploaded_file_paths = []

        try:
            for file in uploaded_files:
                if file and file.filename.endswith('.xlsx'):
                    filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                    file.save(filename)
                    uploaded_file_paths.append(filename)

                    output_filename = f"{os.path.splitext(file.filename)[0]}_{year}.xlsx"
                    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
                    
                    generate_schedule(filename, output_path, year)
                    output_files.append(output_path)
                else:
                    return f"Geçersiz dosya formatı: {file.filename}. Sadece .xlsx dosyaları kabul edilir.", 400

            if not output_files:
                return "Hiçbir geçerli Excel dosyası işlenemedi.", 400

            # Create a zip file in memory
            memory_file = io.BytesIO()
            with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fpath in output_files:
                    zf.write(fpath, os.path.basename(fpath))
            memory_file.seek(0)

            # Clean up generated files and uploaded files
            for fpath in output_files:
                os.remove(fpath)
            for fpath in uploaded_file_paths:
                os.remove(fpath)

            return send_file(memory_file, as_attachment=True, download_name='hat_saatleri.zip', mimetype='application/zip')

        except Exception as e:
            # Clean up in case of an error
            for fpath in output_files:
                if os.path.exists(fpath):
                    os.remove(fpath)
            for fpath in uploaded_file_paths:
                if os.path.exists(fpath):
                    os.remove(fpath)
            return str(e), 400
            
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
