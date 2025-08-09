# app.py
from flask import Flask, render_template, request, send_from_directory
import os
import pandas as pd
import qrcode
import zipfile
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='static', template_folder='templates')

# مسیر جدید خروجی‌ها در کنار app.py
BASE_EXPORT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exports")
os.makedirs(BASE_EXPORT, exist_ok=True)

ALLOWED_EXT = {'xlsx'}

def allowed_filename(filename):
    return '.' in filename and filename.rsplit('.',1)[1].lower() in ALLOWED_EXT

def base26_letters(n):
    letters = ''
    while True:
        n, r = divmod(n, 26)
        letters = chr(65 + r) + letters
        if n == 0:
            break
        n -= 1
    return letters

def get_month_code(month_num):
    month_map = {
        1: 'A', 2: 'Z', 3: 'Y', 4: 'C', 5: 'B', 6: 'D',
        7: 'E', 8: 'F', 9: 'G', 10: 'H', 11: 'J', 12: 'K'
    }
    return month_map.get(month_num, '?')

def generate_date_code(year, month):
    year_digit = str(year)[-1]
    month_code = get_month_code(month)
    return f"T{month_code}{month_code}{year_digit}"

def generate_serial_code(board_num, model_num, date_code, row_num):
    board_type = f"B{board_num:02d}"
    model = f"M{model_num:02d}"
    n = row_num - 101
    if n < 0:
        n = 0
    block = 1 + (n // 676)
    suffix = base26_letters(n % 676)
    code = f"KZ{block}{suffix}"
    serial = f"{board_type}-{model}-{date_code}-S{code}"
    return serial

def generate_qr_code_png(data, file_path):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="white", back_color="black")
    img.save(file_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    context = {
        "error": None,
        "result_ready": False,
        "serial_count": 0,
        "export_folder": None,
        "excel_filename": None,
        "csv_filename": None,
        "zip_filename": None,
        "timestamp": datetime.utcnow().strftime("%Y%m%d%H%M%S")
    }

    if request.method == 'POST':
        try:
            board_num = int(request.form.get('board_num', 0))
            model_num = int(request.form.get('model_num', 0))
            year = int(request.form.get('year', 0))
            month = int(request.form.get('month', 0))
            file = request.files.get('excel_file')

            if not file or file.filename == "":
                context['error'] = "فایل اکسل آپلود نشده."
                return render_template('index.html', **context)
            if not allowed_filename(file.filename):
                context['error'] = "فرمت فایل باید .xlsx باشد."
                return render_template('index.html', **context)
            if not (1 <= board_num <= 99):
                context['error'] = "نوع برد باید بین 1 تا 99 باشد."
                return render_template('index.html', **context)
            if not (1 <= model_num <= 99):
                context['error'] = "مدل دستگاه باید بین 1 تا 99 باشد."
                return render_template('index.html', **context)
            if not (1300 <= year <= 1499):
                context['error'] = "سال باید بین 1300 تا 1499 باشد."
                return render_template('index.html', **context)
            if not (1 <= month <= 12):
                context['error'] = "ماه باید بین 1 تا 12 باشد."
                return render_template('index.html', **context)

            df = pd.read_excel(file)
            if 'row' not in df.columns or 'imei' not in df.columns:
                context['error'] = "فایل اکسل باید شامل ستون‌های 'row' و 'imei' باشد."
                return render_template('index.html', **context)

            date_code = generate_date_code(year, month)
            timestamp = context['timestamp']
            export_name = f"output_{board_num}_{model_num}_{year}{month:02d}_{timestamp}"
            export_path = os.path.join(BASE_EXPORT, export_name)
            os.makedirs(export_path, exist_ok=True)

            output_rows = []
            for _, item in df.iterrows():
                row_num = int(item['row'])
                imei = str(item['imei']).strip()
                serial = generate_serial_code(board_num, model_num, date_code, row_num)

                serial_qr_filename = secure_filename(f"qr_serial_{row_num}.png")
                serial_qr_path = os.path.join(export_path, serial_qr_filename)
                generate_qr_code_png(serial, serial_qr_path)

                imei_qr_filename = secure_filename(f"qr_imei_{row_num}.png")
                imei_qr_path = os.path.join(export_path, imei_qr_filename)
                generate_qr_code_png(imei, imei_qr_path)

                # مسیر نسبی برای CSV
                rel_serial = os.path.relpath(serial_qr_path, os.path.dirname(__file__))
                rel_imei = os.path.relpath(imei_qr_path, os.path.dirname(__file__))

                output_rows.append({
                    "row": row_num,
                    "imei": imei,
                    "serial": serial,
                    "@serial_qr_path": rel_serial,
                    "@imei_qr_path": rel_imei
                })

            out_df = pd.DataFrame(output_rows).sort_values(by="row")
            excel_filename = f"{export_name}.xlsx"
            csv_filename = "output_fixed_relative.csv"  # ثابت
            out_df.to_excel(os.path.join(export_path, excel_filename), index=False)
            out_df.to_csv(os.path.join(export_path, csv_filename), index=False)

            zip_filename = f"{export_name}_qr.zip"
            zip_path = os.path.join(export_path, zip_filename)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for r in output_rows:
                    zipf.write(os.path.join(export_path, os.path.basename(r["@serial_qr_path"])),
                               arcname=os.path.basename(r["@serial_qr_path"]))
                    zipf.write(os.path.join(export_path, os.path.basename(r["@imei_qr_path"])),
                               arcname=os.path.basename(r["@imei_qr_path"]))

            context.update({
                "result_ready": True,
                "serial_count": len(output_rows),
                "export_folder": export_name,
                "excel_filename": excel_filename,
                "csv_filename": csv_filename,
                "zip_filename": zip_filename
            })

        except Exception as e:
            context['error'] = f"خطا: {str(e)}"

    return render_template('index.html', **context)

@app.route('/download/<export_folder>/<filename>')
def download_file(export_folder, filename):
    path = os.path.join(BASE_EXPORT, export_folder)
    return send_from_directory(path, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
