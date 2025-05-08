from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import json
import pandas as pd
from collections import defaultdict
import os
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)
app.secret_key = 'c821091fd2bfe9fa039086c8ed6349f0d2a5955141721cbe'  # Змініть на реальний секретний ключ

# Налаштування для завантаження файлів
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'jsonl'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Створюємо папку для завантажень, якщо її немає
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_jsonl_to_xlsx(input_file):
    sku_data = defaultdict(int)
    null_sku_count = 0
    total_sold = 0

    try:
        for line in input_file:
            line = line.decode('utf-8').strip()
            if not line:
                continue

            try:
                data = json.loads(line)
                sku = data.get("product_variant_sku")
                sold = int(data.get("net_items_sold", 0))

                if sku is None:
                    null_sku_count += sold
                else:
                    sku_data[sku] += sold
                total_sold += sold
            except (json.JSONDecodeError, ValueError) as e:
                print(f"Помилка при обробці рядка: {line}. Помилка: {e}")
                continue

        # Створюємо DataFrame
        df = pd.DataFrame(
            [('NULL_SKU', null_sku_count)] + list(sku_data.items()),
            columns=['Product Variant SKU', 'Net Items Sold']
        )

        # Додаємо рядок із загальною сумою
        total_row = pd.DataFrame([['TOTAL', total_sold]],
                                 columns=['Product Variant SKU', 'Net Items Sold'])
        df = pd.concat([df, total_row], ignore_index=True)

        # Створюємо excel файл у пам'яті
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sales Report')
            worksheet = writer.sheets['Sales Report']

            # Додаємо форматування
            header_format = writer.book.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1})

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Авто-розмір стовпців
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col) + 2)
                worksheet.set_column(i, i, max_len)

        output.seek(0)  # Цей рядок має бути ПОЗА циклом for та ПОЗА блоком with

        return output, total_sold, null_sku_count

    except Exception as e:
        print(f"Сталася помилка при обробці файлу: {e}")
        raise


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Перевіряємо, чи є файл у запиті
        if 'file' not in request.files:
            flash('Не вибрано файл')
            return redirect(request.url)

        file = request.files['file']

        # Якщо користувач не вибрав файл
        if file.filename == '':
            flash('Не вибрано файл')
            return redirect(request.url)

        # Якщо файл дозволеного формату
        if file and allowed_file(file.filename):
            try:
                # Обробляємо файл
                output, total_sold, null_sku_count = process_jsonl_to_xlsx(file)

                # Повертаємо результат користувачу
                return send_file(
                    output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name='sales_report.xlsx'
                )

            except Exception as e:
                flash(f'Помилка при обробці файлу: {str(e)}')
                return redirect(request.url)
        else:
            flash('Дозволені тільки файли з розширенням .jsonl')
            return redirect(request.url)

    return render_template('upload.html')


if __name__ == '__main__':
    app.run(debug=True)