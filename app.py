from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
from openpyxl import load_workbook
from Main import read_tables_from_sheet, remove_none_values, remove_total_row, Highest_Payment, Lowest_Payment
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)    
    if file and file.filename.endswith('.xlsx'):
        unique_filename = str(uuid.uuid4()) + '.xlsx'
        file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
        file.save(file_path)
        process_file(file_path)
        return 'File processed successfully!'
    return 'Invalid file format. Please upload an .xlsx file.'

def process_file(filename):
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.active
    tables = read_tables_from_sheet(sheet)
    
    for i, table in enumerate(tables):
        table = remove_none_values(table)
        table = remove_total_row(table)
        if len(table) > 1:
            df = pd.DataFrame(table[1:], columns=table[0])
            print(f"Table {i+1} DataFrame:")
            print(df)
            Highest_Payment(table)
            Lowest_Payment(table)

if __name__ == '__main__':
    app.run(debug=True)