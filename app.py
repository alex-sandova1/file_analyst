from flask import Flask, request, render_template, redirect, url_for, send_file
import pandas as pd
from openpyxl import load_workbook
from analysis import *
from visualize import *
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(PROCESSED_FOLDER):
    os.makedirs(PROCESSED_FOLDER)

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
        results, processed_file_path, pie_chart_path = process_file(file_path)
        return render_template('results.html', results=results, processed_file=processed_file_path, pie_chart=pie_chart_path)
    return 'Invalid file format. Please upload an .xlsx file.'

def process_file(filename):
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.active
    tables = read_tables_from_sheet(sheet)
    
    results = []
    pie_chart_path = None
    for i, table in enumerate(tables):
        table = remove_none_values(table)
        table = remove_total_row(table)
        if len(table) > 1:
            df = pd.DataFrame(table[1:], columns=table[0])
            highest_payment = Highest_Payment(table)
            lowest_payment = Lowest_Payment(table)
            results.append({
                'table_number': i + 1,
                'dataframe': df.to_html(),
                'highest_payment': highest_payment,
                'lowest_payment': lowest_payment
            })
            # Save processed data to a new Excel file
            processed_filename = f'processed_{uuid.uuid4()}.xlsx'
            processed_file_path = os.path.join(PROCESSED_FOLDER, processed_filename)
            with pd.ExcelWriter(processed_file_path) as writer:
                df.to_excel(writer, index=False, sheet_name=f'Table_{i+1}')
            
            # Generate pie chart
            pie_chart_filename = f'pie_chart_{uuid.uuid4()}.png'
            pie_chart_path = os.path.join(PROCESSED_FOLDER, pie_chart_filename)
            pie_graph(table, table[0][1], pie_chart_path)
    return results, processed_file_path, pie_chart_path

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)