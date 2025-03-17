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

@app.route('/submit_data', methods=['POST'])
def submit_data():
    data1 = request.form['data1']
    data2 = request.form['data2']
    
    # Process the input data
    results = process_input_data(data1, data2)
    
    return render_template('results.html', results=results)

def process_input_data(data1, data2):
    # Example processing logic
    results = {
        'data1': data1,
        'data2': data2,
        'summary': f"Received data1: {data1} and data2: {data2}"
    }
    return results

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)