from flask import Flask, request, render_template, send_file, jsonify
import pandas as pd
from PyPDF2 import PdfReader
import os
import re
import shutil

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'excel' not in request.files or 'pdf' not in request.files:
        return jsonify({"error": "Please upload both Excel and PDF files."}), 400

    excel_file = request.files['excel']
    pdf_files = request.files.getlist('pdf')

    # Save Excel file temporarily
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return jsonify({"error": f"Error reading Excel file: {str(e)}"}), 400

    # Ensure the Excel file has at least 2 columns (column B is index 1)
    if len(df.columns) < 2:
        return jsonify({"error": "Excel file must have at least 2 columns."}), 400

    # Add "File Name" and "TOTAL CYCLE TIME" columns if they don't exist
    if 'File Name' not in df.columns:
        df['File Name'] = ''
    
    if 'TOTAL CYCLE TIME' not in df.columns:
        df['TOTAL CYCLE TIME'] = ''

    for pdf_file in pdf_files:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        pdf_file.save(pdf_path)

        try:
            pdf_reader = PdfReader(pdf_path)
            pdf_text = ""
            for page in pdf_reader.pages:
                pdf_text += page.extract_text()
        except Exception as e:
            return jsonify({"error": f"Error reading PDF file {pdf_file.filename}: {str(e)}"}), 400

        # Extract Item Number from the file name
        item_number = extract_item_number_from_file_name(pdf_file.filename)

        # Extract TOTAL CYCLE TIME from PDF text
        cycle_time = extract_cycle_time(pdf_text)

        if item_number:
            matching_row = df.iloc[:, 1].astype(str).str.strip() == str(item_number).strip()
            if matching_row.any():
                df.loc[matching_row, 'File Name'] = pdf_file.filename
                df.loc[matching_row, 'TOTAL CYCLE TIME'] = cycle_time or 'No instances of "TOTAL CYCLE TIME" found.'
                print(f"Updated row with Item No. {item_number}: File Name = {pdf_file.filename}, TOTAL CYCLE TIME = {cycle_time}")
            else:
                # Add new row if no match found
                new_row = pd.DataFrame({'File Name': [pdf_file.filename], 
                                        'TOTAL CYCLE TIME': [cycle_time or 'No instances of "TOTAL CYCLE TIME" found.'], 
                                        df.columns[1]: [item_number]})
                df = pd.concat([df, new_row], ignore_index=True)
                print(f"Added new row for PDF file: {pdf_file.filename}")
        else:
            print(f"Could not extract item number from PDF file name: {pdf_file.filename}")

    # Save updated Excel file
    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Clean up temp files
    for file in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

    # Provide download link
    return send_file(updated_excel_path, as_attachment=True, download_name='updated_excel.xlsx')

def extract_item_number_from_file_name(file_name):
    # Use regex to extract the item number from the file name
    regex = r"Item\s*(\d+)"
    match = re.search(regex, file_name, re.IGNORECASE)
    return match.group(1).strip() if match else None

def extract_cycle_time(text):
    # Use regex to find "TOTAL CYCLE TIME" and extract the time
    regex = r"TOTAL CYCLE TIME\s*:\s*(\d+\s*HOURS?,\s*\d+\s*MINUTES?,\s*\d+\s*SECONDS?)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1).strip() if match else None

if __name__ == '__main__':
    app.run(debug=True)
