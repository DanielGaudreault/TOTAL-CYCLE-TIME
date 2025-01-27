from flask import Flask, request, render_template, send_file
import pandas as pd
from PyPDF2 import PdfReader
import os
import re

app = Flask(__name__)

# Ensure the uploads directory exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check if files are uploaded
    if 'excel' not in request.files or 'pdf' not in request.files:
        return "Please upload both Excel and PDF files."

    excel_file = request.files['excel']
    pdf_files = request.files.getlist('pdf')

    # Save Excel file temporarily
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    # Read Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    # Ensure the Excel file has at least 2 columns (column B is index 1)
    if len(df.columns) < 2:
        return "Excel file must have at least 2 columns."

    # Add "File Name" and "TOTAL CYCLE TIME" columns if they don't exist
    if 'File Name' not in df.columns:
        df['File Name'] = ''

    if 'TOTAL CYCLE TIME' not in df.columns:
        df['TOTAL CYCLE TIME'] = ''

    # Process each PDF file
    for pdf_file in pdf_files:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        pdf_file.save(pdf_path)

        # Read PDF file
        try:
            pdf_reader = PdfReader(pdf_path)
            pdf_text = ""
            for page in pdf_reader.pages:
                pdf_text += page.extract_text()
        except Exception as e:
            return f"Error reading PDF file: {e}"

        # Extract Item Number from the file name
        item_number = extract_item_number_from_file_name(pdf_file.filename)

        # Extract TOTAL CYCLE TIME from PDF text
        cycle_time = extract_cycle_time(pdf_text)

        if item_number:
            # Match item number with column B (index 1) in Excel
            matching_row = df.iloc[:, 1] == item_number

            if matching_row.any():
                # Update the "File Name" and "TOTAL CYCLE TIME" columns for the matching row
                df.loc[matching_row, 'File Name'] = pdf_file.filename
                df.loc[matching_row, 'TOTAL CYCLE TIME'] = cycle_time or ''
            else:
                print(f"No matching item number found for PDF file: {pdf_file.filename}")
        else:
            print(f"Could not extract item number from PDF file name: {pdf_file.filename}")

    # Save updated Excel file
    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Provide download link
    return send_file(updated_excel_path, as_attachment=True)

def extract_item_number_from_file_name(file_name):
    # Use regex to extract the item number from the file name
    # Example: "Item12345.pdf" -> "12345"
    regex = r"Item\s*(\d+)"
    match = re.search(regex, file_name, re.IGNORECASE)
    return match.group(1).strip() if match else None

def extract_cycle_time(text):
    # Use regex to find "TOTAL CYCLE TIME" and extract the time
    regex = r"TOTAL CYCLE TIME\s*:\s*(\d+\s*HOURS?,\s*\d+\s*MINUTES?,\s*\d+\s*SECONDS?)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1) if match else None

if __name__ == '__main__':
    app.run(debug=True)
