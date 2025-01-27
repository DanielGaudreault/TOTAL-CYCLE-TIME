from flask import Flask, request, render_template, send_file, jsonify
import pandas as pd
from PyPDF2 import PdfReader
import os

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
        return jsonify({"error": "Please upload both Excel and PDF files."}), 400

    excel_file = request.files['excel']
    pdf_files = request.files.getlist('pdf')

    # Save Excel file temporarily
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    # Read Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return jsonify({"error": f"Error reading Excel file: {e}"}), 400

    # Ensure the Excel file has the required columns
    if 'File Name' not in df.columns or 'Part Number' not in df.columns:
        return jsonify({"error": "Excel file must contain 'File Name' and 'Part Number' columns."}), 400

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
            return jsonify({"error": f"Error reading PDF file: {e}"}), 400

        # Extract TOTAL CYCLE TIME from PDF text
        cycle_time = extract_cycle_time(pdf_text)
        if cycle_time:
            # Match the PDF file name or part number with the Excel row
            pdf_name = pdf_file.filename
            part_number = extract_part_number(pdf_text)  # Extract part number from PDF text

            # Update the matching row in the Excel file
            if pdf_name in df['File Name'].values:
                df.loc[df['File Name'] == pdf_name, 'TOTAL CYCLE TIME'] = cycle_time
            elif part_number and part_number in df['Part Number'].values:
                df.loc[df['Part Number'] == part_number, 'TOTAL CYCLE TIME'] = cycle_time

    # Save updated Excel file
    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Provide download link
    return send_file(updated_excel_path, as_attachment=True)

def extract_cycle_time(text):
    # Use regex to find "TOTAL CYCLE TIME" and extract the time
    import re
    regex = r"TOTAL CYCLE TIME\s*:\s*(\d+\s*HOURS?,\s*\d+\s*MINUTES?,\s*\d+\s*SECONDS?)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1) if match else None

def extract_part_number(text):
    # Use regex to find the part number (adjust the regex as needed)
    import re
    regex = r"Part Number\s*:\s*(\w+)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1) if match else None

if __name__ == '__main__':
    app.run(debug=True)
