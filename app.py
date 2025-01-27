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

    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return jsonify({"error": f"Error reading Excel file: {str(e)}"}), 400

    if len(df.columns) < 2:
        return jsonify({"error": "Excel file must have at least 2 columns."}), 400

    if 'File Name' not in df.columns:
        df['File Name'] = ''
    
    if 'TOTAL CYCLE TIME' not in df.columns:
        df['TOTAL CYCLE TIME'] = ''

    resultsArray = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        pdf_file.save(pdf_path)

        try:
            pdf_reader = PdfReader(pdf_path)
            pdf_text = ""
            for page in pdf_reader.pages:
                pdf_text += page.extract_text()
            
            print(f"Text from {pdf_file.filename}: {pdf_text[:100]}...")  # Only print first 100 characters for brevity

            cycle_time = extract_cycle_time(pdf_text)
            resultsArray.append({'fileName': pdf_file.filename, 'cycleTime': cycle_time})

        except Exception as e:
            return jsonify({"error": f"Error processing PDF file {pdf_file.filename}: {str(e)}"}), 400

    excelData = df.values.tolist()

    for result in resultsArray:
        normalizedFileName = result['fileName'].lower().replace('.pdf', '')  # Adjust the extension as needed
        rowIndex = next((index for index, row in enumerate(excelData) if row[0] and row[0].lower().replace('.pdf', '') == normalizedFileName), None)

        if rowIndex is not None:
            print(f"Updating row for {result['fileName']}")
            excelData[rowIndex][1] = result['cycleTime'] or 'No instances of "TOTAL CYCLE TIME" found.'
        else:
            print(f"No match for {result['fileName']}, adding new row.")
            newRow = [result['fileName'], result['cycleTime'] or 'No instances of "TOTAL CYCLE TIME" found.']
            excelData.append(newRow)

    # Convert back to DataFrame for saving
    df = pd.DataFrame(excelData, columns=df.columns)

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

    return send_file(updated_excel_path, as_attachment=True, download_name='updated_excel.xlsx')

def extract_cycle_time(text):
    regex = r"(?i)total\s*cycle\s*time\s*:\s*(\d+\s*(?:hours?|hr)\s*,\s*\d+\s*(?:minutes?|min)\s*,\s*\d+\s*(?:seconds?|sec))"
    match = re.search(regex, text)
    return match.group(1).strip() if match else None

if __name__ == '__main__':
    app.run(debug=True)
