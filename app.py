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

@app.route('/process', methods=['POST'])
def process_files():
    if 'excel' not in request.files or 'files' not in request.files:
        return jsonify({"error": "Please upload both Excel and document files."}), 400

    excel_file = request.files['excel']
    uploaded_files = request.files.getlist('files')

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

    results = []

    for file in uploaded_files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        try:
            if file.filename.lower().endswith('.pdf'):
                pdf_reader = PdfReader(file_path)
                text = "".join(page.extract_text() for page in pdf_reader.pages)
            else:
                with open(file_path, 'r') as f:
                    text = f.read()

            cycle_time = extract_cycle_time(text)
            results.append({'fileName': file.filename, 'cycleTime': cycle_time})

        except Exception as e:
            return jsonify({"error": f"Error processing {file.filename}: {str(e)}"}), 400

    for result in results:
        # Normalize file name by removing extension for matching
        normalizedFileName = os.path.splitext(result['fileName'])[0].lower()
        rowIndex = df[df.iloc[:, 0].astype(str).str.lower().str.replace('.pdf', '') == normalizedFileName].index

        if not rowIndex.empty:
            df.loc[rowIndex[0], 'File Name'] = result['fileName']
            df.loc[rowIndex[0], 'TOTAL CYCLE TIME'] = result['cycleTime'] or 'No instances of "TOTAL CYCLE TIME" found.'
        else:
            # Add new row if no match found
            new_row = pd.DataFrame({
                'File Name': [result['fileName']],
                'TOTAL CYCLE TIME': [result['cycleTime'] or 'No instances of "TOTAL CYCLE TIME" found.'],
                df.columns[0]: [normalizedFileName]  # Use the first column for the file name without extension
            })
            df = pd.concat([df, new_row], ignore_index=True)

    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Clean up temp files
    for temp_file in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_file)
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
