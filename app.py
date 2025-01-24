from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import PyPDF2
import os
from io import BytesIO

app = Flask(__name__)

# Function to extract text from a PDF file
def extract_text_from_pdf(file):
    reader = PyPDF2.PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

# Function to extract "TOTAL CYCLE TIME" from text
def extract_cycle_time(text):
    lines = text.split('\n')
    for line in lines:
        if "TOTAL CYCLE TIME" in line:
            # Use regex to extract the time part (e.g., "0 HOURS, 4 MINUTES, 16 SECONDS")
            import re
            match = re.search(r"(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)", line, re.IGNORECASE)
            return match.group(0) if match else None
    return None

@app.route('/process', methods=['POST'])
def process_files():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded."}), 400

    files = request.files.getlist('files')
    results = []

    for file in files:
        if file.filename.endswith('.pdf'):
            text = extract_text_from_pdf(file)
        else:
            text = file.read().decode('utf-8')

        cycle_time = extract_cycle_time(text)
        results.append({
            "Filename": file.filename,
            "TOTAL CYCLE TIME": cycle_time if cycle_time else "Not found"
        })

    # Save results to an Excel file
    df = pd.DataFrame(results)
    output_file = "results.xlsx"
    df.to_excel(output_file, index=False)

    return jsonify({
        "results": results,
        "outputFile": output_file
    })

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_from_directory('.', filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
