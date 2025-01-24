from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import PyPDF2
from io import BytesIO

app = Flask(__name__)

# Function to extract text from a PDF file
def extract_text_from_pdf(file):
    reader = PyPDF2.PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

# Function to extract TOTAL CYCLE TIME from text
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
    # Get uploaded Excel file and files to scan
    excel_file = request.files['excelFile']
    files_to_scan = request.files.getlist('filesToScan')

    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Ensure the Excel file has a column named "Filename" (or adjust as needed)
    if "Filename" not in df.columns:
        return jsonify({"error": "The Excel file must contain a 'Filename' column."}), 400

    # Create a dictionary to store the results
    results = []

    # Scan uploaded files and extract TOTAL CYCLE TIME
    for file in files_to_scan:
        filename = file.filename
        if filename in df["Filename"].values:
            if filename.endswith('.pdf'):
                text = extract_text_from_pdf(file)
            else:
                text = file.read().decode('utf-8')
            cycle_time = extract_cycle_time(text)
            results.append({"Filename": filename, "TOTAL CYCLE TIME": cycle_time if cycle_time else "Not found"})

    # Update the Excel file with the results
    results_df = pd.DataFrame(results)
    df = df.merge(results_df, on="Filename", how="left")

    # Save the updated Excel file
    output_file = "updated_output.xlsx"
    df.to_excel(output_file, index=False)

    return jsonify({
        "message": "Processing complete!",
        "outputFile": output_file,
    })

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_from_directory('.', filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
