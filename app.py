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

# Function to extract the TOTAL CYCLE TIME line from text
def extract_cycle_time_line(text):
    lines = text.split('\n')
    for line in lines:
        if "TOTAL CYCLE TIME" in line:
            return line.strip()  # Return the entire line containing "TOTAL CYCLE TIME"
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

    # Scan uploaded files and extract the TOTAL CYCLE TIME line
    for file in files_to_scan:
        filename = file.filename
        if filename in df["Filename"].values:
            if filename.endswith('.pdf'):
                text = extract_text_from_pdf(file)
                cycle_time_line = extract_cycle_time_line(text)
                results.append({"Filename": filename, "TOTAL CYCLE TIME": cycle_time_line if cycle_time_line else "Not found"})

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
