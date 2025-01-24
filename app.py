from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os
import PyPDF2

app = Flask(__name__)

# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
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
    # Get uploaded Excel file and directory path
    excel_file = request.files['excelFile']
    file_directory = request.form['fileDirectory']

    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Ensure the Excel file has a column named "Filename" (or adjust as needed)
    if "Filename" not in df.columns:
        return jsonify({"error": "The Excel file must contain a 'Filename' column."}), 400

    # Scan files and extract TOTAL CYCLE TIME
    results = []
    for filename in df["Filename"]:
        file_path = os.path.join(file_directory, filename)
        if os.path.exists(file_path):
            if file_path.endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            else:
                with open(file_path, 'r') as file:
                    text = file.read()
            cycle_time = extract_cycle_time(text)
            results.append({"Filename": filename, "TOTAL CYCLE TIME": cycle_time if cycle_time else "Not found"})
        else:
            results.append({"Filename": filename, "TOTAL CYCLE TIME": "File not found"})

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
