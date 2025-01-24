from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# Function to calculate TOTAL CYCLE TIME (customize this based on your file structure)
def calculate_total_cycle_time(file_path):
    """
    Example: Calculate TOTAL CYCLE TIME based on timestamps in the file.
    Replace this logic with your actual calculation.
    """
    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    # Example: Extract timestamps and calculate the difference
    timestamps = []
    for line in lines:
        if "Timestamp:" in line:  # Replace with your actual timestamp identifier
            timestamp_str = line.split("Timestamp:")[1].strip()
            timestamp = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")  # Adjust format as needed
            timestamps.append(timestamp)
    
    if len(timestamps) >= 2:
        total_cycle_time = (timestamps[-1] - timestamps[0]).total_seconds()  # Calculate time difference in seconds
        return total_cycle_time
    else:
        return 0  # Return 0 if there are not enough timestamps

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

    # Scan files and calculate TOTAL CYCLE TIME
    results = []
    for filename in df["Filename"]:
        file_path = os.path.join(file_directory, filename)
        if os.path.exists(file_path):
            total_cycle_time = calculate_total_cycle_time(file_path)
            results.append({"Filename": filename, "Total Cycle Time": total_cycle_time})
        else:
            results.append({"Filename": filename, "Total Cycle Time": "File not found"})

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
