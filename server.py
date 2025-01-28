import os
import cgi
from http.server import SimpleHTTPRequestHandler, HTTPServer
import pandas as pd
from io import BytesIO
from urllib.parse import urlparse

UPLOAD_FOLDER = './uploads'

# Ensure the uploads folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

class MyHandler(SimpleHTTPRequestHandler):

    def do_POST(self):
        if self.path == '/upload':
            # Get the length of the uploaded content
            content_length = int(self.headers['Content-Length'])

            # Parse the form data
            form_data = self.rfile.read(content_length)
            _, pdict = cgi.parse_header(self.headers['Content-Type'])
            pdict['boundary'] = pdict['boundary'].encode('utf-8')
            fields = cgi.parse_multipart(BytesIO(form_data), pdict)

            # Save the uploaded Excel and PDF files
            excel_file = fields.get('excelFile')[0]
            pdf_files = fields.get('pdfFolder[]')

            excel_filename = self.save_file(excel_file, 'uploaded_excel.xlsx')
            pdf_filenames = [self.save_file(pdf, pdf.filename) for pdf in pdf_files]

            # Process the Excel file with PDF names
            self.process_excel(excel_filename, pdf_filenames)

            # Respond back to the frontend with a success message
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            response = {
                "message": "File uploaded and Excel updated successfully!"
            }
            self.wfile.write(bytes(str(response), 'utf-8'))

    def save_file(self, file_data, filename):
        """ Save uploaded files to the server """
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        with open(filepath, 'wb') as f:
            f.write(file_data)
        return filepath

    def process_excel(self, excel_filename, pdf_filenames):
        """ Process the Excel file and update columns based on PDF file names """
        df = pd.read_excel(excel_filename)

        # Example: update 'Column C' with the PDF filenames and 'Column D' with cycle time
        for i, pdf_filename in enumerate(pdf_filenames):
            setup_name = os.path.basename(pdf_filename)
            if len(df) > i:  # To avoid index out of bounds
                df.at[i, 'Column C'] = setup_name  # Add PDF filename to Column C
                df.at[i, 'Column D'] = "Total Cycle Time"  # Placeholder for cycle time

        # Save the updated Excel file
        updated_filename = 'updated_' + os.path.basename(excel_filename)
        updated_filepath = os.path.join(UPLOAD_FOLDER, updated_filename)
        df.to_excel(updated_filepath, index=False)

if __name__ == "__main__":
    # Start the server on port 8000
    server_address = ('', 8000)
    httpd = HTTPServer(server_address, MyHandler)
    print("Serving on port 8000...")
    httpd.serve_forever()
