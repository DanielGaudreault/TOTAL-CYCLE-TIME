# PDF Data Extractor to Excel Updater

A simple web application to extract project names and cycle times from PDFs and update an Excel sheet.

## Live Demo

Access the application here: [PDF Data Extractor to Excel Updater](https://danielgaudreault.github.io/TOTAL-CYCLE-TIME/)

## Prerequisites

- **Node.js** installed for running an HTTP server (optional but recommended).
- A modern web browser (Chrome, Firefox, etc.).

## Step-by-Step Guide

### Step 1: Setup Your Project

1. **Create Project Directory:**
   - Make a new directory for your project:
     ```bash
     mkdir pdf-excel-updater
     cd pdf-excel-updater
     ```

2. **Project Structure:**
   - Ensure your project has the following structure:
     ```
     pdf-excel-updater/
     ├── index.html
     ├── script.js
     ├── styles.css
     ```

### Step 2: Install Dependencies

For this project, we'll use a simple HTTP server from Node.js for serving files:

1. **Install Node.js:**
   - If you don't have Node.js installed, download and install it from [nodejs.org](https://nodejs.org/en/download/).

2. **Install http-server (optional, but recommended):**
   - Open your terminal/command prompt in the project directory and run:
     ```bash
     npm install -g http-server
     ```

### Step 3: Prepare Your Files

- **HTML (`index.html`):** Contains your form for file uploads and the table to display results.
- **JavaScript (`script.js`):** Includes all the logic for PDF parsing and Excel updating.
- **CSS (`styles.css`):** For basic styling (if you have one).

### Step 4: Serve Your Webpage

#### Option 1: Serve Locally with `http-server`

1. **Start the Server:**
   - Navigate to your project directory and run:
     ```bash
     http-server
     ```

   This will start your server on `localhost:8080` or another port if 8080 is in use.

2. **Open in Browser:**
   - Navigate to `http://localhost:8080/` in your web browser.

#### Option 2: Python's Simple HTTP Server (For Python users)

1. **Start Python Server:**
   - If you have Python installed, you can use:
     ```bash
     python -m http.server
     ```

   This defaults to `localhost:8000`.

2. **Open in Browser:**
   - Navigate to `https://danielgaudreault.github.io/TOTAL-CYCLE-TIME/` in your browser.

### Step 5: Using the Application

1. **Upload PDFs:**
   - Use the file input to select PDF files, then click `Process Files`.

2. **View Results:**
   - The results of your PDF parsing should appear in the table.

3. **Update Excel:**
   - Select an Excel file with the `Update to Excel` section, then click `Update to Excel`.

4. **Download Updated Excel:**
   - If the script runs successfully, you'll get a prompt to download the updated Excel file.

### Troubleshooting

- **File Permissions:** Ensure your browser has permission to access local files if you're opening the page directly from the file system.
- **Browser Security:** Some operations might be blocked if not served from a server. Use `http-server` or Python's HTTP server.
- **Console Logs:** Check your browser's console for any error messages or logs to debug issues.

### Notes

- This application assumes you have the necessary CDN versions of `pdf.js` and `xlsx.js` included in your HTML for PDF parsing and Excel manipulation respectively.
- If you encounter persistent issues, consider moving Excel manipulation to a server-side script.

---

Feel free to modify this guide according to your specific setup or additional requirements.
