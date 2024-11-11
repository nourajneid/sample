import urllib.parse
from flask import Flask, request, send_file, jsonify, render_template_string
import os
import time
import traceback
import tempfile
import shutil
from threading import Thread
from waitress import serve
from ppt_workbook_update import analyze_excel_markers, modify_embedded_excel_in_pptx
from refreshCharts import refreshCharts, vba_code  # only works on Windows and will not work with Linux or Mac
from openpyxl import load_workbook

app = Flask(__name__)

# Directory where temporary files will be stored
TEMP_ROOT = os.path.join(os.getcwd(), 'temp')
LOCK_FILE = "powerpoint_process.lock"

# HTML content rendered directly via Flask (for testing only, prod is using bodhi vue)
index_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload PowerPoint and Excel Files</title>
</head>
<body>
    <h1>Upload PowerPoint and Excel Files</h1>
    <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data">
        <label for="ppt_file">PowerPoint File (.pptx, .pptm):</label><br>
        <input type="file" id="ppt_file" name="ppt_file" accept=".pptx,.pptm" required><br><br>
        
        <label for="excel_file">Excel File (.xlsx):</label><br>
        <input type="file" id="excel_file" name="excel_file" accept=".xlsx" required><br><br>

        <button type="submit">Upload and Process</button>
    </form>

    <script>
        document.getElementById('uploadForm').onsubmit = function(event) {
            event.preventDefault();

            const formData = new FormData();
            const pptFile = document.getElementById('ppt_file').files[0];
            const excelFile = document.getElementById('excel_file').files[0];

            formData.append('ppt_file', pptFile);
            formData.append('excel_file', excelFile);

            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error("Error processing files");
                }
                return response.json();
            })
            .then(data => {
                // Trigger download via download link
                const downloadLink = document.createElement('a');
                downloadLink.href = `${window.location.origin}/download?filename=${encodeURIComponent(data.filename)}`;
                downloadLink.download = data.filename;
                downloadLink.click();
            })
            .catch(error => {
                alert("Failed to process files: " + error.message);
            });
        };
    </script>
</body>
</html>
"""

def is_locked():
    """Check if lock file exists and ensure it is not older than 2 minutes."""
    if os.path.exists(LOCK_FILE):
        lock_time = os.path.getmtime(LOCK_FILE)
        if time.time() - lock_time > 120:
            # Lock is older than 2 minutes, update timestamp
            print("Lock is older than 2 minutes. Updating timestamp.")
            create_lock()
        return True
    return False

def create_lock():
    """Create lock file with the current timestamp."""
    with open(LOCK_FILE, "w") as lock_file:
        lock_file.write("This file is used to lock the PowerPoint process.")

def remove_lock():
    """Remove lock file."""
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

def clean_old_temp_dirs():
    """
    Routine to clean up temporary directories older than 5 minute.
    """
    while True:
        now = time.time()
        for temp_dir in os.listdir(TEMP_ROOT):
            temp_dir_path = os.path.join(TEMP_ROOT, temp_dir)
            # Check if it is a directory and get its age
            if os.path.isdir(temp_dir_path):
                dir_age = now - os.path.getmtime(temp_dir_path)
                if dir_age > 5 * 60:  # Older than 5 minutes
                    try:
                        shutil.rmtree(temp_dir_path)
                        print(f"Removed old temp directory: {temp_dir_path}")
                    except Exception as e:
                        print(f"Error removing {temp_dir_path}: {e}")
        time.sleep(30)  # Check every 30 seconds

@app.route('/')
def index():
    """Render the HTML form."""
    return render_template_string(index_html)
@app.route('/upload', methods=['POST'])
def upload_and_process():
    print("first in route")
    """Handle the file upload and process the files."""
    skip_macro = request.form.get('skip_macro') == 'true'  # Retrieve the skip_macro value
    if not skip_macro:
        if is_locked():
            print("file was lock")
            return jsonify({"error": "Another process is currently using PowerPoint. Please try again later or use the skip macro option."}), 400
        # Create lock to ensure exclusive access if macro is not skipped
        create_lock()

    if 'ppt_file' not in request.files or 'excel_file' not in request.files:
        return jsonify({"error": "Please upload both PowerPoint and Excel files"}), 400

    ppt_file = request.files['ppt_file']
    excel_file = request.files['excel_file']
    skip_macro = request.form.get('skip_macro') == 'true'  # Retrieve the skip_macro value
    print("skip_macro is ", skip_macro)
    # Create temporary directories to store uploaded files and output
    temp_dir = tempfile.mkdtemp(prefix='powerpoint_', dir=TEMP_ROOT)
    try:
        # Save uploaded PowerPoint and Excel files to the temporary directory
        ppt_file_path = os.path.join(temp_dir, ppt_file.filename)
        excel_file_path = os.path.join(temp_dir, excel_file.filename)
        ppt_file.save(ppt_file_path)
        excel_file.save(excel_file_path)

        # Load the Excel workbook once
        workbook = load_workbook(excel_file_path, data_only=True, read_only=True)

        # Output file path for the modified PowerPoint
        updated_filename = ppt_file.filename.replace('.ppt', '_updated.ppt') # this is to handle both pptx and pptm
        output_ppt_file = os.path.join(temp_dir, updated_filename)
        
        # Configuration dictionary for modifying embedded Excel files
        config_modify = {
            "use_filesystem": False,
            'ppt_file_path': ppt_file_path,
            'workbook': workbook,
            'mapping': analyze_excel_markers(workbook)
        }

        # Modify embedded Excel files in the PowerPoint based on the marker mapping
        presentation = modify_embedded_excel_in_pptx(config_modify)

        # Save the modified presentation to the output file path
        with open(output_ppt_file, 'wb') as output_file:
            presentation.save(output_file)

        workbook.close() # Explicitly delete the workbook to release the file handle

        if not skip_macro:
            try:
                print("before refreshing")
                refreshCharts(output_ppt_file)  # window only
            except Exception as e:
                print(f"refreshCharts failed: {e}")
                # Add a header indicating refreshCharts failure
                error_message = f"The chart refresh is incomplete. The workbook for the chart has been updated, but the chart cache could not be recomputed. You can manually refresh the chart clicking on 'I will run the macro myself' and run the macro by yourself" # you could add {urllib.parse.quote(vba_code)}
                print("error_message", error_message)
            else:
                error_message = None
        else:
            print("Skipping macro as requested.")
            error_message = None

        # Send the relative path to the client for download
        relative_file_path = os.path.relpath(output_ppt_file, TEMP_ROOT)
        return jsonify({"filename": relative_file_path})
    
    except Exception as e:
        error_trace = traceback.format_exc()
        print("error ", error_trace)
        # Optionally, you can log the error_trace somewhere (e.g., to a file or a logging service)
        return jsonify({"error": str(e), "stack": error_trace}), 500
    
    finally:
        # Remove lock to allow other processes if macro was not skipped
        if not skip_macro:
            remove_lock()
            # Don't remove temp_dir immediately to allow download
            print("end")

@app.route('/download', methods=['GET'])
def download_file():
    print("in download")
    """Provide the generated file for download."""
    filename = request.args.get('filename')
    if not filename:
        return jsonify({"error": "Filename is required"}), 400
    
    if not filename.startswith('powerpoint_'):
        print("filename not matching pattern", filename)
        return jsonify({"error": "Invalid filename"}), 400
    # Path to the file in the temp directory
    file_path = os.path.join(TEMP_ROOT, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    return send_file(file_path, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    # Ensure TEMP_ROOT exists
    os.makedirs(TEMP_ROOT, exist_ok=True)

    # Start the cleanup thread to remove old temp directories
    cleanup_thread = Thread(target=clean_old_temp_dirs, daemon=True)
    cleanup_thread.start()

    # Run Flask using Waitress server on port 8000
    serve(app, host='0.0.0.0', port=8000)
