# app.py

import os
import logging
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename
from main import main as process_cv  # Import the main processing function
from file_handler import validate_file
from formatter import format_data
from document_generator import create_document
from config import OPENAI_API_KEY

# Load environment variables
load_dotenv()

# Configuration
UPLOAD_FOLDER = 'Documents/To_Process'
PROCESSED_FOLDER = 'Documents/Processed'
ALLOWED_EXTENSIONS = {'pdf', 'docx'}

# Initialize Flask app
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.secret_key = os.urandom(24)  # For flash messages

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def allowed_file(filename):
    """Check if the file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    """Render the homepage and handle file uploads."""
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'cv_file' not in request.files:
            flash('No file part in the request.')
            return redirect(request.url)
        
        file = request.files['cv_file']
        
        # If user does not select file, browser may submit an empty part
        if file.filename == '':
            flash('No file selected for uploading.')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            try:
                # Ensure upload directory exists
                os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                
                # Save the uploaded file
                file.save(file_path)
                logging.info(f"File {filename} uploaded successfully.")
                
                # Validate the file
                validate_file(file_path)
                logging.info("File validation completed.")
                
                # Process the CV and get the output path
                output_path = process_cv(file_path)
                logging.info("CV processing completed.")
                
                # Check if the output file exists
                if os.path.exists(output_path):
                    output_filename = os.path.basename(output_path)
                    logging.info(f"Processed document {output_filename} is ready.")
                    return redirect(url_for('download_file', filename=output_filename))
                else:
                    flash('Processing failed. Output file not found.')
                    return redirect(request.url)
            
            except Exception as e:
                logging.error(f"Error processing file: {e}")
                flash(f"An error occurred while processing the file: {e}")
                return redirect(request.url)
        else:
            flash('Allowed file types are PDF and DOCX.')
            return redirect(request.url)
    
    return render_template('index.html')

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """Provide the processed CV for download."""
    # Ensure the filename is secure
    filename = secure_filename(filename)
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename, as_attachment=True)

@app.route('/results', methods=['GET'])
def results():
    """Display the results of CV processing."""
    # Implement this route if you want to show processed data on a webpage
    pass

if __name__ == '__main__':
    # Ensure processed directory exists
    os.makedirs(PROCESSED_FOLDER, exist_ok=True)
    
    # Run the Flask app
    app.run(debug=True, host='0.0.0.0', port=5001)
