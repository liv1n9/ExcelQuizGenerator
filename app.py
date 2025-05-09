import os
import logging
import tempfile
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from werkzeug.utils import secure_filename
from utils.excel_processor import validate_excel_file, get_random_questions
from utils.document_generator import generate_zip_files

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Create Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "default_secret_key")

# Configure upload settings
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
TEMP_FOLDER = tempfile.gettempdir()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check if the post request has the file part
        if 'excelFile' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        
        file = request.files['excelFile']
        
        # If user does not select file
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Check if it's an allowed file type
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx or .xls)'}), 400
        
        # Get form data
        num_questions = request.form.get('numQuestions')
        num_versions = request.form.get('numVersions')
        
        # Validate form data
        if not num_questions or not num_versions:
            return jsonify({'error': 'Please provide number of questions and versions'}), 400
        
        try:
            num_questions = int(num_questions)
            num_versions = int(num_versions)
        except ValueError:
            return jsonify({'error': 'Number of questions and versions must be integers'}), 400
        
        if num_questions <= 0 or num_versions <= 0:
            return jsonify({'error': 'Number of questions and versions must be positive'}), 400
        
        # Save the file to a temporary location
        filename = secure_filename(file.filename)
        file_path = os.path.join(TEMP_FOLDER, filename)
        file.save(file_path)
        
        # Validate Excel file format
        validation_result = validate_excel_file(file_path)
        if 'error' in validation_result:
            os.remove(file_path)  # Remove temporary file
            return jsonify(validation_result), 400
        
        # Generate the ZIP files
        questions_df = pd.read_excel(file_path)
        
        # Make sure we have enough questions
        total_questions = len(questions_df)
        if total_questions < num_questions:
            os.remove(file_path)  # Remove temporary file
            return jsonify({'error': f'Excel file contains only {total_questions} questions, but {num_questions} were requested'}), 400
        
        zip_files = generate_zip_files(questions_df, num_questions, num_versions)
        
        # Clean up temporary file
        os.remove(file_path)
        
        return jsonify({
            'success': True,
            'message': 'Files generated successfully',
            'regular_zip': zip_files['regular'],
            'highlighted_zip': zip_files['highlighted']
        })
    
    except Exception as e:
        logger.exception("Error processing request")
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(TEMP_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        logger.exception("Error downloading file")
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.errorhandler(404)
def page_not_found(e):
    return render_template('index.html'), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'An internal server error occurred'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
