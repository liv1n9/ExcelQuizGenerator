import os
import logging
import tempfile
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from werkzeug.utils import secure_filename
from utils.excel_processor import validate_excel_file, validate_excel_format, get_random_questions
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
            return jsonify({'error': 'Không tìm thấy file'}), 400
        
        file = request.files['excelFile']
        
        # If user does not select file
        if file.filename == '':
            return jsonify({'error': 'Chưa chọn file'}), 400
        
        # Check if it's an allowed file type
        if not allowed_file(file.filename):
            return jsonify({'error': 'Loại file không hợp lệ. Vui lòng tải lên file Excel (.xlsx hoặc .xls)'}), 400
        
        # Get form data
        num_questions = request.form.get('numQuestions')
        num_versions = request.form.get('numVersions')
        
        # Validate form data
        if not num_questions or not num_versions:
            return jsonify({'error': 'Vui lòng nhập số câu hỏi và số phiên bản'}), 400
        
        try:
            num_questions = int(num_questions)
            num_versions = int(num_versions)
        except ValueError:
            return jsonify({'error': 'Số câu hỏi và số phiên bản phải là số nguyên'}), 400
        
        if num_questions <= 0 or num_versions <= 0:
            return jsonify({'error': 'Số câu hỏi và số phiên bản phải là số dương'}), 400
        
        # Save the file to a temporary location
        if file.filename:
            filename = secure_filename(file.filename)
            file_path = os.path.join(TEMP_FOLDER, filename)
            file.save(file_path)
        else:
            return jsonify({'error': 'Tên file không hợp lệ'}), 400
        
        # Validate Excel file format
        validation_result = validate_excel_file(file_path)
        if 'error' in validation_result:
            os.remove(file_path)  # Remove temporary file
            return jsonify(validation_result), 400
        
        # Generate the ZIP files
        # Read all sheets from Excel file
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        # Concatenate all sheets into one dataframe
        all_questions = []
        for sheet_name in sheet_names:
            sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Validate sheet format
            validation_result = validate_excel_format(sheet_df)
            if 'error' in validation_result:
                os.remove(file_path)  # Remove temporary file
                return jsonify({'error': f'Lỗi trong sheet "{sheet_name}": {validation_result["error"]}'}), 400
            
            all_questions.append(sheet_df)
        
        # Combine all sheets
        questions_df = pd.concat(all_questions, ignore_index=True)
        
        # Make sure we have enough questions
        total_questions = len(questions_df)
        if total_questions < num_questions:
            os.remove(file_path)  # Remove temporary file
            return jsonify({'error': f'File Excel chỉ chứa {total_questions} câu hỏi từ tất cả các sheet, nhưng bạn yêu cầu {num_questions} câu'}), 400
        
        zip_files = generate_zip_files(questions_df, num_questions, num_versions)
        
        # Clean up temporary file
        os.remove(file_path)
        
        return jsonify({
            'success': True,
            'message': 'Tạo file thành công',
            'regular_zip': zip_files['regular'],
            'highlighted_zip': zip_files['highlighted']
        })
    
    except Exception as e:
        logger.exception("Error processing request")
        return jsonify({'error': f'Đã xảy ra lỗi: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(TEMP_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        logger.exception("Error downloading file")
        return jsonify({'error': f'Đã xảy ra lỗi: {str(e)}'}), 500

@app.errorhandler(404)
def page_not_found(e):
    return render_template('index.html'), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Đã xảy ra lỗi nội bộ máy chủ'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
