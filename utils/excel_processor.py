import pandas as pd
import numpy as np
import logging

logger = logging.getLogger(__name__)

def validate_excel_format(df):
    """
    Validates that a dataframe has the required columns and format
    """
    try:
        # Define required columns
        required_columns = ['Câu hỏi', 'A', 'B', 'C', 'D', 'đáp án']
        
        # Check if all required columns are present
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return {'error': f'Thiếu các cột yêu cầu: {", ".join(missing_columns)}'}
        
        # Check if any required columns are empty
        empty_columns = []
        for col in required_columns:
            # Check if the column has any NaN values
            if pd.isna(df[col]).any(skipna=False):
                empty_columns.append(col)
        
        if empty_columns:
            return {'error': f'Các cột sau đây chứa giá trị trống: {", ".join(empty_columns)}'}
        
        # Validate answer column values
        valid_answers = ['A', 'B', 'C', 'D']
        # Check for invalid answers
        invalid_answers = []
        for answer in df['đáp án'].unique():
            if answer not in valid_answers:
                invalid_answers.append(str(answer))
        
        if len(invalid_answers) > 0:
            return {'error': f'Tìm thấy giá trị đáp án không hợp lệ: {", ".join(invalid_answers)}. Đáp án hợp lệ là: A, B, C, D'}
        
        return {'success': True}
    
    except Exception as e:
        logger.exception("Error validating Excel format")
        return {'error': f'Lỗi xử lý dữ liệu Excel: {str(e)}'}

def validate_excel_file(file_path):
    """
    Validates that the uploaded Excel file has at least one sheet with the required format
    """
    try:
        # Check if we can open the file as Excel
        try:
            excel_file = pd.ExcelFile(file_path)
        except Exception:
            return {'error': 'File không phải là file Excel hợp lệ'}
        
        # Get all sheet names
        sheet_names = excel_file.sheet_names
        if not sheet_names:
            return {'error': 'File Excel không chứa sheet nào'}
            
        # Validate first sheet to check format
        first_sheet = pd.read_excel(file_path, sheet_name=sheet_names[0])
        return validate_excel_format(first_sheet)
    
    except Exception as e:
        logger.exception("Error validating Excel file")
        return {'error': f'Lỗi xử lý file Excel: {str(e)}'}

def get_random_questions(df, num_questions):
    """
    Gets a random sample of questions from the dataframe
    """
    # Check if we have enough questions
    total_questions = len(df)
    if total_questions < num_questions:
        raise ValueError(f"Not enough questions. Requested {num_questions}, but only {total_questions} are available.")
    
    # Get random sample
    sample_indices = np.random.choice(total_questions, num_questions, replace=False)
    sampled_questions = df.iloc[sample_indices].reset_index(drop=True)
    
    return sampled_questions
