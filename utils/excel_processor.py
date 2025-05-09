import pandas as pd
import numpy as np
import logging

logger = logging.getLogger(__name__)

def validate_excel_file(file_path):
    """
    Validates that the uploaded Excel file has the required columns
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Define required columns
        required_columns = ['Câu hỏi', 'A', 'B', 'C', 'D', 'đáp án']
        
        # Check if all required columns are present
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return {'error': f'Missing required columns: {", ".join(missing_columns)}'}
        
        # Check if any required columns are empty
        empty_columns = []
        for col in required_columns:
            if df[col].isna().any():
                empty_columns.append(col)
        
        if empty_columns:
            return {'error': f'The following columns contain empty values: {", ".join(empty_columns)}'}
        
        # Validate answer column values
        valid_answers = ['A', 'B', 'C', 'D']
        invalid_answers = df[~df['đáp án'].isin(valid_answers)]['đáp án'].unique()
        if len(invalid_answers) > 0:
            return {'error': f'Invalid answer values found: {", ".join(str(a) for a in invalid_answers)}. Valid answers are: A, B, C, D'}
        
        return {'success': True}
    
    except Exception as e:
        logger.exception("Error validating Excel file")
        return {'error': f'Error processing Excel file: {str(e)}'}

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
