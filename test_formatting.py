"""
Test script to verify subscript/superscript formatting preservation
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.cell.rich_text import TextBlock, CellRichText
from utils.excel_processor import read_excel_with_formatting, extract_rich_text_from_cell
from utils.document_generator import create_word_document
import os

def create_test_excel():
    """Create a test Excel file with subscript and superscript"""
    wb = Workbook()
    ws = wb.active
    
    # Add headers
    headers = ['Câu hỏi', 'A', 'B', 'C', 'D', 'đáp án']
    ws.append(headers)
    
    # Create rich text with subscript for H2O
    h2o_text = CellRichText()
    h2o_text.append(TextBlock(Font(), 'H'))
    h2o_text.append(TextBlock(Font(vertAlign='subscript'), '2'))
    h2o_text.append(TextBlock(Font(), 'O'))
    
    # Create rich text with superscript for x^2
    x2_text = CellRichText()
    x2_text.append(TextBlock(Font(), 'x'))
    x2_text.append(TextBlock(Font(vertAlign='superscript'), '2'))
    
    # Add a test question
    ws.append(['Công thức của nước là gì?', h2o_text, 'H3O', 'H2SO4', 'HCl', 'A'])
    ws.append(['Kết quả của x^2 khi x=3?', '6', x2_text, '9', '12', 'C'])
    
    test_file = '/tmp/test_formatting.xlsx'
    wb.save(test_file)
    print(f"✓ Created test Excel file: {test_file}")
    return test_file

def test_read_formatting(file_path):
    """Test reading formatting from Excel"""
    print("\n=== Testing read_excel_with_formatting ===")
    df = read_excel_with_formatting(file_path)
    
    print(f"✓ Read {len(df)} rows")
    print(f"Columns: {df.columns.tolist()}")
    
    # Check if _formatting column exists
    if '_formatting' in df.columns:
        print("✓ _formatting column exists")
        
        # Check first row formatting
        first_row_formatting = df.iloc[0]['_formatting']
        print(f"\nFirst row formatting: {first_row_formatting}")
        
        if 'A' in first_row_formatting:
            print(f"Option A formatting: {first_row_formatting['A']}")
            # Check if subscript is detected
            has_subscript = any(is_sub for _, is_sub, _ in first_row_formatting['A'])
            if has_subscript:
                print("✓ Subscript detected in Option A (H2O)")
            else:
                print("✗ Subscript NOT detected in Option A")
        
        # Check second row formatting
        if len(df) > 1:
            second_row_formatting = df.iloc[1]['_formatting']
            if 'B' in second_row_formatting:
                print(f"\nOption B (row 2) formatting: {second_row_formatting['B']}")
                has_superscript = any(is_sup for _, _, is_sup in second_row_formatting['B'])
                if has_superscript:
                    print("✓ Superscript detected in Option B (x^2)")
                else:
                    print("✗ Superscript NOT detected in Option B")
    else:
        print("✗ _formatting column NOT found")
    
    return df

def test_create_document(df):
    """Test creating Word document with formatting"""
    print("\n=== Testing create_word_document ===")
    
    doc = create_word_document(df, highlight_answers=False, 
                              class_name="Test Class", 
                              subject_name="Chemistry",
                              version=0, num_columns=1)
    
    output_file = '/tmp/test_formatted_output.docx'
    doc.save(output_file)
    print(f"✓ Created Word document: {output_file}")
    print(f"  Please open this file and verify that:")
    print(f"  - H2O has subscript '2'")
    print(f"  - x^2 has superscript '2'")

if __name__ == '__main__':
    print("Starting formatting test...")
    
    # Create test Excel
    test_file = create_test_excel()
    
    # Test reading with formatting
    df = test_read_formatting(test_file)
    
    # Test creating document
    test_create_document(df)
    
    print("\n=== Test Complete ===")
    print("Please check the output file: /tmp/test_formatted_output.docx")
