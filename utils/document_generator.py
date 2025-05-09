import os
import zipfile
import tempfile
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
import logging

logger = logging.getLogger(__name__)

def create_word_document(questions_df, highlight_answers=False):
    """
    Creates a Word document with the given questions
    If highlight_answers is True, the correct answers are highlighted
    """
    doc = Document()
    
    # Set document to landscape orientation
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    # Swap width and height is handled automatically when setting orientation
    
    # Set narrower margins to maximize space
    section.left_margin = Inches(0.5)  # 0.5 inch
    section.right_margin = Inches(0.5)  # 0.5 inch
    section.top_margin = Inches(0.5)  # 0.5 inch
    section.bottom_margin = Inches(0.5)  # 0.5 inch
    
    # Set document style with smaller font
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)  # Reduced font size
    
    # Set up 2-column layout to better utilize space
    section = doc.sections[0]
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num', '2')
    
    # Process questions
    total_questions = len(questions_df)
    
    for index, row in questions_df.iterrows():
        # Add question number and text (in bold)
        question_text = f"{index + 1}. {row['Câu hỏi']}"
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(question_text)
        run.bold = True  # Make questions bold
        paragraph.paragraph_format.space_after = Pt(2)  # Reduce space after paragraph
        
        # Add options with less indentation and spacing
        options = ['A', 'B', 'C', 'D']
        correct_answer = row['đáp án']
        
        for option in options:
            option_text = f"{option}: {row[option]}"
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.left_indent = Pt(12)  # Less indentation
            paragraph.paragraph_format.space_before = Pt(0)  # No space before
            paragraph.paragraph_format.space_after = Pt(0)  # No space after
            
            if highlight_answers and option == correct_answer:
                # Highlight correct answer with bold
                run = paragraph.add_run(option_text)
                run.bold = True
            else:
                paragraph.add_run(option_text)
        
        # Add minimal space between questions
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
    
    return doc

def generate_zip_files(questions_df, num_questions, num_versions):
    """
    Generates two ZIP files:
    1. Regular version without highlighted answers
    2. Version with highlighted answers
    
    Returns dictionary with paths to both ZIP files
    """
    temp_dir = tempfile.gettempdir()
    
    # Create unique filenames for the ZIP files
    regular_zip_filename = f"regular_quiz_{num_questions}q_{num_versions}v.zip"
    highlighted_zip_filename = f"highlighted_quiz_{num_questions}q_{num_versions}v.zip"
    
    regular_zip_path = os.path.join(temp_dir, regular_zip_filename)
    highlighted_zip_path = os.path.join(temp_dir, highlighted_zip_filename)
    
    # Create both ZIP files
    with zipfile.ZipFile(regular_zip_path, 'w') as regular_zip, \
         zipfile.ZipFile(highlighted_zip_path, 'w') as highlighted_zip:
        
        for version in range(1, num_versions + 1):
            # Get random questions for this version
            version_questions = get_random_questions(questions_df, num_questions)
            
            # Create regular document
            regular_doc = create_word_document(version_questions, highlight_answers=False)
            
            # Create highlighted document
            highlighted_doc = create_word_document(version_questions, highlight_answers=True)
            
            # Save documents to temporary files
            regular_doc_path = os.path.join(temp_dir, f"quiz_version_{version}.docx")
            highlighted_doc_path = os.path.join(temp_dir, f"quiz_version_{version}_answers.docx")
            
            regular_doc.save(regular_doc_path)
            highlighted_doc.save(highlighted_doc_path)
            
            # Add documents to ZIP files
            regular_zip.write(regular_doc_path, f"quiz_version_{version}.docx")
            highlighted_zip.write(highlighted_doc_path, f"quiz_version_{version}_answers.docx")
            
            # Clean up temporary document files
            os.remove(regular_doc_path)
            os.remove(highlighted_doc_path)
    
    return {
        'regular': regular_zip_filename,
        'highlighted': highlighted_zip_filename
    }

def get_random_questions(df, num_questions):
    """
    Gets a random sample of questions from the dataframe
    """
    # Return a random sample of the dataframe
    return df.sample(n=num_questions).reset_index(drop=True)
