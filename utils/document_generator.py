import os
import zipfile
import tempfile
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import logging

logger = logging.getLogger(__name__)

def create_word_document(questions_df, highlight_answers=False, class_name="", subject_name=""):
    """
    Creates a Word document with the given questions
    If highlight_answers is True, the correct answers are highlighted
    Optional class_name and subject_name parameters for document header
    """
    doc = Document()
    
    # Set document to landscape orientation
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    
    # Set narrower margins to maximize space
    section.left_margin = Inches(0.3)  # 0.3 inch
    section.right_margin = Inches(0.3)  # 0.3 inch
    section.top_margin = Inches(0.3)  # 0.3 inch
    section.bottom_margin = Inches(0.3)  # 0.3 inch
    
    # Add document header with subject and class name
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Create title text
    title_text = "ĐỀ THI"
    if subject_name:
        title_text += f" {subject_name}"
    if class_name:
        title_text += f" - {class_name}"
    
    # Add the title in bold, centered
    run = header.add_run(title_text)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)
    
    # Add student information section without underlines
    info_para = doc.add_paragraph()
    info_para.add_run("Mã sinh viên:                     Họ tên:                    ").font.name = 'Times New Roman'
    info_para.add_run().font.size = Pt(8)
    
    # Add a small space after the student info section
    doc.add_paragraph()
    
    # Set up 2-column layout for questions section
    section.start_type = WD_SECTION.NEW_PAGE
    
    # Create two columns
    sectPr = section._sectPr
    if not sectPr.xpath('./w:cols'):
        cols = OxmlElement('w:cols')
        cols.set(qn('w:num'), '2')
        sectPr.append(cols)
    else:
        cols = sectPr.xpath('./w:cols')[0]
        cols.set(qn('w:num'), '2')
    
    # Try to set font style as a default for the document
    try:
        style = doc.styles['Normal']
        # We'll handle font formatting at the run level instead
    except Exception as e:
        logger.debug(f"Error accessing style: {str(e)}")
        
    # Process questions
    total_questions = len(questions_df)
    
    for index, row in questions_df.iterrows():
        # Add question number and text (in bold)
        question_text = f"{index + 1}. {row['Câu hỏi']}"
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(1)  # Minimal space
        
        # Add question in bold with Times New Roman font
        run = paragraph.add_run(question_text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(8)  # Smaller font
        
        # Add options with minimal spacing and indentation
        options = ['A', 'B', 'C', 'D']
        correct_answer = row['đáp án']
        
        for option in options:
            option_text = f"{option}: {row[option]}"
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.left_indent = Pt(8)  # Smaller indentation
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            
            if highlight_answers and option == correct_answer:
                # Highlight correct answer with bold
                run = paragraph.add_run(option_text)
                run.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(8)
            else:
                run = paragraph.add_run(option_text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(8)
        
        # Add minimal separator between questions
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
    
    return doc

def generate_zip_files(questions_df, num_questions, num_versions, class_name="", subject_name=""):
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
            
            # Create regular document with class and subject name
            regular_doc = create_word_document(
                version_questions, 
                highlight_answers=False,
                class_name=class_name,
                subject_name=subject_name
            )
            
            # Create highlighted document with class and subject name
            highlighted_doc = create_word_document(
                version_questions, 
                highlight_answers=True,
                class_name=class_name,
                subject_name=subject_name
            )
            
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
