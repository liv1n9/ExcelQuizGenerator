# Quiz Generator Application

## Overview

A Vietnamese language web application that generates randomized multiple-choice quiz exams from Excel question banks. The system allows educators to upload Excel files containing questions and answers, then automatically generates multiple versions of exams in Word document format. The application supports two output types: standard student exams and teacher versions with highlighted correct answers.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture

**Technology Stack**: Bootstrap 5 with dark theme, vanilla JavaScript, Font Awesome icons

**Design Pattern**: Server-side rendered templates with AJAX-based file upload and download

The frontend uses Flask's Jinja2 templating engine to serve a single-page application. The UI is built with Bootstrap 5's dark theme for a modern appearance and includes:

- Form validation using Bootstrap's built-in validation classes
- File upload interface with drag-and-drop support (indicated by CSS animations)
- Loading overlay for asynchronous operations
- Alert system for user feedback
- Download buttons for generated exam files

**Rationale**: Server-side rendering with progressive enhancement provides better initial load performance and SEO while maintaining interactivity through AJAX for file operations.

### Backend Architecture

**Framework**: Flask (Python web framework)

**Application Structure**: Simple monolithic architecture with utility modules

The application follows a lightweight, modular structure:

- `app.py`: Main Flask application with route handlers
- `main.py`: Application entry point
- `utils/excel_processor.py`: Excel file validation and question randomization logic
- `utils/document_generator.py`: Word document generation with python-docx

**Key Design Decisions**:

1. **File Upload Handling**: Uses temporary file storage (`tempfile.gettempdir()`) to avoid persisting uploaded files, reducing security risks and storage requirements

2. **Validation Layer**: Two-stage validation process - file type validation at upload, then content validation (required columns, answer format, empty values)

3. **Randomization**: Supports multiple Excel sheets with uniform format, aggregating questions across sheets before random selection

4. **Document Generation**: Creates landscape-oriented Word documents with customizable headers (class name, subject) and two versions (regular and answer-highlighted) packaged in ZIP files

**Pros**: Simple to deploy, easy to maintain, minimal dependencies
**Cons**: Single-threaded processing may struggle with concurrent users or large files

### Data Storage Solutions

**File Storage**: Temporary filesystem storage using Python's `tempfile` module

**Data Processing**: In-memory processing with pandas DataFrames

The application does not use a persistent database. Instead:

- Uploaded Excel files are temporarily stored during processing
- Questions are loaded into pandas DataFrames for manipulation
- Generated Word documents are stored temporarily before being sent to the user

**Rationale**: For this use case, persistent storage is unnecessary since each session is independent. This approach reduces infrastructure complexity and eliminates data privacy concerns since no user data is retained.

**Alternatives Considered**: 
- Database storage for question banks: Rejected due to added complexity for a tool meant for one-time batch processing
- Cloud storage for generated files: Rejected to avoid vendor lock-in and reduce latency

### Authentication and Authorization

**Current Implementation**: None

The application does not implement authentication or authorization. It's designed as a utility tool for individual educators to use locally or in trusted environments.

**Future Considerations**: If deployed publicly, would need to add session management and user accounts to prevent abuse.

## External Dependencies

### Python Libraries

**Core Framework**:
- `Flask`: Web framework for routing and request handling
- `Werkzeug`: WSGI utilities (included with Flask) for secure filename handling

**Data Processing**:
- `pandas`: Excel file reading and DataFrame manipulation for question randomization
- `numpy`: Numerical operations (used by pandas)
- `openpyxl` or `xlrd`: Excel file format support (implicit pandas dependency)

**Document Generation**:
- `python-docx`: Word document (.docx) creation with formatting support
  - Handles document structure, paragraphs, tables
  - Provides text formatting (bold, font size, color)
  - Supports page layout (orientation, margins)

**File Operations**:
- `zipfile`: Standard library for creating ZIP archives of generated documents
- `tempfile`: Standard library for temporary file management

### Frontend Libraries (CDN-hosted)

**UI Framework**:
- Bootstrap 5.x with Replit dark theme: Responsive UI components and grid system
- Font Awesome 6.4.0: Icon library for UI elements

### File Format Requirements

**Input**: Excel files (.xlsx or .xls) with required columns:
- "Câu hỏi" (Question)
- "A", "B", "C", "D" (Answer options)
- "đáp án" (Correct answer)
- Optional: "Phân loại" (Classification/Category)

**Output**: 
- Word documents (.docx) in landscape orientation
- ZIP archives containing multiple exam versions

### Environment Variables

**SESSION_SECRET**: Flask session encryption key (defaults to "default_secret_key" if not set)

### Deployment Requirements

- Python 3.11+
- Web server capable of handling file uploads (default: Flask development server on port 5000)
- Sufficient temporary storage for concurrent file processing operations