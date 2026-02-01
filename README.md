# Syllabus Parser

A Flask web application for parsing academic syllabi and generating comprehensive result sheets.

## Features

- **Step 1**: Upload syllabus PDF to extract course information
- **Step 2**: Upload student list Excel file
- **Step 3**: Upload BE Results Excel file with marks
- **Step 4**: Generate comprehensive result sheet combining all data

### Data Extraction
- Extracts SEM, SESSION, SGPA from results sheets
- Parses course codes, credits, and mark components
- Supports persistent student records across sessions

## Installation

```bash
pip install -r requirements.txt
python app.py
```

## Usage

1. Open http://localhost:5501
2. Upload syllabus PDF in Step 1
3. Upload student list in Step 2
4. Upload BE Results in Step 3
5. Generate and download result sheet

## Tech Stack

- Flask (Python)
- SQLite
- OpenPyXL
- PDFPlumber

## Database

The application uses SQLite with the following tables:
- `student_master` - Persistent student records
- `classes`, `courses`, `students`, `marks` - Session data
- `uploads` - File upload history for duplicate detection
