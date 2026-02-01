# Syllabus Parser

A Flask-based web application for parsing academic syllabus documents and generating marks distribution tables, student course matrices, and comprehensive result sheets.

## Features

- **Step 1: Syllabus Upload** - Parse PDF syllabus to extract course information (Course Code, Title, Credits, Marks Distribution)
- **Step 2: Student List Upload** - Upload Excel file with student details (USN, Name, Section, etc.)
- **Step 3: BE Results Upload** - Parse examination results with marks (ISE-TH, MSE, ESE, SGPA, etc.)
- **Step 4: Result Sheet Generation** - Generate comprehensive result sheets combining all data

## Column Mappings

| Column | Source | Description |
|--------|--------|-------------|
| SUB_CE_MRKS | ISE-TH | Internal/Continuous Evaluation marks |
| SUB_TH_MRKS | MSE + ESE | Theory marks (mid-sem + end-sem) |
| SUB_PR_MRKS | ISE-PR/TU | Practical marks |
| SUB_VV_MRKS | ESE (for Projects only) | Viva marks |
| SGPA | From Step 3 or calculated | Semester GPA |

## Installation

1. Clone the repository:
```bash
git clone https://github.com/frcrcfaculty-commits/syllabus-parser.git
cd syllabus-parser
```

2. Create a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python app.py
```

5. Open http://localhost:5501 in your browser

## Database

The application uses SQLite with the following tables:
- `student_master` - Persistent student records across sessions
- `classes`, `courses`, `students`, `marks` - Session data
- `uploads` - File upload history for duplicate detection

## Tech Stack

- **Backend**: Python, Flask
- **Database**: SQLite
- **PDF Parsing**: pdfplumber
- **Excel Handling**: openpyxl
- **Frontend**: HTML, CSS, JavaScript

## License

MIT License
