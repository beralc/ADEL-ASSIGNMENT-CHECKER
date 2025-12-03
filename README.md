# ADEL Assignment Checker

A Flask-based web application for bulk grading and providing AI-powered feedback on student assignments.

## Features

- **Bulk Processing**: Upload ZIP files containing multiple student assignments (PDF/DOCX)
- **AI-Powered Feedback**: Uses GPT-5-mini to generate personalized feedback based on custom rubrics
- **Real-time Progress**: Server-Sent Events for live processing updates
- **CSV Integration**: Matches student files to Moodle CSV exports
- **Multiple Task Types**: Support for reading, oral, and essay assignments
- **Smart Name Matching**: Handles Unicode corruption and fuzzy matching (80% similarity threshold)
- **Export Options**: Generate updated CSV and Excel reports with feedback and grades

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Create a `.env` file with your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   SECRET_KEY=your_secret_key_here
   ```
4. Run the application:
   ```bash
   python app.py
   ```
5. Open http://127.0.0.1:5000 in your browser

## Usage

1. Select the task type (reading, oral, or essay)
2. Upload a ZIP file containing student assignments
3. Upload the corresponding Moodle CSV file
4. Click "Process" to start bulk grading
5. Download the updated CSV and Excel reports when complete

## Task Configuration

Edit the JSON instruction files to customize feedback criteria:
- `instructions-reading.json`
- `instructions-oral.json`
- `instructions-essay.json`

## Requirements

- Python 3.7+
- Flask
- OpenAI API access
- PyPDF2, python-docx, pandas

## License

MIT
