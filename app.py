import os
import json
import re
import logging
import time
import zipfile
import tempfile
import shutil
import unicodedata
from flask import Flask, render_template, request, jsonify, Response, send_file
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from openai import OpenAI
import PyPDF2
import docx
import pandas as pd
import uuid

# Set logging to INFO to see matching details
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'docx', 'zip', 'csv'}
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024 * 1024  # 2GB max
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Store for processing results and progress
processing_sessions = {}

# Initialize OpenAI client
client = OpenAI()

# Task type to JSON file mapping
TASK_CONFIGS = {
    'reading': 'instructions-reading.json',
    'oral': 'instructions-oral.json',
    'essay': 'instructions-essay.json'
}

def allowed_file(filename, file_types=None):
    if file_types is None:
        file_types = app.config['ALLOWED_EXTENSIONS']
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in file_types

def extract_text(file_path):
    text = ""
    if file_path.endswith('.pdf'):
        with open(file_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() or ""
    elif file_path.endswith('.docx'):
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    return text

def enforce_instructions(feedback):
    replacements = {
        "commendable": "good",
        "excellent": "very good",
        "innovative": "new",
        "unique": "special",
        "fostering": "helping",
        "comprehensive": "complete",
        "ensure": "make sure",
        "showcases": "shows",
        "great": "good",
    }
    for word, repl in replacements.items():
        feedback = re.sub(rf'\b{word}\b', repl, feedback, flags=re.IGNORECASE)

    if not feedback.startswith("You") and not re.match(r'^\w+,', feedback):
        feedback = "You, " + feedback

    max_length = 1000
    if len(feedback) > max_length:
        feedback = feedback[:max_length] + "..."

    return feedback

def generate_feedback(file_path, instructions):
    text = extract_text(file_path)
    try:
        prompt = (
            f"You are reviewing a student assignment. The text below was extracted from a PDF/document.\n\n"
            f"⚠️ CRITICAL INSTRUCTION - PDF ARTIFACTS ⚠️\n"
            f"The extracted text contains PDF decoding artifacts such as:\n"
            f"- Extra spaces in words (e.g., 'm aking' instead of 'making', 'rem ain' instead of 'remain')\n"
            f"- Broken words (e.g., 'ber nar do' instead of 'bernardo')\n"
            f"- Missing spaces or unusual formatting\n\n"
            f"These are 100% NOT student mistakes. They are technical artifacts from PDF text extraction.\n"
            f"DO NOT mention, reference, or penalize ANY spacing issues, broken words, or formatting problems.\n"
            f"COMPLETELY IGNORE all spacing and word-breaking artifacts.\n\n"
            f"⚠️ CRITICAL INSTRUCTION - DO NOT CORRECT GRAMMAR ⚠️\n"
            f"DO NOT point out or correct grammar mistakes in your feedback.\n"
            f"DO NOT mention spelling errors.\n"
            f"DO NOT provide language corrections.\n"
            f"Focus on the CONTENT and IDEAS only, not on language mechanics.\n\n"
            f"⚠️ CRITICAL INSTRUCTION - PUNCTUATION ⚠️\n"
            f"DO NOT use em dashes (—) or en dashes (–) in your feedback.\n"
            f"ONLY use regular hyphens (-) for punctuation.\n"
            f"Use parentheses ( ) or commas instead of dashes for clarification.\n\n"
            f"Focus ONLY on:\n"
            f"- The actual content and ideas\n"
            f"- Whether the assignment meets the objectives\n"
            f"- The quality and originality of the activities designed\n"
            f"- How well the assignment follows the instructions\n"
            f"- Whether required elements are present (e.g., TAVI and TALO for reading)\n\n"
            f"Please review the following student assignment and provide short feedback using simple language "
            f"based on these instructions:\n\n{json.dumps(instructions, indent=2)}\n\n"
            f"Student's text:\n{text}"
        )
        response = client.chat.completions.create(
            model="gpt-5-mini",
            messages=[
                {"role": "system", "content": "You are a teaching assistant who carefully distinguishes between genuine student errors and text extraction artifacts."},
                {"role": "user", "content": prompt}
            ]
        )
        feedback = response.choices[0].message.content
        return enforce_instructions(feedback)
    except Exception as e:
        logging.error(f"Error generating feedback for {file_path}: {e}")
        return f"Error: {e}"

def clean_corrupted_name(name):
    """Remove corrupted box-drawing and other non-letter characters from names"""
    # Remove box-drawing characters and other corruption artifacts
    cleaned = []
    for c in name:
        cat = unicodedata.category(c)
        # Keep letters, spaces, hyphens, and apostrophes
        if cat.startswith('L') or c in ' -\'':
            cleaned.append(c)
        # Skip box-drawing (So), combining marks (Mn), and other weird characters
    return ''.join(cleaned)

def extract_first_name_from_filename(filename):
    name_part = os.path.splitext(filename)[0]
    # Try to normalize to NFC, then clean corruption
    name_part = unicodedata.normalize('NFC', name_part)
    name_part = clean_corrupted_name(name_part)
    words = name_part.strip().split()
    return words[0].capitalize() if words else "Student"

def normalize_name(name):
    def remove_accents_and_corruption(text):
        # First normalize to NFD to separate base characters from combining marks
        nfd = unicodedata.normalize('NFD', text)
        # Keep only ASCII letters and spaces - this removes both accents AND corruption
        result = []
        for c in nfd:
            cat = unicodedata.category(c)
            # Only keep ASCII letters (a-z, A-Z) and spaces
            # This removes: accents, umlauts (ü), box characters, etc.
            if c.isascii():
                if c.isalpha() or c.isspace():
                    result.append(c)
        return ''.join(result)

    name = remove_accents_and_corruption(name)
    # Final cleanup: ensure only alphanumeric and spaces
    name = ''.join(c for c in name if c.isalnum() or c.isspace())
    return ' '.join(name.strip().lower().split()[:2])

def split_feedback_and_score(comment):
    match = re.search(r"\bscore\s*:\s*(\d(?:\.\d)?)", comment, re.IGNORECASE)
    if match:
        score = match.group(1)
        comment = comment[:match.start()].strip()
        return comment, score
    return comment, ""

def calculate_match_percentage(name1, name2):
    """Calculate similarity percentage between two names"""
    name1_normalized = normalize_name(name1)
    name2_normalized = normalize_name(name2)

    if name1_normalized == name2_normalized:
        return 100

    # Split into words
    words1 = name1_normalized.split()
    words2 = name2_normalized.split()

    if not words1 or not words2:
        return 0

    # Check if first two words are similar (allowing for corruption)
    matches = 0
    total = max(len(words1), len(words2), 2)

    for i in range(min(2, min(len(words1), len(words2)))):
        w1 = words1[i]
        w2 = words2[i]

        # Exact match
        if w1 == w2:
            matches += 1
        else:
            # Check similarity - if 80% of characters match, consider it a match
            # This handles corruption like "fautima" vs "fatima"
            if len(w1) >= 3 and len(w2) >= 3:
                common_chars = sum(1 for c in w1 if c in w2)
                similarity = common_chars / max(len(w1), len(w2))
                if similarity >= 0.8:  # 80% similar
                    matches += 0.9  # Partial match

    return int((matches / 2) * 100)  # Out of 2 words

def process_bulk_marking(session_id, task_type, zip_path, csv_path):
    """Process all files and yield progress updates"""
    try:
        # Load instructions based on task type
        instructions_file = TASK_CONFIGS.get(task_type)
        if not instructions_file:
            yield f"data: {json.dumps({'error': 'Invalid task type'})}\n\n"
            return

        with open(instructions_file, 'r') as f:
            instructions = json.load(f)

        # Read CSV - preserve all original data
        df = pd.read_csv(csv_path, encoding='utf-8')

        # Create a normalized version of names for matching ONLY (don't modify the original)
        # Store original Full name values
        df_normalized_names = df["Full name"].apply(lambda x: unicodedata.normalize('NFC', str(x)) if pd.notna(x) else x)

        # Ensure columns exist - but don't change existing data
        if "Feedback comments" not in df.columns:
            df["Feedback comments"] = ""  # Empty string for new column

        if "Grade" not in df.columns:
            df["Grade"] = None  # None for new column

        # Extract ZIP to temp directory
        temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Get all PDF/DOCX files - normalize immediately to avoid NFD display issues
        raw_files = [f for f in os.listdir(temp_dir) if allowed_file(f, {'pdf', 'docx'})]
        files = [unicodedata.normalize('NFC', f) for f in raw_files]
        total_files = len(files)

        # Debug: Check if normalization is working
        if raw_files and raw_files[0] != files[0]:
            logging.info(f"Normalization working: {repr(raw_files[0][:20])} -> {repr(files[0][:20])}")
        else:
            logging.info(f"No normalization needed or not working: {repr(raw_files[0][:20]) if raw_files else 'no files'}")

        results = []

        for idx, fname in enumerate(files, 1):
            try:
                # fname is already NFC-normalized at this point
                # But we need to use the original (potentially NFD) filename for file operations
                fname_original = unicodedata.normalize('NFD', fname)
                path = os.path.join(temp_dir, fname_original)

                # Clean corrupted characters from the filename for display and matching
                fname_clean = clean_corrupted_name(fname)
                student_name_from_file = os.path.splitext(fname_clean)[0].strip()
                first_name = extract_first_name_from_filename(fname)

                # Debug: Log the filename and normalized version
                logging.info(f"Processing file {idx}/{total_files}: {student_name_from_file}")

                # Generate feedback
                full_feedback = generate_feedback(path, instructions)
                comment, score = split_feedback_and_score(full_feedback)

                if comment:
                    # Remove leading name but keep "You have..."
                    comment = re.sub(r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s*,\s*", "", comment, flags=re.IGNORECASE)
                    if not comment.lower().startswith(first_name.lower()):
                        comment = f"{first_name}, {comment[0].lower() + comment[1:]}" if comment else ""

                # Match with CSV - try exact match first
                normalized_filename = normalize_name(student_name_from_file)
                # Use the pre-normalized names for matching
                normalized_names = df_normalized_names.apply(normalize_name)

                matches = normalized_names[normalized_names == normalized_filename]

                if not matches.empty:
                    # Exact match found - use ORIGINAL name from CSV, not normalized
                    matched_name = df.loc[matches.index[0], "Full name"]

                    # Extract correct first name from matched CSV name
                    correct_first_name = matched_name.split()[0]

                    # Fix the first name in the comment to use the correct one from CSV
                    if comment:
                        # Replace the corrupted first name with the correct one
                        comment = re.sub(r'^[^,]+,', f'{correct_first_name},', comment)

                    df.loc[df["Full name"] == matched_name, "Feedback comments"] = comment
                    df.loc[df["Full name"] == matched_name, "Grade"] = float(score) if score else None

                    match_percentage = 100
                    match_status = "success"
                    logging.info(f"✓ Exact match: {student_name_from_file} -> {matched_name}")
                else:
                    # Try fuzzy match using character-level similarity
                    best_match = None
                    best_score = 0

                    for idx, csv_name in enumerate(df["Full name"]):
                        # Use the fuzzy matching function
                        similarity_percentage = calculate_match_percentage(student_name_from_file, csv_name)

                        if similarity_percentage > best_score:
                            best_score = similarity_percentage
                            best_match = (idx, csv_name)

                    # If we found a match with at least 80% similarity, use it
                    # This allows for corruption like "Faütima" vs "Fátima"
                    if best_match and best_score >= 80:
                        matched_name = best_match[1]

                        # Extract correct first name from matched CSV name
                        correct_first_name = matched_name.split()[0]

                        # Fix the first name in the comment to use the correct one from CSV
                        if comment:
                            # Replace the corrupted first name with the correct one
                            comment = re.sub(r'^[^,]+,', f'{correct_first_name},', comment)

                        df.loc[df["Full name"] == matched_name, "Feedback comments"] = comment
                        df.loc[df["Full name"] == matched_name, "Grade"] = float(score) if score else None
                        match_percentage = int(best_score)  # best_score is already a percentage
                        match_status = "success"
                        logging.warning(f"⚠ Fuzzy match ({match_percentage}%): {student_name_from_file} -> {matched_name}")
                    else:
                        matched_name = None
                        match_percentage = 0
                        match_status = "no_match"
                        logging.error(f"✗ No match: {student_name_from_file} (normalized: {normalized_filename})")

                result = {
                    "file_name": fname_clean,  # Cleaned filename without corruption
                    "student_name": student_name_from_file,  # Already using cleaned version
                    "matched_name": matched_name,
                    "match_percentage": match_percentage,
                    "match_status": match_status,
                    "score": score,  # Grade score from AI feedback
                    "comment": comment,
                    "comment_preview": comment[:150] + "..." if len(comment) > 150 else comment
                }

                results.append(result)

                # Send progress update
                progress_data = {
                    "type": "progress",
                    "current": idx,
                    "total": total_files,
                    "percentage": int((idx / total_files) * 100),
                    "result": result
                }

                yield f"data: {json.dumps(progress_data, ensure_ascii=False)}\n\n"

            except Exception as e:
                # Use cleaned filename for error display
                fname_clean_err = clean_corrupted_name(fname) if fname else "Unknown file"
                error_data = {
                    "type": "error",
                    "file": fname_clean_err,
                    "message": str(e)
                }
                yield f"data: {json.dumps(error_data, ensure_ascii=False)}\n\n"

        # Save results
        ts = time.strftime("%Y%m%d%H%M%S")

        # Save updated CSV - preserve formatting as much as possible
        csv_filename = os.path.basename(csv_path)
        output_csv = os.path.join(app.config['UPLOAD_FOLDER'],
                                  os.path.splitext(csv_filename)[0] + f"_with_feedback_{ts}.csv")
        # Use quoting=csv.QUOTE_MINIMAL to preserve original quoting style
        # UTF-8 with BOM so Excel/Word recognize the encoding automatically
        df.to_csv(output_csv, index=False, quoting=1, encoding='utf-8-sig')  # BOM for Excel compatibility

        # Save Excel
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], f"bulk_feedback_{ts}.xlsx")
        with pd.ExcelWriter(excel_path) as writer:
            df_feedback = pd.DataFrame(results)[['file_name', 'student_name', 'comment']]
            df_grades = pd.DataFrame(results)[['file_name', 'student_name', 'score']]
            df_combined = pd.DataFrame(results)

            df_feedback.to_excel(writer, sheet_name="Feedback", index=False)
            df_grades.to_excel(writer, sheet_name="Scores", index=False)
            df_combined.to_excel(writer, sheet_name="Combined", index=False)

        # Update session data with results (preserve existing data)
        processing_sessions[session_id].update({
            'results': results,
            'csv_path': output_csv,
            'excel_path': excel_path,
            'csv_filename': os.path.basename(output_csv),
            'excel_filename': os.path.basename(excel_path)
        })

        # Send completion
        completion_data = {
            "type": "complete",
            "total": total_files,
            "csv_filename": os.path.basename(output_csv),
            "excel_filename": os.path.basename(excel_path)
        }
        yield f"data: {json.dumps(completion_data, ensure_ascii=False)}\n\n"

        # Cleanup
        shutil.rmtree(temp_dir)

    except Exception as e:
        error_data = {
            "type": "fatal_error",
            "message": str(e)
        }
        yield f"data: {json.dumps(error_data, ensure_ascii=False)}\n\n"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    """Handle file upload and start processing"""
    try:
        task_type = request.form.get('task_type')
        zip_file = request.files.get('zip_file')
        csv_file = request.files.get('csv_file')

        if not task_type or task_type not in TASK_CONFIGS:
            return jsonify({'error': 'Invalid task type'}), 400

        if not zip_file or not csv_file:
            return jsonify({'error': 'Both ZIP and CSV files are required'}), 400

        if not allowed_file(zip_file.filename, {'zip'}):
            return jsonify({'error': 'Invalid ZIP file'}), 400

        if not allowed_file(csv_file.filename, {'csv'}):
            return jsonify({'error': 'Invalid CSV file'}), 400

        # Save files
        zip_filename = secure_filename(zip_file.filename)
        csv_filename = secure_filename(csv_file.filename)

        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        csv_path = os.path.join(app.config['UPLOAD_FOLDER'], csv_filename)

        zip_file.save(zip_path)
        csv_file.save(csv_path)

        # Generate session ID
        session_id = str(uuid.uuid4())

        # Store file paths for this session
        processing_sessions[session_id] = {
            'zip_path': zip_path,
            'csv_path': csv_path,
            'task_type': task_type
        }

        # Return session ID for SSE connection
        return jsonify({
            'session_id': session_id,
            'task_type': task_type
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/stream/<session_id>')
def stream(session_id):
    """Server-Sent Events endpoint for real-time updates"""
    task_type = request.args.get('task_type')

    # Get file paths for this specific session
    if session_id not in processing_sessions:
        return Response(
            f"data: {json.dumps({'type': 'fatal_error', 'message': 'Session not found'}, ensure_ascii=False)}\n\n",
            mimetype='text/event-stream; charset=utf-8'
        )

    session_data = processing_sessions[session_id]
    zip_path = session_data['zip_path']
    csv_path = session_data['csv_path']

    if not os.path.exists(zip_path) or not os.path.exists(csv_path):
        return Response(
            f"data: {json.dumps({'type': 'fatal_error', 'message': 'Files not found'}, ensure_ascii=False)}\n\n",
            mimetype='text/event-stream; charset=utf-8'
        )

    return Response(
        process_bulk_marking(session_id, task_type, zip_path, csv_path),
        mimetype='text/event-stream; charset=utf-8'
    )

@app.route('/download/<session_id>/<file_type>')
def download(session_id, file_type):
    """Download generated files"""
    if session_id not in processing_sessions:
        return jsonify({'error': 'Session not found'}), 404

    session_data = processing_sessions[session_id]

    if file_type == 'csv':
        return send_file(
            session_data['csv_path'],
            as_attachment=True,
            download_name=session_data['csv_filename']
        )
    elif file_type == 'excel':
        return send_file(
            session_data['excel_path'],
            as_attachment=True,
            download_name=session_data['excel_filename']
        )
    else:
        return jsonify({'error': 'Invalid file type'}), 400

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 2GB. Consider splitting large batches into smaller groups.'}), 413

if __name__ == '__main__':
    app.run(debug=True)
