from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
from werkzeug.utils import secure_filename
from docx import Document
from utils.formatter import format_document
from utils.analyzer import analyze_document

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html', file_info=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(request.url)

    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        doc = Document(filepath)

        # Analyze and format the document
        issues = analyze_document(doc)
        formatted_doc, changes_report = format_document(doc)

        processed_filename = f"formatted_{filename}"
        processed_filepath = os.path.join(app.config['OUTPUT_FOLDER'], processed_filename)
        os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
        formatted_doc.save(processed_filepath)

        return render_template(
            'results.html',
            original=filename,
            processed=processed_filename,
            issues=changes_report,
            file_info={
                'name': filename,
                'size': f"{os.path.getsize(filepath)/1024:.1f} KB"
            }
        )

    flash('Only .docx files allowed')
    return redirect(request.url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    app.run(host='0.0.0.0', port=5000)
