from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
import openai
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

openai.api_key = os.getenv('OPENAI_API_KEY')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def analyze_document(doc):
    issues = []
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        if not first_para.text.strip().isupper():
            issues.append("First paragraph formatted as Title")
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.vertical_alignment != WD_ALIGN_VERTICAL.CENTER:
                    issues.append("Fixed table cell alignment")
    return issues

def justify_tables(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.lower()
                if any(char.isdigit() for char in text):
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                elif text.isupper() or len(text) < 30:
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    return doc

def enhance_with_ai(doc):
    try:
        full_text = "\n".join(para.text for para in doc.paragraphs)
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{
                "role": "system", 
                "content": "Improve this document professionally. Fix grammar, justify tables appropriately, and maintain original meaning."
            }, {
                "role": "user", 
                "content": full_text
            }],
            temperature=0.3
        )
        improved_text = response.choices[0].message.content
        for para in list(doc.paragraphs):
            para._element.getparent().remove(para._element)
        for line in improved_text.split('\n'):
            doc.add_paragraph(line)
    except Exception as e:
        print(f"AI Error: {e}")
    return doc

@app.route('/')
def index():
    return render_template('index.html')

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
        issues = analyze_document(doc)
        doc = justify_tables(doc)
        doc = enhance_with_ai(doc)
        
        processed_filename = f"enhanced_{filename}"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
        doc.save(processed_filepath)
        
        return render_template('results.html',
                            original=filename,
                            processed=processed_filename,
                            issues=issues,
                            file_info={
                                'name': filename,
                                'size': f"{os.path.getsize(filepath)/1024:.1f} KB"
                            })
    
    flash('Only .docx files allowed')
    return redirect(request.url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)