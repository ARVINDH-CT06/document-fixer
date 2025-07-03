from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import re
import openai
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

# Initialize OpenAI (make sure to set your API key in environment variables)
openai.api_key = os.getenv('OPENAI_API_KEY')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def analyze_document_structure(doc):
    """Analyze the document structure and identify issues"""
    issues = []
    
    # Check title formatting
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        if first_para.style.name != 'Title':
            issues.append("First paragraph should be formatted as Title")
    
    # Check headings
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip() and not para.style.name.startswith('Heading'):
            # Check if it might be a heading by text characteristics
            if len(para.text) < 50 and para.text.isupper():
                issues.append(f"Paragraph {i+1} might be a heading but isn't formatted as one")
    
    # Check table alignments
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.vertical_alignment != WD_ALIGN_VERTICAL.CENTER:
                    issues.append(f"Table cell vertical alignment should be centered")
                if cell.paragraphs[0].alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
                    issues.append(f"Table cell text alignment should be centered")
    
    return issues

def fix_document_formatting(doc):
    """Apply standard formatting fixes to the document"""
    
    # Ensure first paragraph is title
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        first_para.style = doc.styles['Title']
        first_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Standardize fonts
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
    
    # Fix table alignments
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    return doc

def enhance_with_ai(doc):
    """Use AI to improve document content and structure"""
    full_text = "\n".join([para.text for para in doc.paragraphs])
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a professional document editor. Improve the following document by correcting grammar, enhancing clarity, and ensuring professional tone while preserving all key information and meaning."},
                {"role": "user", "content": full_text}
            ],
            temperature=0.3
        )
        
        improved_text = response.choices[0].message.content
        
        # Clear existing content
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
        
        # Add improved content
        for line in improved_text.split('\n'):
            doc.add_paragraph(line)
            
    except Exception as e:
        print(f"AI enhancement failed: {e}")
    
    return doc

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Process the document
        doc = Document(filepath)
        
        # Analyze document
        issues = analyze_document_structure(doc)
        
        # Fix formatting
        doc = fix_document_formatting(doc)
        
        # Enhance with AI
        doc = enhance_with_ai(doc)
        
        # Save processed document
        processed_filename = f"processed_{filename}"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
        doc.save(processed_filepath)
        
        return render_template('results.html', 
                             original=filename,
                             processed=processed_filename,
                             issues=issues)
    
    return redirect(request.url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)