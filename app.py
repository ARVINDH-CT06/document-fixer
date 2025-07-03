from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
import openai
from collections import defaultdict
import re

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

openai.api_key = os.getenv('OPENAI_API_KEY')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def detect_document_type(doc):
    """Identify Indian academic document types"""
    first_text = "\n".join(p.text for p in doc.paragraphs[:5]).lower()
    
    doc_types = {
        'syllabus': ['syllabus', 'course plan', 'curriculum'],
        'lab_record': ['experiment', 'apparatus', 'observation'],
        'assignment': ['assignment', 'homework', 'question bank'],
        'project_report': ['project', 'abstract', 'implementation'],
        'internship_report': ['internship', 'organization', 'learning outcomes']
    }
    
    for doc_type, keywords in doc_types.items():
        if any(keyword in first_text for keyword in keywords):
            return doc_type
    return 'general'

def analyze_academic_content(doc, doc_type):
    """Specialized analysis for Indian academic docs"""
    findings = defaultdict(list)
    
    # Check essential sections
    required_sections = {
        'syllabus': ['course objectives', 'outcomes', 'textbooks', 'references'],
        'lab_record': ['aim', 'procedure', 'result', 'viva questions'],
        'assignment': ['questions', 'solutions', 'diagrams'],
        'project_report': ['introduction', 'methodology', 'results', 'references']
    }
    
    full_text = "\n".join(p.text for p in doc.paragraphs).lower()
    for section in required_sections.get(doc_type, []):
        if section not in full_text:
            findings['missing_sections'].append(section)
    
    # Check numbering consistency
    if doc_type == 'assignment':
        if not re.search(r'q\d+\.|question\s+\d+', full_text):
            findings['format_issues'].append("Questions not properly numbered")
    
    return findings

def enhance_with_ai(doc, doc_type):
    """AI trained for Indian academic standards"""
    academic_prompts = {
        'syllabus': "You are an expert university professor. Improve this syllabus with clear outcomes, proper sectioning and academic rigor:",
        'lab_record': "You are an engineering lab instructor. Format this lab record with proper aim, apparatus, procedure, results and viva questions:",
        'assignment': "You are a college professor. Enhance this assignment with clear questions, proper numbering and model solutions:",
        'project_report': "You are a PhD guide. Improve this project report with academic writing standards, proper sections and technical accuracy:"
    }
    
    prompt = academic_prompts.get(doc_type, 
        "Improve this academic document with proper formatting, grammar and clarity while preserving all key information:")
    
    try:
        full_text = "\n".join(p.text for p in doc.paragraphs)
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": full_text}
            ],
            temperature=0.2  # Lower for academic precision
        )
        
        improved_text = response.choices[0].message.content
        
        # Clear and rebuild document while preserving original formatting
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
        
        for line in improved_text.split('\n'):
            if line.strip():
                doc.add_paragraph(line.strip())
                
    except Exception as e:
        print(f"AI Enhancement Error: {str(e)}")
    
    return doc

def format_tables(doc):
    """Indian academic table standards"""
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                text = cell.text.lower().strip()
                
                # Indian academic table formatting rules
                if any(x in text for x in ['s.no', 'roll no', 'code']):
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                elif text.replace('.','').isdigit():  # Numeric data
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                elif len(text) < 30:  # Short text
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:  # Long text
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
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
        doc_type = detect_document_type(doc)
        analysis = analyze_academic_content(doc, doc_type)
        doc = enhance_with_ai(doc, doc_type)
        doc = format_tables(doc)
        
        processed_filename = f"enhanced_{filename}"
        processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
        doc.save(processed_filepath)
        
        return render_template('results.html',
                            original=filename,
                            processed=processed_filename,
                            doc_type=doc_type.replace('_',' ').title(),
                            analysis=analysis,
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