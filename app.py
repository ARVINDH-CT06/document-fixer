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
import traceback

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

# Initialize OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def analyze_document_structure(doc):
    """Analyze document structure and identify issues"""
    issues = []
    
    # Check title formatting
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        if not first_para.text.strip().isupper() and len(first_para.text) < 50:
            issues.append("First paragraph should be formatted as Title")
    
    # Check headings
    for i, para in enumerate(doc.paragraphs[1:], start=2):
        if para.text.strip() and len(para.text) < 80 and not para.style.name.startswith('Heading'):
            issues.append(f"Paragraph {i} might be a heading but isn't formatted as one")
    
    # Check table alignments
    for table_idx, table in enumerate(doc.tables, start=1):
        for row in table.rows:
            for cell in row.cells:
                if cell.vertical_alignment != WD_ALIGN_VERTICAL.CENTER:
                    issues.append(f"Table {table_idx}: Cell vertical alignment should be centered")
    return issues

def fix_document_formatting(doc):
    """Apply professional formatting to the document"""
    # Ensure Title style exists
    try:
        title_style = doc.styles['Title']
    except KeyError:
        title_style = doc.styles.add_style('Title', WD_STYLE_TYPE.PARAGRAPH)
        title_style.font.name = 'Calibri Light'
        title_style.font.size = Pt(18)
        title_style.font.bold = True
        title_style.font.color.rgb = RGBColor(0x2B, 0x54, 0x8B)  # Navy blue
    
    # Apply title formatting
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        first_para.style = title_style
        first_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Standardize body text
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # Black
    
    # Perfect table formatting
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    return doc

def enhance_with_ai(doc):
    """Use AI to enhance document content"""
    try:
        full_text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())
        
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a professional editor. Improve grammar, clarity and professionalism while preserving all key information and original meaning."},
                {"role": "user", "content": full_text}
            ],
            temperature=0.3
        )
        
        improved_text = response.choices[0].message.content
        
        # Clear and rebuild document
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
        
        for line in improved_text.split('\n'):
            doc.add_paragraph(line.strip())
            
    except Exception as e:
        print(f"AI Enhancement Error: {str(e)}")
        traceback.print_exc()
    return doc

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
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
            issues = analyze_document_structure(doc)
            doc = fix_document_formatting(doc)
            doc = enhance_with_ai(doc)
            
            processed_filename = f"enhanced_{filename}"
            processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
            doc.save(processed_filepath)
            
            return render_template('results.html', 
                                original=filename,
                                processed=processed_filename,
                                issues=issues)
        
        flash('Invalid file type. Only .docx files allowed')
        return redirect(request.url)
    
    except Exception as e:
        flash(f'Error processing file: {str(e)}')
        return redirect(request.url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)