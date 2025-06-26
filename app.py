from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def beautify_document(input_path, output_path):
    doc = Document(input_path)

    standard_font = "Calibri"
    standard_size = Pt(12)

    for para in doc.paragraphs:
        para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in para.runs:
            run.font.name = standard_font
            run.font.size = standard_size
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), standard_font)
        if para.text.strip() == "":
            p = para._element
            p.getparent().remove(p)

    for table in doc.tables:
        table.style = 'Table Grid'
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    for run in para.runs:
                        run.font.name = standard_font
                        run.font.size = standard_size
                        r = run._element
                        r.rPr.rFonts.set(qn('w:eastAsia'), standard_font)

    doc.save(output_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file.filename.endswith('.docx'):
            filename = secure_filename(uploaded_file.filename)
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            output_path = os.path.join(OUTPUT_FOLDER, f\"beautified_{filename}\")
            uploaded_file.save(input_path)
            beautify_document(input_path, output_path)
            return send_file(output_path, as_attachment=True)

        return \"Please upload a .docx file only.\"

    return render_template('index.html')

if __name__== '__main__':
    from waitress import serve
    serve(app, host='0.0.0.0', port=10000)