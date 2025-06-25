from flask import Flask, render_template, request, send_file
from docx import Document
import os
from io import BytesIO
from waitress import serve

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['file']
    if uploaded_file.filename.endswith('.docx'):
        doc = Document(uploaded_file)

        # Justify all paragraphs (keeps all original content)
        for para in doc.paragraphs:
            para.alignment = 3  # 3 = Justify

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='aligned_output.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    return "❌ Invalid file format. Please upload a .docx file only."

if __name__ == '__main__':
    serve(app, host='0.0.0.0', port=10000)