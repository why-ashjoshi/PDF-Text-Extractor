from flask import Flask, request, render_template, send_file
import fitz  # PyMuPDF
import os
from docx import Document
from io import BytesIO

app = Flask(__name__)

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text.strip()  # Remove any leading/trailing whitespace

def save_text_as_docx(text, save_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(save_path)
    print(f"Document saved to {save_path}")  # Debug: Confirm file save

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'pdf' not in request.files:
            return 'No file part'
        
        file = request.files['pdf']
        if file.filename == '':
            return 'No selected file'
        
        if file and file.filename.endswith('.pdf'):
            pdf_path = 'uploaded_file.pdf'
            file.save(pdf_path)
            
            text = extract_text_from_pdf(pdf_path)
            
            # Save to .docx
            docx_path = 'extracted_text.docx'
            save_text_as_docx(text, docx_path)
            
            # Create an in-memory buffer
            buffer = BytesIO()
            with open(docx_path, 'rb') as f:
                buffer.write(f.read())
            buffer.seek(0)
            
            # Cleanup
            os.remove(pdf_path)
            os.remove(docx_path)
            
            return send_file(buffer, as_attachment=True, download_name='extracted_text.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
