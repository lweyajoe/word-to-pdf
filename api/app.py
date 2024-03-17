from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
from docx import Document
import pdfkit
from urllib.parse import quote

app = Flask(__name__)

# Configuration for PDFKit - path to wkhtmltopdf executable
pdfkit_config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')

@app.route('/')
def index():
    """Renders the index.html template when accessing the root URL."""
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    """Handles the file upload and conversion process."""
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file and allowed_file(file.filename):
        # Check if the file is a valid Word document
        if is_valid_word_file(file):
            # Save the uploaded Word file
            word_filename = secure_filename(file.filename)
            word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
            file.save(word_path)

            # Convert Word to PDF
            pdf_filename = os.path.splitext(word_filename)[0] + ".pdf"
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
            convert_to_pdf(word_path, pdf_path)

            # Return the converted PDF file
            return send_file(pdf_path, as_attachment=True)
        else:
            return "Invalid Word file"

    return "File type not allowed"

def is_valid_word_file(file):
    """Checks if the uploaded file is a valid Word document."""
    # Check if the file has a valid Word document extension
    valid_extensions = {'.docx', '.doc'}
    return '.' in file.filename and file.filename.rsplit('.', 1)[1].lower() in valid_extensions or \
           file.content_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or \
           file.content_type == 'application/vnd.openxmlformats-officedocument.themeManager+xml' or \
           file.content_type == 'application/msword'

def convert_to_pdf(input_path, output_path):
    """Converts a Word document to PDF."""
    # Read the Word document
    doc = Document(input_path)
    
    # Convert each paragraph to HTML
    html_content = ""
    for para in doc.paragraphs:
        html_content += "<p>{}</p>".format(para.text)
    
    # Convert HTML to PDF
    pdfkit.from_string(html_content, output_path, configuration=pdfkit_config)

def allowed_file(filename):
    """Checks if the filename has a valid extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'doc', 'docx'}

if __name__ == '__main__':
    app.config['UPLOAD_FOLDER'] = 'uploads'
    app.run(debug=False)
