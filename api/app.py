from flask import Flask, render_template, request, send_file
import os
from docx import Document
import pdfkit

app = Flask(__name__)

# Configuration for PDFKit - path to wkhtmltopdf executable
pdfkit_config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')  # Path may vary, check your system

# Define the root route
@app.route('/')
def index():
    return render_template('index.html')  # Render the HTML template for the upload page

# Define the route for file conversion
@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part"  # Error if no file is uploaded

    file = request.files['file']

    if file.filename == '':
        return "No selected file"  # Error if no file is selected

    if file and allowed_file(file.filename):  # Check if the file type is allowed
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

    return "File type not allowed"  # Error if the file type is not allowed

# Function to convert Word document to PDF
def convert_to_pdf(input_path, output_path):
    # Read the Word document
    doc = Document(input_path)
    
    # Convert each paragraph to HTML
    html_content = ""
    for para in doc.paragraphs:
        html_content += "<p>{}</p>".format(para.text)
    
    # Convert HTML to PDF
    pdfkit.from_string(html_content, output_path, configuration=pdfkit_config)

# Function to check if the file type is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'doc', 'docx'}

# Main entry point of the application
if __name__ == '__main__':
    app.config['UPLOAD_FOLDER'] = 'uploads'  # Define the folder for uploaded files
    app.run(debug=True)  # Run the Flask application in debug mode
