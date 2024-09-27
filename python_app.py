from flask import Flask, request, render_template, send_file
import pandas as pd
from docx import Document
import os
import zipfile
from io import BytesIO
from docx.shared import RGBColor  # Add this import to handle font color

app = Flask(__name__)

# Define the route for the main page
@app.route('/')
def index():
    return render_template('index.html')

# Define the route for generating offer letters
@app.route('/generate', methods=['POST'])
def generate():
    names = request.form.get('names')
    names_list = names.splitlines()  # Split the pasted names by line

    # Path to the DOCX template
    template_file = 'offer_template.docx'

    # Create a BytesIO object to store the ZIP file in memory
    zip_buffer = BytesIO()

    # Create a temporary directory to save offer letters
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for name in names_list:
            # Skip empty lines
            if not name.strip():
                continue

            # Load the template document
            doc = Document(template_file)

            # Replace the placeholder with the actual name
            for paragraph in doc.paragraphs:
                if '{Name}' in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace('{Name}', name.strip())
                        run.font.color.rgb = RGBColor(0, 0, 0)  # Ensure text color is black

            # Save each offer letter in the temporary folder
            offer_letter_name = f"Offer_Letter_{name.strip()}.docx"
            doc_stream = BytesIO()
            doc.save(doc_stream)
            doc_stream.seek(0)

            # Write each offer letter to the zip file
            zip_file.writestr(offer_letter_name, doc_stream.getvalue())

    # Return the ZIP file for download
    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True, download_name='offer_letters.zip', mimetype='application/zip')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 8000)))
