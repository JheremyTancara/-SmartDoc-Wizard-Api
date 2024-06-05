from flask import Flask, send_file, request, jsonify
from flask_cors import CORS
from docx import Document
import os
import io

app = Flask(__name__)
CORS(app)

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    data = request.json
    replacement_text = data.get('replacement_text', 'first') if data is not None else 'first'
    value1_text = data.get('value1', 'second') if data is not None else 'second'
    value2_text = data.get('value2', 'third') if data is not None else 'second'
    value3_text = data.get('value3', 'fourth') if data is not None else 'third'
    value4_text = data.get('value4', 'fifth') if data is not None else 'fifth'

    if replacement_text.lower() == 'cartas':
        template_path = os.path.abspath('templates/template1.docx')
    elif replacement_text.lower() == 'informes':
        template_path = os.path.abspath('templates/template2.docx')
    else:
        return jsonify({"error": "Invalid document type"}), 400

    doc = Document(template_path)
    
    for paragraph in doc.paragraphs:
        if '{value1}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{value1}', value1_text)
        if '{value2}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{value2}', value2_text)
        if '{value3}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{value3}', value3_text)
        if '{value4}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{value4}', value4_text)
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    return send_file(doc_io, as_attachment=True, download_name='modified_template.docx')

if __name__ == '__main__':
    app.run(debug=True)
