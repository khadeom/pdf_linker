from flask import Flask, render_template, request
import PyPDF2
import os

app = Flask(__name__)

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/display', methods=['POST'])
def display():
    word = request.form['word']
    pdf_path = request.form['pdf_path']

    if not os.path.exists(pdf_path):
        return "PDF file not found!"

    pdf_text = extract_text_from_pdf(pdf_path)

    return render_template('display.html', word=word, pdf_text=pdf_text)

if __name__ == '__main__':
    app.run(debug=True)
