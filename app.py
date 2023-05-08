import os
from flask import Flask, request, redirect, url_for, send_from_directory, flash, render_template
from werkzeug.utils import secure_filename

from pdf2docx import pdf_to_png
from pdf2docx import add_images_to_docx

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'your_secret_key_here'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # Convert the PDF file to PNG images
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            pdf_to_png(pdf_path)
            # Convert the PNG images to a Word document
            image_path = os.path.splitext(pdf_path)[0]
            word_doc = os.path.splitext(pdf_path)[0] + ".docx"
            add_images_to_docx(image_path, word_doc)
            return redirect(url_for('download_file', filename=os.path.basename(word_doc)))
    return render_template('upload.html')


@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(os.path.dirname(os.path.join(UPLOAD_FOLDER, filename)), filename, as_attachment=True)


if __name__ == '__main__':
    app.run()
