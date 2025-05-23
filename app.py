import os
from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
from converter.controller import ConverterController

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.secret_key = 'secret_key'  
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER


os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        
        if 'file' not in request.files:
            flash('Please select a file to upload.')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('Did not select a file')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            
            try:
                controller = ConverterController(filepath, app.config['OUTPUT_FOLDER'])
                controller.convert_all_sheets()
                flash('Convert successful!')
                return redirect(url_for('download_files', folder_name=filename))
            except Exception as e:
                flash(f'Error: {e}')
                return redirect(request.url)

        else:
            flash('Only Excel files are allowed.')
            return redirect(request.url)

    return render_template('index.html')

@app.route('/downloads/<folder_name>')
def download_files(folder_name):

    output_files = []
    prefix = os.path.splitext(folder_name)[0]  # filename
    for f in os.listdir(app.config['OUTPUT_FOLDER']):
        if f.startswith(prefix) and f.endswith('.docx'):
            output_files.append(f)

    return render_template('downloads.html', files=output_files)

@app.route('/outputs/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)
