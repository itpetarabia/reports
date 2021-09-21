import os
from report import generate_report
from flask import Flask, request, redirect, url_for, render_template, send_from_directory, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__, static_url_path="/static")

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/upload/'
OUTPUT_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/output/'
EXCEL_TEMPLATE = os.path.dirname(os.path.abspath(__file__)) + '/Branch_Daily_Sales_Report_Sample.xlsx'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['EXCEL_TEMPLATE'] = EXCEL_TEMPLATE
ALLOWED_EXTENSIONS = {'csv', 'xlsx'}


DIR_PATH = os.path.dirname(os.path.realpath(__file__))
# limit upload size upto 8mb
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'data' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['data']
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            process_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), 'dsr.xlsx')
            return redirect(url_for('uploaded_file', filename='dsr.xlsx'))
    return render_template('index.html')


def process_file(inputcsv, filename):
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    generate_report(inputcsv, output_path)


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)