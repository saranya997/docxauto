import process
import os
import pythoncom
import win32com.client
from flask_cors import CORS, cross_origin
from flask import Flask, flash, request, redirect, url_for, send_from_directory, Response, jsonify, json, send_file,session
from werkzeug.utils import secure_filename
UPLOAD_FOLDER = './uploads'
ALLOWED_EXTENSIONS = set(['docx','doc'])
filename=''

app = Flask(__name__)
app.secret_key = "super secret key"
CORS(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def convert():
    file_path= os.getcwd() + "\\uploads\\load.doc"
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.application")
    docx_file = '{0}{1}'.format(os.getcwd() + "\\uploads\\load.doc", 'x')
    wordDoc = word.Documents.Open(file_path, False, False, False)
    wordDoc.SaveAs2(docx_file, FileFormat = 16)
    wordDoc.Close()

def root_dir():  
    return os.path.abspath(os.path.dirname(__file__))

def get_file(filename): 
    try:
        src = os.path.join(root_dir(), filename)
        return open(src).read()
    except IOError as exc:
        return str(exc)
		

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

	

		   
@app.route('/search', methods=['GET', 'POST'])
def search():
	print(request.method)
	print(request.json['search'])
	status = process.searchdoc(request.json['search'])
	return jsonify(
        status=status
    )
		   
@app.route('/download', methods=['GET', 'POST'])
def download():
	return send_file('./output/output.docx', as_attachment=True)
		   
@app.route('/', methods=['GET', 'POST'])
def index():
	content = get_file('index.html')
	return Response(content, mimetype="text/html")
	
@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = "load.doc"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            convert()
            return "uploaded successfully"