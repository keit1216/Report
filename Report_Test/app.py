
from flask import Flask
UPLOAD_FOLDER = '/home/keit1216/keit1216_mlEngine/report/report_test'
app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['REPORT_NAME'] = 'new'