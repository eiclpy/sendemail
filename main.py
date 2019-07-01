# coding:utf-8
import os
import sys

from flask import Flask, render_template, request
from flask_dropzone import Dropzone
from flask_httpauth import HTTPBasicAuth
import time
import multiprocessing
from gevent.pywsgi import WSGIServer
from config import listen, port, users

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
auth = HTTPBasicAuth()


@auth.get_password
def get_password(username):
    for user in users:
        if user['username'] == username:
            return user['password']
    return None


app.config.update(
    UPLOADED_PATH=os.path.join(basedir, 'uploads'),
    # Flask-Dropzone config:
    DROPZONE_ALLOWED_FILE_CUSTOM=True,
    DROPZONE_ALLOWED_FILE_TYPE='.zip',
    DROPZONE_MAX_FILE_SIZE=32,
    DROPZONE_MAX_FILES=1,
    DROPZONE_REDIRECT_VIEW='completed',
)

dropzone = Dropzone(app)
pros = None
last_ts = ''


def dosend(fname):
    ret = str(os.system('python sender.py '+fname))
    with open('retVal', 'w') as f:
        f.write(ret)
        f.flush()
    sys.stdout.write('return : %s\n' % ret)


@app.route('/', methods=['GET'])
@auth.login_required
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@auth.login_required
def upload():
    global pros
    global last_ts
    if pros is not None and pros.is_alive():
        return 'email is sending', 404
    f = request.files.get('file')
    fname = time.strftime('%Y-%m-%d-%H-%M-%S')
    last_ts = fname
    # fname = os.path.join(app.config['UPLOADED_PATH'], fname)
    f.save(os.path.join(app.config['UPLOADED_PATH'], fname+'.zip'))
    pros = multiprocessing.Process(target=dosend, args=(fname,))
    pros.start()

    return render_template('index.html')


@app.route('/completed')
@auth.login_required
def completed():
    global pros
    global last_ts
    retVal = 0
    loginfo = ''
    if pros is not None and pros.is_alive():
        stat = 'Sending'
    else:
        stat = 'Finish'
        with open('retVal', 'r') as f:
            retVal = f.read()
    return render_template('query.html', sendStatus=stat, errorcode=retVal, lines=loginfo.split('\n'))

@app.route('/getlog')
@auth.login_required
def getlog():
    global pros
    global last_ts
    retVal = 0
    loginfo = ''
    if pros is not None and pros.is_alive():
        stat = 'Sending'
    else:
        stat = 'Finish'
        with open('retVal', 'r') as f:
            retVal = f.read()
        logfile = './logs/mail_{}.log'.format(last_ts)
        if os.path.exists(logfile):
            with open(logfile, 'r', encoding='utf8') as f:
                loginfo = f.read()
            # loginfo = loginfo.replace('\n', '<br/>')
    return render_template('query.html', sendStatus=stat, errorcode=retVal, lines=loginfo.split('\n'))


if __name__ == '__main__':
    if not os.path.exists('logs'):
        os.makedirs('logs')
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    
    http_server = WSGIServer((listen, port), app)
    http_server.serve_forever()
