import zipfile
import io
import pathlib
import flask as fl
import pandas as pd
import numpy as np
from itertools import chain
import os
from flask import Flask, Response

from pandas import ExcelWriter
import csv, xlrd
from datetime import date
from flask import Flask, flash, request, redirect, url_for, render_template, send_from_directory, make_response, \
    current_app, session
import encrypt
from werkzeug.utils import secure_filename
from requests_toolbelt import MultipartEncoder
from flask import send_file
from flask import session
import random
from selenium import webdriver
import threading

# import user

UPLOAD_FOLDER = os.path.dirname(os.path.realpath(__file__))
UPLOAD_FOLDER = UPLOAD_FOLDER+'\Excel_files'
ALLOWED_EXTENSIONS = set(['xlsx'])


def drive():
    import webbrowser
    import time
    time.sleep(5)
    url = 'http://127.0.0.1:5000/'
    webbrowser.open_new_tab(url)


app = Flask(__name__)
app.secret_key = "super secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = UPLOAD_FOLDER


def chainer(s):
    return list(chain.from_iterable(s.str.split(';')))


###############User Input Form################
@app.route('/cc', methods=['GET', 'POST'])
def hello():
    if request.method == 'GET':
        return render_template('checkbox.html')
    else:
        if "Add" in request.form:
            GDS_code = None
            Multi_code = None
            f = request.args.get('f')
            fd = request.args.get('fd')
            GDS_code = request.form.get("GDS_code")
            if GDS_code == "":
                GDS_code = 'N'
            Multi_code = request.form.get("Multi_code")
            if Multi_code == "":
                Multi_code = 'N'
            Sabre = request.form.get("Sabre")
            Worldspan = request.form.get("Worldspan")
            Amaedus = request.form.get("Amaedus")
            Galileo = request.form.get("Galileo")
            Web = request.form.get("Web")
            return redirect(
                url_for('userinput', GDS_code=GDS_code, Multi_code=Multi_code, Sabre=Sabre, Worldspan=Worldspan,
                        Amaedus=Amaedus, Galileo=Galileo, Web=Web, f=f, fd=fd))
        else:
            GDS_code = None
            Multi_code = None
            f = request.args.get('f')
            fd = request.args.get('fd')
            GDS_code = request.form["GDS_code"]
            if GDS_code == "":
                GDS_code = 'N'
            Multi_code = request.form["Multi_code"]
            if Multi_code == "":
                Multi_code = 'N'
            Sabre = request.form["Sabre"]
            Worldspan = request.form["Worldspan"]
            Amaedus = request.form["Amaedus"]
            Galileo = request.form["Galileo"]
            Web = request.form["Web"]
            return redirect(
                url_for('backend', GDS_code=GDS_code, Multi_code=Multi_code, Sabre=Sabre, Worldspan=Worldspan,
                        Amaedus=Amaedus, Galileo=Galileo, Web=Web, f=f, fd=fd))


#################Upload Workbook########################
@app.route('/file', methods=['GET', 'POST'])
def upload_file():
    fd_data = request.args.get('fd_data')
    if fd_data == 'no':
        fd = 0
        if request.method == 'GET':
            return render_template('upload.html')
        elif request.method == 'POST':
            data_files = request.files['file']
            data_files.save(os.path.join(app.config['UPLOAD_FOLDER'], "MAIN.xlsx"))
            f = pd.DataFrame({'Corp Acct#': "1602"}, index=[0])
            f = f.to_json(orient='split')
            return redirect(url_for('hello', f=f, fd=fd))
    else:
        fd = 1
        if request.method == 'GET':
            return render_template('upload2.html')
        elif request.method == 'POST':
            data_files = request.files['file']
            data_files2 = request.files['file2']
            data_files.save(os.path.join(app.config['UPLOAD_FOLDER'], "MAIN.xlsx"))
            data_files2.save(os.path.join(app.config['UPLOAD_FOLDER'], "FAIR.xlsx"))
            f = pd.DataFrame({'Corp Acct#': "160802"}, index=[0])
            f = f.to_json(orient='split')
            return redirect(url_for('hello', f=f, fd=fd))


@app.route('/', methods=['GET', 'POST'])
def fd():
    if request.method == 'POST':
        fd_data = request.form['fd']
        return redirect(url_for('upload_file', fd_data=fd_data))
    # full_filename = "C://Users//purushv//AppData//Local//Programs//Python//Python36//venv//Scripts//dist//flask_up//templates/bw_logo.jpg"
    return render_template('fd.html')


'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():

    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('hello'))
            #return redirect(url_for('uploaded_file',
                                    #filename=filename))
    return render_template('upload.html')
'''


##############Backend for Transformation##########################
@app.route('/cc/backend_in_process')
def backend():
    GDS_code = request.args.get('GDS_code')
    Multi_code = request.args.get('Multi_code')
    Sabre = request.args.get('Sabre')
    Worldspan = request.args.get('Worldspan')
    Amaedus = request.args.get('Amaedus')
    Galileo = request.args.get('Galileo')
    Web = request.args.get('Web')

    f = request.args.get('f')
    fd = request.args.get('fd')

    df_Main = pd.read_excel(UPLOAD_FOLDER + '//MAIN.xlsx', sheet_name='Sheet1')
    today = date.today().strftime('%m/%d/%Y')
    f = pd.read_json(f, orient='split')

    df_Main.loc[df_Main['CountryName'] != 'United States', 'BWIRateCode'] = ''
    df_Main.loc[df_Main['CountryName'] == 'United States', 'BWIRateCode'] = ''
    c = pd.DataFrame({'Corp Acct#': df_Main['CRSIdentifier'], 'Resort': df_Main['CRSHotelID'],
                      'BWI Rate Code': df_Main['BWIRateCode'], 'GDS Rate Codes': GDS_code,
                      'Multi Rate Code': Multi_code, 'Begin Date': today, 'End Date': '12/31/2099', 'Sabre': Sabre,
                      'Worldspan': Worldspan, 'Amaedus': Amaedus, 'Galileo': Galileo, 'Web': Web,
                      'RoomDescription': df_Main['RoomDescription']})
    c = c.drop_duplicates(subset=['Resort'], keep='first')
    c['Map to ROH'] = np.where(c['RoomDescription'] == 'ROH', 'Y', 'N')
    c = c.drop(columns=['RoomDescription'])
    f = pd.concat([f, c], sort=False)
    # f = f.dropna(axis=0, subset=['Resort'])
    f.to_csv(UPLOAD_FOLDER + "//User.csv", index=False)
    f = f.to_csv(None, index=False).encode()
    # d=f.to_json(orient='split')

    # encrypt.etl(d)
    return redirect(url_for('down', fd=fd))


@app.route('/cc/add')
def userinput():
    GDS_code = request.args.get('GDS_code')
    Multi_code = request.args.get('Multi_code')
    Sabre = request.args.get('Sabre')
    Worldspan = request.args.get('Worldspan')
    Amaedus = request.args.get('Amaedus')
    Galileo = request.args.get('Galileo')
    Web = request.args.get('Web')

    f = request.args.get('f')
    fd = request.args.get('fd')

    df_Main = pd.read_excel(UPLOAD_FOLDER + "//MAIN.xlsx", sheet_name='Sheet1')
    df_Main.loc[df_Main['CountryName'] != 'United States', 'BWIRateCode'] = ''
    df_Main.loc[df_Main['CountryName'] == 'United States', 'BWIRateCode'] = ''
    today = date.today().strftime('%m/%d/%Y')
    f = pd.read_json(f, orient='split')
    c = pd.DataFrame({'Corp Acct#': df_Main['CRSIdentifier'], 'Resort': df_Main['CRSHotelID'],
                      'BWI Rate Code': df_Main['BWIRateCode'], 'GDS Rate Codes': GDS_code,
                      'Multi Rate Code': Multi_code, 'Begin Date': today, 'End Date': '12/31/2099', 'Sabre': Sabre,
                      'Worldspan': Worldspan, 'Amaedus': Amaedus, 'Galileo': Galileo, 'Web': Web,
                      'RoomDescription': df_Main['RoomDescription']})
    c = c.drop_duplicates(subset=['Resort'], keep='first')
    c['Map to ROH'] = np.where(c['RoomDescription'] == 'ROH', 'Y', 'N')
    c = c.drop(columns=['RoomDescription'])
    f = pd.concat([f, c], sort=False)
    print(f)
    f = f.to_json(orient='split')

    # encrypt.etl(d)
    return redirect(url_for('hello', f=f, fd=fd))


################################################################################################################################

@app.route('/dd', methods=['GET', 'POST'])
def down():
    if request.method == 'POST':
        bwi = None
        fd = request.args.get('fd')
        rt = request.form['rt']
        ol = request.form['ol']
        lra = request.form['lra']
        ri = request.form['ri']
        bwi = request.form['bwi']
        return redirect(url_for('down1', fd=fd, rt=rt, ol=ol, lra=lra, ri=ri, bwi=bwi))
    return render_template('down.html')


r1 = random.randint(1, 100000)
r2 = random.randint(1, 100000)
r3 = random.randint(1, 100000)


@app.route('/dd1', methods=['GET', 'POST'])
def down1():
    bwi = None
    rt = request.args.get('rt')
    ol = request.args.get('ol')
    lra = request.args.get('lra')
    ri = request.args.get('ri')
    fd = request.args.get('fd')
    bwi = request.args.get('bwi')
    repeat = 0
    qa = []
    ww = []
    a = encrypt.etl(fd, rt, ol, lra, ri, bwi, repeat, qa, ww)

    if a == 0:
        return render_template('dcsv.html')
    elif a==8:
        return render_template('error_only.html')
    else:
        return render_template('errorcsv.html')


##########Download_the_CSV#############


@app.route('/' + str(r1) + '/' + str(r2) + '/' + str(r3) + '/success' + str(r3), methods=['GET', 'POST'])
def index():
    file = UPLOAD_FOLDER + "//pd.xlsx"
    return send_file(file, as_attachment=True)


@app.route('/' + str(r1) + '/' + str(r2) + '/' + str(r3) + '/error' + str(r3), methods=['GET', 'POST'])
def texte():
    file = UPLOAD_FOLDER + "//df.csv"
    return send_file(file, as_attachment=True)


@app.route('/' + str(r1) + '/' + str(r2) + '/' + str(r3) + '/audit' + str(r3), methods=['GET', 'POST'])
def audit():
    uploads = os.path.join(current_app.root_path, app.config['UPLOAD_FOLDER'])
    return send_from_directory(directory=uploads, filename="Audit_Rpt.xlsx", as_attachment=True)


if __name__ == '__main__':
    import threading

    thread2 = threading.Thread(target=app.run)
    thread2.start()
    drive()
    thread2.join()
    '''
    #drive()
    app.debug = True
    app.run(debug = True)

    '''

