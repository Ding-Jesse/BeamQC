from multiprocessing import allow_connection_pickling
import os
from time import sleep
from urllib import response
from flask import Flask, request, redirect, url_for, render_template,send_from_directory,session,g, Response, stream_with_context,jsonify
from werkzeug.utils import secure_filename
from main import main_functionV3, main_col_function,storefile
import functools
import json
import time
from datetime import timedelta
from auth import createPhoneCode,sendPhoneMessage
from beam_count import count_beam_main
app = Flask(__name__)

# UPLOAD_FOLDER = 'C:/Users/User/Desktop/BeamQC/INPUT'
# OUTPUT_FOLDER = 'C:/Users/User/Desktop/BeamQC/OUTPUT'
ALLOWED_EXTENSIONS = set(['dwg','DWG'])
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
# app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
# app.config['SECRET_KEY'] = b'_5#y2L"F4Q8z\n\xec]/'
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('home.html')

def login_required(view):
    @functools.wraps(view)
    def wrapped_view(**kwargs):
        user_agree = session.get('user_agree')
        if user_agree is None:
            return redirect(url_for('login'))

        return view(**kwargs)

    return wrapped_view

@app.before_request
def before_request():
    session.permanent = True
    # app.permanent_session_lifetime = timedelta(minutes=15)

@app.route('/tool1', methods=['GET', 'POST'])
@login_required
def upload_file():
    if request.method == 'POST':
        
        uploaded_beams = request.files.getlist("file1")
        uploaded_plans = request.files.getlist("file2")
        uploaded_columns = request.files.getlist("file_col")
        beam_type = '大梁'
        sbeam_type = '小梁'
        beam_file =[]
        plan_file = []
        column_file = []
        txt_file =''
        dwg_type = 'single'
        project_name = request.form['project_name']
        text_col_layer = request.form['text_col_layer']
        line_layer = request.form['line_layer']
        text_layer = request.form['text_layer']
        block_layer = request.form['block_layer']
        floor_layer = request.form['floor_layer']
        big_beam_layer = request.form['big_beam_layer']
        big_beam_text_layer = request.form['big_beam_text_layer']
        sml_beam_layer = request.form['sml_beam_layer']
        sml_beam_text_layer = request.form['sml_beam_text_layer']
        size_layer = request.form['size_layer']
        col_layer = request.form['col_layer']

        xs_col = request.form.get('xs-col')
        xs_beam = request.form.get('xs-beam')
        sizing = request.form.get('sizing')
        mline_scaling = request.form.get('mline_scaling')

        beam_ok = False
        plan_ok = False
        column_ok = False
        filenames = ['']
        project_name = time.strftime("%Y-%m-%d-%H-%M", time.localtime())+project_name
        progress_file = f'./OUTPUT/{project_name}_progress'
        if len(uploaded_beams) > 1: dwg_type = 'muti'
        for uploaded_beam in uploaded_beams:
            if uploaded_beam and allowed_file(uploaded_beam.filename) and xs_beam:
                beam_ok, beam_new_file = storefile(uploaded_beam,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                filename_beam = secure_filename(uploaded_beam.filename)
                beam_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_beam}'))
        for uploaded_column in uploaded_columns:
            if uploaded_column and allowed_file(uploaded_column.filename) and xs_col:
                column_ok, column_new_file = storefile(uploaded_column,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                filename_column = secure_filename(uploaded_column.filename)
                column_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_column}'))
        for uploaded_plan in uploaded_plans:
            if uploaded_plan and allowed_file(uploaded_plan.filename):
                plan_ok, plan_new_file = storefile(uploaded_plan,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                filename_plan = secure_filename(uploaded_plan.filename)
                plan_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_plan}'))
                col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-column-{filename_plan}')
        if beam_ok and len(plan_file)==1:filenames.append(os.path.split(plan_new_file)[1])
        if len(beam_file)==1:filenames.append(os.path.split(beam_new_file)[1])
        if len(column_file)==1:filenames.append(os.path.split(column_new_file)[1])
        if column_ok and len(plan_file)==1:filenames.append(os.path.split(col_plan_new_file)[1])
            # filename_plan = secure_filename(uploaded_plan.filename)
            # plan_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_plan}')
            # plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-{filename_plan}')
            # col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_COL_MARKON-{filename_plan}')
            # uploaded_plan.save(plan_file)
            # plan_ok = True
        if beam_ok and plan_ok:
            # main function
            txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{beam_type}.txt')
            sb_txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{sbeam_type}.txt')
            main_functionV3(beam_file,plan_file,beam_new_file,plan_new_file,txt_file,sb_txt_file,text_layer,block_layer,floor_layer,size_layer,big_beam_layer,big_beam_text_layer,sml_beam_layer,sml_beam_text_layer,project_name,progress_file,sizing,mline_scaling)
            filenames_beam = [f'{project_name}-{beam_type}.txt',f'{project_name}-{sbeam_type}.txt']
            filenames.extend(filenames_beam) 
        if column_ok and plan_ok:
            # main function
            txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-column.txt')
            main_col_function(column_file,plan_file,column_new_file,col_plan_new_file,txt_file,text_col_layer,line_layer,block_layer,floor_layer,col_layer,project_name,progress_file)
            filenames_column = [f'{project_name}-column.txt']
            filenames.extend(filenames_column)
        if column_ok or beam_ok:
            if 'filenames' in session:
                session['filenames'].extend(filenames)
            else:
                session['filenames'] = filenames
            return render_template('tool1_result.html', filenames=filenames)
    return render_template('tool1.html')

@app.route('/results')
@login_required
def result_page():
    filenames = session.get('filenames',[])
    count_filenames = session.get('count_filenames',[])
    # if filenames is None or len(filenames)==0:
    #     return render_template('tool1_result.html', filenames=[])
    # else:
    return render_template('tool1_result.html', filenames=filenames,count_filenames = count_filenames)

@app.route('/results/<filename>/')
def result_file(filename):
    response = send_from_directory(app.config['OUTPUT_FOLDER'],
                               filename, as_attachment = True)
    response.cache_control.max_age = 0
    return response

@app.route('/tutorial')
def tutorial():
    return render_template('tutorial.html')

@app.route('/secret')
def secret():
    return render_template('secret.html')

@app.route('/NOT_FOUND')
def NOT_FOUND():
    return render_template('404.html')

@app.route('/tool5')
def tool5():
    return render_template('tool5.html')

@app.route('/sendVerifyCode',methods=['POST'])
def sendVerifyCode():
    if request.method == 'POST':
        content = request.form['phone']
        response = Response()
        # response.data = str('{"phone":'+content+'}').encode()
        # response.data = jsonify({'validate':})
        response.data = json.dumps({'validate':f'send text to {content}'})
        response.status_code = 200
        response.content_type = 'application/json'
        sendPhoneMessage(content)
        print(session["phoneVerifyCode"])
        return response

@app.route('/tool2', methods=['GET'])
def tool2():
    # return render_template('tool2.html')
    if 'isverify' not in session:
        return render_template('verifycode.html')
    elif session['isverify'] == 'expire':
        return render_template('verifycode.html')
    else:
        return render_template('tool2.html')

@app.route('/checkcode', methods=['POST'])
def checkcode():
    user_code = request.form.get('user_code')
    if 'phoneVerifyCode' not in session:
        response = Response()
        response.status_code = 404
        response.data = json.dumps({'validate':f'Wrong Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'expire'
        return response
    if user_code == session["phoneVerifyCode"]['code']:
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate':f'Correct Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'valid'
        return response
    else:
        response = Response()
        response.status_code = 404
        response.data = json.dumps({'validate':f'Wrong Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'expire'
        return response

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['user_agree'] = 'agree'
        return redirect(url_for('home'))
    return render_template('statement.html', template_folder='./')

@app.route('/count_beam',methods=['POST'])
def count_beam():
    uploaded_beams = request.files.getlist("file_beam")
    project_name = request.form['project_name']
    project_name = time.strftime("%Y-%m-%d-%H-%M", time.localtime())+project_name
    beam_filename = ''
    temp_file = ''
    rebar_file = f'{project_name}-數量.txt'
    tie_file = f'{project_name}-箍筋數量.txt'
    rebar_input_file = os.path.join(app.config['OUTPUT_FOLDER'],rebar_file)
    for uploaded_beam in uploaded_beams:
        beam_ok, beam_new_file = storefile(uploaded_beam,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],request.form['project_name'])
        beam_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{secure_filename(uploaded_beam.filename)}')
        temp_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-temp.pkl')
        print(f'beam_filename:{beam_filename},temp_file:{temp_file}')
    layer_config = {
        'rebar_data_layer':request.form['rebar_data_layer'], # 箭頭和鋼筋文字的塗層
        'rebar_layer':request.form['rebar_layer'], # 鋼筋和箍筋的線的塗層
        'tie_text_layer':request.form['tie_text_layer'], # 箍筋文字圖層
        'block_layer':request.form['block_layer'], # 框框的圖層
        'beam_text_layer' :request.form['beam_text_layer'], # 梁的字串圖層
        'bounding_block_layer':request.form['bounding_block_layer']
        }
    if beam_filename != '' and temp_file != '':
        # count_beam_main(beam_filename=beam_filename,layer_config=layer_config,temp_file=temp_file,rebar_file=rebar_input_file,tie_file=tie_file)
        if 'count_filenames' in session:
            session['count_filenames'].extend([rebar_file,tie_file])
        else:
            session['count_filenames'] = [rebar_file,tie_file]
    response = Response()
    
    response.status_code = 200
    response.data = json.dumps({'validate':f'Send Success'})
    response.content_type = 'application/json'
    # print(request.form['project_name'])
    time.sleep(1)
    return response

# @app.route("/listen/<project_name>/")
# def listen(project_name):

#   def respond_to_client():
#     while True:
#       f = open(f'./OUTPUT/{project_name}_progress', 'a+', encoding="utf-8") 
#       lines = f.readlines() #一行一行讀
#       color = 'white'
#       _data = json.dumps({"color":color, "counter":''.join(lines)}, ensure_ascii=False)
#       yield f"id: 1\ndata: {_data}\nevent: online\n\n"
#       time.sleep(5)
#       f.close
#   return Response(respond_to_client(), mimetype='text/event-stream')

@app.errorhandler(404)
def page_not_found(e):
    return redirect(url_for('NOT_FOUND'))

if __name__ == '__main__':
    app.config.from_object('config.config.DevConfig')
    # app.secret_key = 'dev'
    app.run(host = '192.168.0.143',debug=True,port=8080)