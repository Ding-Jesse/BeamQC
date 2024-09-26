import functools
import json
import os
import sys
import time
import traceback
import uuid

from flask import (Flask, Response, redirect, render_template, request,
                   send_from_directory, session, stream_with_context, url_for, jsonify)
from flask_mail import Mail, Message

from auth import sendPhoneMessage
from flask_session import Session
from src.beam_count import (count_beam_multiprocessing, count_beam_multifiles,
                            read_parameter_json)
from src.column_count import count_column_multiprocessing, count_column_multifiles
from src.joint_scan import joint_scan_main
from src.main import (Upload_Error_log, main_col_function, main_functionV3,
                      storefile)

app = Flask(__name__)
app.config.from_object('config.config.Config')
Session(app)
mail = Mail(app)
connected_clients: dict[str, list] = {}
ALLOWED_EXTENSIONS = set(['dwg', 'DWG'])


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


def write_result_log(file_path, result_content: dict):
    localtime = time.asctime(time.localtime(time.time()))
    with open(file_path, 'a', encoding='utf-8') as result_log:
        result_log.write(f'\n Time:{localtime} \n')
        for topic, content in result_content.items():
            result_log.write(f'{topic}:{content} \n')


@app.route('/')
def home():
    return render_template('home.html')


@app.route('/test_url', methods=['POST'])
def test_url():
    try:
        Upload_Error_log({
            'temp': 'test'
        })
        print('success')
    except Exception as ex:
        print(ex)


def login_required(view):
    @functools.wraps(view)
    def wrapped_view(**kwargs):
        user_agree = session.get('user_agree')
        client_id = session.get('client_id')
        if client_id is None:
            client_id = str(uuid.uuid4())
            print(f'set client id:{client_id}')
            session['client_id'] = client_id
            connected_clients[client_id] = []
        if user_agree is None:
            return redirect(url_for('login'))

        return view(**kwargs)

    return wrapped_view


@app.before_request
def before_request():
    if request.path == '/compare_beam':
        if 'last_post_time' not in session:
            session['last_post_time'] = time.time()
            # print(time.time())
        else:
            if time.time() - session['last_post_time'] < 60:
                print(f'last_post_time exist {session["last_post_time"]}')
                time.sleep(60)
                # session['post_status'] = 'disable'
                # return redirect(url_for('tool1'))
            else:
                session['last_post_time'] = time.time()
    # print(app.config['SESSION_PERMANENT'])
    # print(app.config['PERMANENT_SESSION_LIFETIME'])
    # session.permanent = True
    # app.permanent_session_lifetime = timedelta(minutes=15)


@app.route('/tool1', methods=['GET'])
@login_required
def tool1():
    # print(session.get('filenames',[]))
    return render_template('tool1.html')


def send_error_response(warning_message: str):
    response = Response()
    response.status_code = 200
    response.data = json.dumps({'validate': f'{warning_message}'})
    response.content_type = 'application/json'
    return response


@app.route('/compare_beam', methods=['POST'])
# @login_required
def upload_file():
    if request.method == 'POST':
        try:
            status = 'progress'
            uploaded_beams = request.files.getlist("file1")
            uploaded_plans = request.files.getlist("file2")
            uploaded_columns = request.files.getlist("file_col")
            email_address = request.form['email_address']
            beam_file = []
            plan_file = []
            column_file = []

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
            table_line_layer = request.form['table_line_layer']
            column_block_layer = request.form['column_block_layer']

            layer_config = {
                'text_layer': text_layer.split('\r\n'),
                'block_layer': block_layer.split('\r\n'),
                'floor_layer': floor_layer.split('\r\n'),
                'big_beam_layer': big_beam_layer.split('\r\n'),
                'big_beam_text_layer': big_beam_text_layer.split('\r\n'),
                'sml_beam_layer': sml_beam_layer.split('\r\n'),
                'size_layer': size_layer.split('\r\n'),
                'sml_beam_text_layer': sml_beam_text_layer.split('\r\n'),
                'line_layer': table_line_layer.split('\r\n'),
            }
            col_layer_config = {
                'text_layer': text_col_layer.split('\r\n'),
                'line_layer': line_layer.split('\r\n'),
                'block_layer': block_layer.split('\r\n'),
                'floor_layer': floor_layer.split('\r\n'),
                'col_layer': col_layer.split('\r\n'),
                'size_layer': size_layer.split('\r\n'),
                'table_line_layer': table_line_layer.split('\r\n'),
                'column_block_layer': column_block_layer.split('\r\n')
            }

            xs_col = request.form.get('xs-col')
            xs_beam = request.form.get('xs-beam')
            sizing = request.form.get('sizing')
            mline_scaling = request.form.get('mline_scaling')

            beam_ok = False
            plan_ok = False
            column_ok = False
            filenames = []
            project_name = f'{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}_{project_name}'
            start_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

            client_id = session.get('client_id', None)
            if client_id:
                if client_id not in connected_clients:
                    connected_clients[client_id] = []
                connected_clients[client_id].append(project_name)

            for uploaded_beam in uploaded_beams:
                if uploaded_beam and allowed_file(uploaded_beam.filename) and xs_beam:
                    beam_ok, beam_new_file, input_beam_file = storefile(
                        uploaded_beam, app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], project_name)

                    beam_file.append(input_beam_file)
            for uploaded_column in uploaded_columns:
                if uploaded_column and allowed_file(uploaded_column.filename) and xs_col:
                    column_ok, column_new_file, input_column_file = storefile(
                        uploaded_column, app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], project_name)

                    column_file.append(input_column_file)
            for uploaded_plan in uploaded_plans:
                if uploaded_plan and allowed_file(uploaded_plan.filename):
                    plan_ok, plan_new_file, input_plan_file = storefile(
                        uploaded_plan, app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], project_name)

                    plan_file.append(input_plan_file)
                    col_plan_new_file = f'{os.path.splitext(plan_new_file)[0]}_column.dwg'

            if beam_ok and len(plan_file) == 1:
                filenames.append(os.path.split(plan_new_file)[1])
            if len(beam_file) == 1:
                filenames.append(os.path.split(beam_new_file)[1])
            if len(column_file) == 1:
                filenames.append(os.path.split(column_new_file)[1])
            if column_ok and len(plan_file) == 1:
                filenames.append(os.path.split(col_plan_new_file)[1])

            if beam_ok and plan_ok:

                output_file = main_functionV3(beam_filenames=beam_file,
                                              plan_filenames=plan_file,
                                              beam_new_filename=beam_new_file,
                                              plan_new_filename=plan_new_file,
                                              output_directory=app.config['OUTPUT_FOLDER'],
                                              project_name=project_name,
                                              layer_config=layer_config,
                                              sizing=sizing,
                                              mline_scaling=mline_scaling,
                                              client_id=client_id)
                filenames_beam = [output_file]
                filenames.extend(filenames_beam)
            if column_ok and plan_ok:
                # main function
                output_file = main_col_function(col_filenames=column_file,
                                                plan_filenames=plan_file,
                                                col_new_filename=column_new_file,
                                                plan_new_filename=col_plan_new_file,
                                                output_directory=app.config['OUTPUT_FOLDER'],
                                                project_name=project_name,
                                                layer_config=col_layer_config,
                                                client_id=client_id)
                filenames_column = [output_file]
                filenames.extend(filenames_column)
            if column_ok or beam_ok:
                if 'filenames' in session:
                    session['filenames'].extend(filenames)
                else:
                    session['filenames'] = filenames
            if (email_address):
                try:
                    print(f'send_email:{email_address}, filenames:{filenames}')
                    sendResult(email_address, filenames, "配筋圖核對結果")
                except Exception as e:
                    print(e)
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate': '完成，請至輸出結果查看'})
            response.content_type = 'application/json'
            time.sleep(1)
            status = 'success'
        except ConnectionRefusedError:
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate': '發送請求過於頻繁，請稍等'})
            response.content_type = 'application/json'
            status = 'error'
        except Exception:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            detailed_traceback = traceback.extract_tb(exc_traceback)
            error_detail = {}
            for entry in detailed_traceback:
                error_detail.update({entry.filename: {
                    "Line": entry.lineno,
                    "Function": entry.name,
                    "Code Context": entry.line,
                    "Error Type": exc_type,
                    "Value": exc_value
                }})
            error_detail.update({
                'date': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                'project name': project_name,
            })
            Upload_Error_log(data=error_detail,
                             collection_name="Error Log")

            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate': '發生錯誤'})
            response.content_type = 'application/json'

        end_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        Upload_Error_log(data={
            'email address': email_address,
            'start_time': start_time,
            'end_time': end_time,
            'project name': project_name,
            'status': status
        }, collection_name="App Log")
        connected_clients[client_id].remove(project_name)
        return response
    return 400
    # return render_template('tool1_result.html', filenames=filenames)
    # return render_template('tool1.html')


@app.route('/results')
@login_required
def result_page():
    filenames = session.get('filenames', [])
    count_filenames = session.get('count_filenames', [])
    # print(session.get('filenames',[]))
    # if filenames is None or len(filenames)==0:
    #     return render_template('tool1_result.html', filenames=[])
    # else:
    return render_template('tool1_result.html', filenames=filenames, count_filenames=count_filenames)


@app.route('/results/<filename>/', methods=['GET', 'POST'])
def result_file(filename):
    if (not filename in session.get('filenames', []) and not filename in session.get('count_filenames', [])):
        return redirect('/')
    response = send_from_directory(app.config['OUTPUT_FOLDER'],
                                   filename, as_attachment=True)
    response.cache_control.max_age = 0
    return response


@app.route('/demo/<filename>/', methods=['GET', 'POST'])
def demo_file(filename):
    response = send_from_directory(app.config['DEMO_FOLDER'],
                                   filename, as_attachment=True)
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


@app.route('/joint_scan', methods=['GET', 'POST'])
@login_required
def tool5():
    if request.method == 'POST':
        try:

            beam_pkl = request.form.get('beam_pkl_file')
            column_pkl = request.form.get('column_pkl_file')

            beam_pkl = os.path.join(
                session['pkl_file_folder'], f'{beam_pkl}.pkl')
            column_pkl = os.path.join(
                session['pkl_file_folder'], f'{column_pkl}.pkl')

            uploaded_plans = request.files['file_plan']
            project_name = request.form['project_name']

            uploaded_xlsx = request.files['file_floor_xlsx']

            if uploaded_plans:
                xlsx_ok, xlsx_new_file, input_plan_file = storefile(uploaded_plans, app.config['UPLOAD_FOLDER'],
                                                                    app.config['OUTPUT_FOLDER'],
                                                                    f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
                plan_filename = input_plan_file
            if uploaded_xlsx:
                xlsx_ok, xlsx_new_file, input_xlsx_file = storefile(
                    uploaded_xlsx, app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
                # xlsx_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_xlsx.filename)}')
                xlsx_filename = input_xlsx_file

            layer_config = {
                'block_layer': request.form['block_layer'].split('\r\n'),
                'floor_text_layer': request.form['floor_text_layer'].split('\r\n'),
                'beam_name_text_layer': request.form['beam_name_text_layer'].split('\r\n'),
                # 框框的圖層
                'column_name_text_layer': request.form['column_name_text_layer'].split('\r\n'),
                'beam_mline_layer': request.form['beam_mline_layer'].split('\r\n'),
                'column_block_layer': request.form['column_block_layer'].split('\r\n'),
            }
            output_floor = request.form['docx_output_floor'].split('\r\n')
            output_serial = request.form['docx_output_serial'].split('\r\n')
            new_plan_view, excel_filename, docx_filename = joint_scan_main(plan_filename=plan_filename,
                                                                           layer_config=layer_config,
                                                                           output_folder=app.config['OUTPUT_FOLDER'],
                                                                           project_name=project_name,
                                                                           beam_pkl=beam_pkl,
                                                                           column_pkl=column_pkl,
                                                                           column_beam_joint_xlsx=xlsx_filename,
                                                                           client_id=session.get(
                                                                               'client_id'),
                                                                           output_floor=output_floor,
                                                                           output_serial=output_serial)
            if 'count_filenames' in session:
                session['count_filenames'].extend(
                    [new_plan_view, excel_filename, docx_filename])
            else:
                session['count_filenames'] = [
                    new_plan_view, excel_filename, docx_filename]
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate': f'計算完成，請至輸出結果查看'})
            response.content_type = 'application/json'
        except Exception:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            detailed_traceback = traceback.extract_tb(exc_traceback)
            error_detail = {}
            for entry in detailed_traceback:
                error_detail.update({entry.filename: {
                    "Line": entry.lineno,
                    "Function": entry.name,
                    "Code Context": entry.line
                }})
            Upload_Error_log(data=error_detail,
                             collection_name="Error Log")
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate': f'發生錯誤'})
            response.content_type = 'application/json'
        return response
        # Process the selected option as needed
        # return jsonify({'status': 'success', 'selectedOption': selected_option})

    beam_pkl_files = [path for path in session['beam_pkl_files']
                      ] if 'beam_pkl_files' in session else []
    column_pkl_files = [path for path in session['column_pkl_files']
                        ] if 'column_pkl_files' in session else []
    return render_template('tool5.html', beam_pkl_files=beam_pkl_files, column_pkl_files=column_pkl_files)


@app.route('/sendVerifyCode', methods=['POST'])
def sendVerifyCode():
    if request.method == 'POST':
        content = request.form['phone']
        response = Response()
        # response.data = str('{"phone":'+content+'}').encode()
        # response.data = jsonify({'validate':})
        response.data = json.dumps({'validate': f'send text to {content}'})
        response.status_code = 200
        response.content_type = 'application/json'
        sendPhoneMessage(content)
        print(session["phoneVerifyCode"])
        return response


@app.route('/tool2', methods=['GET', 'POST'])
@login_required
def tool2():
    if app.config['TESTING']:
        # return render_template('verifycode.html')
        return render_template('tool2.html')
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
        response.data = json.dumps({'validate': f'Wrong Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'expire'
        return response
    if user_code == session["phoneVerifyCode"]['code']:
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate': f'Correct Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'valid'
        return response
    else:
        response = Response()
        response.status_code = 404
        response.data = json.dumps({'validate': f'Wrong Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'expire'
        return response


@app.route('/admin_login', methods=['POST'])
def admin_login():
    user_code = request.form.get('user_code')
    if user_code == "wp32s%v9jhh!n+5i":
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate': f'Correct Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'valid'
        return response
    else:
        response = Response()
        response.status_code = 404
        response.data = json.dumps({'validate': f'Wrong Code'})
        response.content_type = 'application/json'
        session['isverify'] = 'expire'
        return response


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['user_agree'] = 'agree'

        return redirect(url_for('home'))
    return render_template('statement.html', template_folder='./')


@app.route('/count_beam', methods=['POST'])
def count_beam():
    result_log_content = {}
    try:
        beam_filename = ''
        plan_filename = ''
        beam_filenames = []
        uploaded_plans = request.files['file_plan']
        uploaded_xlsx = request.files['file_floor_xlsx']
        uploaded_beams = request.files.getlist("file_beam")
        project_name = request.form['project_name']
        email_address = request.form['email_address']

        template_name = request.form['company']
        measure_type = request.form['measure_type']

        beam_filename = ''
        temp_file = ''
        client_id = session.get('client_id', None)
        if client_id:
            if client_id not in connected_clients:
                connected_clients[client_id] = []
            connected_clients[client_id].append(project_name)
        # rebar_input_file = os.path.join(app.config['OUTPUT_FOLDER'],rebar_file)
        response = Response()

        if len(uploaded_beams) == 0:
            response.status_code = 404
            response.data = json.dumps({'validate': '未上傳檔案'})
            response.content_type = 'application/json'
            return response
        print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} start {project_name}')
        for uploaded_beam in uploaded_beams:
            beam_ok, \
                beam_new_file, \
                input_beam_file = storefile(uploaded_beam,
                                            app.config['UPLOAD_FOLDER'],
                                            app.config['OUTPUT_FOLDER'],
                                            f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # beam_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_beam.filename)}')
            beam_filename = input_beam_file
            beam_filenames.append(beam_filename)
            # temp_file = os.path.join(app.config['UPLOAD_FOLDER'],
            #                          f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-temp.pkl')
            pkl_file_folder = os.path.join(app.config['UPLOAD_FOLDER'])
            # print(f'beam_filename:{beam_filename},temp_file:{temp_file}')
        if uploaded_xlsx:
            xlsx_ok, \
                xlsx_new_file, \
                input_xlsx_file = storefile(uploaded_xlsx,
                                            app.config['UPLOAD_FOLDER'],
                                            app.config['OUTPUT_FOLDER'],
                                            f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # xlsx_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_xlsx.filename)}')
            xlsx_filename = input_xlsx_file
        if uploaded_plans:
            xlsx_ok, xlsx_new_file, input_plan_file = storefile(uploaded_plans,
                                                                app.config['UPLOAD_FOLDER'],
                                                                app.config['OUTPUT_FOLDER'],
                                                                f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            plan_filename = input_plan_file
            # print(f'xlsx_filename:{xlsx_filename}')
        layer_config = {
            # 箭頭和鋼筋文字的塗層
            'rebar_data_layer': request.form['rebar_data_layer'].split('\r\n'),
            # 鋼筋和箍筋的線的塗層
            'rebar_layer': request.form['rebar_layer'].split('\r\n'),
            # 箍筋文字圖層
            'tie_text_layer': request.form['tie_text_layer'].split('\r\n'),
            'block_layer': request.form['block_layer'].split('\r\n'),  # 框框的圖層
            # 梁的字串圖層
            'beam_text_layer': request.form['beam_text_layer'].split('\r\n'),
            'bounding_block_layer': request.form['bounding_block_layer'].split('\r\n'),
            'rc_block_layer': request.form['rc_block_layer'].split('\r\n'),
            's_dim_layer': request.form['s_dim_layer'].split('\r\n'),
            'burst_layer_list': request.form['burst_layer_list'].split('\r\n'),
        }
        plan_layer_config = {
            'block_layer':  request.form['plan_block_layer'].split('\r\n'),
            'floor_text_layer':  request.form['plan_floor_text_layer'].split('\r\n'),
            'name_text_layer':  request.form['plan_serial_text_layer'].split('\r\n'),
        }
        result_log_content['upload_xlsx'] = uploaded_xlsx
        result_log_content['upload_beams'] = uploaded_beams
        result_log_content['project_name'] = project_name
        result_log_content['email_address'] = email_address
        result_log_content['template_name'] = template_name
        result_log_content['layer_config'] = layer_config

        # print(layer_config)
        if beam_filename != '' and beam_ok:
            # rebar_txt,rebar_txt_floor,rebar_excel,rebar_dwg =count_beam_main(beam_filename=beam_filename,layer_config=layer_config,temp_file=temp_file,
            #                                                                     output_folder=app.config['OUTPUT_FOLDER'],project_name=project_name,template_name=template_name)
            # output_file_list, output_dwg_list, pkl = count_beam_multiprocessing(beam_filenames=beam_filenames,
            #                                                                     layer_config=layer_config,
            #                                                                     pkl_file_folder=pkl_file_folder,
            #                                                                     project_name=project_name,
            #                                                                     output_folder=app.config['OUTPUT_FOLDER'],
            #                                                                     template_name=template_name,
            #                                                                     floor_parameter_xlsx=xlsx_filename,
            #                                                                     progress_file=progress_file,
            #                                                                     plan_filename=plan_filename,
            #                                                                     plan_layer_config=plan_layer_config,
            #                                                                     client_id=client_id)
            hyper_parameter: dict = read_parameter_json(
                template_name=template_name)['beam']

            hyper_parameter.update({'layer_config': layer_config,
                                    'plan_layer_config': plan_layer_config,
                                    'measure_type': measure_type})

            output_file_list, output_dwg_list, pkl = count_beam_multifiles(beam_filenames=beam_filenames,
                                                                           pkl_file_folder=pkl_file_folder,
                                                                           project_name=project_name,
                                                                           output_folder=app.config['OUTPUT_FOLDER'],
                                                                           floor_parameter_xlsx=xlsx_filename,
                                                                           plan_filename=plan_filename,
                                                                           client_id=client_id,
                                                                           **hyper_parameter)
            # output_dwg_list = ['P2022-06A 岡山大鵬九村社宅12FB2_20230410_170229_Markon.dwg']
            if 'count_filenames' in session:
                session['count_filenames'].extend(output_file_list)
                session['count_filenames'].extend(output_dwg_list)
            else:
                session['count_filenames'] = output_file_list
                session['count_filenames'].extend(output_dwg_list)

            pkl = os.path.splitext(os.path.basename(pkl))[0]
            session['pkl_file_folder'] = pkl_file_folder

            if 'beam_pkl_files' in session:
                session['beam_pkl_files'].extend(pkl)
            else:
                session['beam_pkl_files'] = [pkl]
        if (email_address):
            try:
                sendResult(email_address, output_file_list, "梁配筋圖數量計算結果")
                sendResult(email_address, output_dwg_list, "梁配筋圖數量計算結果")
                print(
                    f'send_email:{email_address}, filenames:{session["count_filenames"]}')
            except Exception:
                pass
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate': f'計算完成，請至輸出結果查看'})
        response.content_type = 'application/json'
        result_log_content['status'] = 'success'
        # print(request.form['project_name'])
        time.sleep(1)
    except Exception as ex:
        import sys
        import traceback
        error_class = ex.__class__.__name__  # 取得錯誤類型
        detail = ex.args[0]  # 取得詳細內容
        cl, exc, tb = sys.exc_info()  # 取得Call Stack
        for lastCallStack in traceback.extract_tb(tb):
            fileName = lastCallStack[0]  # 取得發生的檔案名稱
            lineNum = lastCallStack[1]  # 取得發生的行號
            funcName = lastCallStack[2]  # 取得發生的函數名稱
            className = sys.exc_info()[0]
            errMsg = "F[{}] {} ".format(error_class, detail)  # 新增常見使用者錯誤
            adminMsg = "File \"{}\", line {}, in {}: [{}] {} ".format(
                fileName, lineNum, funcName, error_class, detail)
            print(adminMsg)
        result_log_content['status'] = f'error, {ex}'
        # print(ex)
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate': f'發生錯誤'})
        response.content_type = 'application/json'
    print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} end {project_name}')
    write_result_log(file_path=r'result\result_log.txt',
                     result_content=result_log_content)
    return response


@app.route('/count_column', methods=['POST'])
def count_column():
    result_log_content = {}
    try:
        uploaded_columns = request.files.getlist("file_column")
        project_name = request.form['project_name']
        email_address = request.form['email_address']
        template_name = request.form['companyColumn']
        measure_type = request.form['measure_type']
        uploaded_xlsx = request.files['file_floor_xlsx']
        uploaded_plans = request.files['file_plan']

        column_filename = ''
        plan_filename = ''
        column_filenames = []
        column_excel = ''
        column_ok = False

        response = Response()
        if len(uploaded_columns) == 0:
            response.status_code = 404
            response.data = json.dumps({'validate': f'未上傳檔案'})
            response.content_type = 'application/json'
            return response

        client_id = session.get('client_id', None)
        if client_id:
            if client_id not in connected_clients:
                connected_clients[client_id] = []
            connected_clients[client_id].append(project_name)

        print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} start {project_name}')
        for uploaded_column in uploaded_columns:
            column_ok, column_new_file, input_column_file = storefile(uploaded_column,
                                                                      app.config['UPLOAD_FOLDER'],
                                                                      app.config['OUTPUT_FOLDER'],
                                                                      f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')

            column_filename = input_column_file
            column_filenames.append(column_filename)
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'],
                                     f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-temp.pkl')

        if uploaded_xlsx:
            xlsx_ok, xlsx_new_file, input_xlsx_file = storefile(uploaded_xlsx,
                                                                app.config['UPLOAD_FOLDER'],
                                                                app.config['OUTPUT_FOLDER'],
                                                                f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')

            xlsx_filename = input_xlsx_file

        if uploaded_plans:
            xlsx_ok, xlsx_new_file, input_plan_file = storefile(uploaded_plans,
                                                                app.config['UPLOAD_FOLDER'],
                                                                app.config['OUTPUT_FOLDER'],
                                                                f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            plan_filename = input_plan_file

        layer_config = {
            'text_layer': request.form['column_text_layer'].split('\r\n'),
            'line_layer': request.form['column_line_layer'].split('\r\n'),
            # 箭頭和鋼筋文字的塗層
            'rebar_text_layer': request.form['column_rebar_text_layer'].split('\r\n'),
            # 鋼筋和箍筋的線的塗層
            'rebar_layer': request.form['column_rebar_layer'].split('\r\n'),
            # 箍筋文字圖層
            'tie_text_layer': request.form['column_tie_text_layer'].split('\r\n'),
            # 箍筋文字圖層
            'tie_layer': request.form['column_tie_layer'].split('\r\n'),
            # 框框的圖層
            'block_layer': request.form['column_block_layer'].split('\r\n'),
            # 斷面圖層
            'column_rc_layer': request.form['column_column_rc_layer'].split('\r\n'),
            'burst_layer_list': request.form['column_burst_layer_list'].split('\r\n'),
        }

        plan_layer_config = {
            'block_layer':  request.form['plan_block_layer'].split('\r\n'),
            'floor_text_layer':  request.form['plan_floor_text_layer'].split('\r\n'),
            'name_text_layer':  request.form['plan_serial_text_layer'].split('\r\n'),
        }

        result_log_content['upload_xlsx'] = uploaded_xlsx
        result_log_content['upload_beams'] = uploaded_columns
        result_log_content['project_name'] = project_name
        result_log_content['email_address'] = email_address
        result_log_content['template_name'] = template_name
        result_log_content['layer_config'] = layer_config
        # print(layer_config)

        if len(column_filenames) != 0 and temp_file != '' and column_ok:
            pkl_file_folder = os.path.join(app.config['UPLOAD_FOLDER'])
            hyper_parameter: dict = read_parameter_json(
                template_name=template_name)['column']

            hyper_parameter.update({'layer_config': layer_config,
                                    'plan_layer_config': plan_layer_config,
                                    'measure_type': measure_type})
            column_excel, column_report, pkl = count_column_multifiles(
                project_name=project_name,
                column_filenames=column_filenames,
                plan_filename=plan_filename,
                floor_parameter_xlsx=xlsx_filename,
                output_folder=app.config['OUTPUT_FOLDER'],
                pkl_file_folder=pkl_file_folder,
                client_id=client_id,
                **hyper_parameter
            )
            # column_excel = count_column_main(column_filename=column_filename,layer_config= layer_config,temp_file= temp_file,
            #                                  output_folder=app.config['OUTPUT_FOLDER'],project_name=project_name,template_name=template_name,floor_parameter_xlsx=xlsx_filename)
            # column_excel, column_report, pkl = count_column_multiprocessing(column_filenames=column_filenames,
            #                                                                 layer_config=layer_config,
            #                                                                 temp_file=temp_file,
            #                                                                 output_folder=app.config['OUTPUT_FOLDER'],
            #                                                                 project_name=project_name,
            #                                                                 template_name=template_name,
            #                                                                 floor_parameter_xlsx=xlsx_filename,
            #                                                                 progress_file=progress_file,
            #                                                                 client_id=client_id)
            if 'count_filenames' in session:
                session['count_filenames'].extend(
                    [column_excel, column_report])
            else:
                session['count_filenames'] = [column_excel, column_report]

            pkl = os.path.splitext(os.path.basename(pkl))[0]
            session['pkl_file_folder'] = pkl_file_folder

            if 'column_pkl_files' in session:
                session['column_pkl_files'].extend(pkl)
            else:
                session['column_pkl_files'] = [pkl]

        if (email_address):
            try:
                sendResult(email_address, [
                           column_excel, column_report], "梁配筋圖數量計算結果")
                print(
                    f'send_email:{email_address}, filenames:{session["count_filenames"]}')
            except Exception:
                pass

        response.status_code = 200
        # response.data = json.dumps({'validate': f'{layer_config}'})
        response.data = json.dumps({'validate': f'計算完成，請至輸出結果查看'})
        response.content_type = 'application/json'
        # return response
    except Exception as ex:

        with open(r'result\error_log.txt', 'a') as error_log:
            error_log.write(
                f'{project_name} | {ex} | column_layer = {layer_config} \n')
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate': f'發生錯誤'})
        response.content_type = 'application/json'
    print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} end {project_name}')
    write_result_log(file_path=r'result\result_log.txt',
                     result_content=result_log_content)
    return response

# @app.route('/send_email',methods=['POST'])


def sendResult(recipients: str, filenames: list, mail_title: str):
    output_folder = app.config['OUTPUT_FOLDER']
    # recipients = "elements.users29@gmail.com"
    # filenames = ["temp-0110_Markon.dwg","temp-0110_20230110_160947_rebar.txt","temp-0110_20230110_160947_rebar_floor.txt","temp-0110_20230110_160949_Count.xlsx"]
    with app.app_context():
        msg = Message(mail_title, recipients=[recipients])
        for filename in filenames:
            # filename = os.path.join(output_folder,filename)
            if ('.txt' in filename):
                content_type = "text/plain"
            if ('.xlsx' in filename):
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if ('.dwg' in filename):
                content_type = "application/x-dwg"
            with app.open_resource(os.path.join(output_folder, filename)) as fp:
                msg.attach(filename=filename, disposition="attachment",
                           content_type=content_type, data=fp.read())
        mail.send(msg)
    return 200


@app.errorhandler(404)
def page_not_found(e):
    return redirect(url_for('NOT_FOUND'))


def tail_logs(filename):
    """Implements tail -f command in Python to yield new log lines as they are written to the file."""
    with open(filename, 'r') as file:
        # Move the cursor to the end of the file
        file.seek(0, 2)

        while True:
            line = file.readline()
            if 'EOF' in line:
                yield line
                break
            if not line:
                time.sleep(0.5)  # Sleep briefly to avoid busy looping
                continue
            yield line


@app.route('/stream-logs')
def stream_logs():
    client_id = session.get('client_id')

    def generate():
        # tail_logs would need to parse and yield lines
        while not os.path.exists(f'logs/app_{client_id}.log'):
            time.sleep(5)
        for log_line in tail_logs(f'logs/app_{client_id}.log'):
            # print(f'user_id={user_id}:{log_line}')
            # if f'user_id={user_id}' in log_line:
            yield f"data: {log_line}\n\n"

    return Response(stream_with_context(generate()), mimetype='text/event-stream')


@app.route('/clear_session')
def clear_session():
    session.clear()
    session.modified = True
    return 'Session data cleared.'


@app.route('/get_template', methods=['POST'])
def get_fillin_template():
    if request.method == 'POST':
        data: dict = request.get_json()

        template = data.get('template')
        element_type = data.get('element_type')

        if element_type == 'beam':
            return jsonify(read_parameter_json(template_name=template)['beam']['layer_config'])
        if element_type == 'column':
            return jsonify(read_parameter_json(template_name=template)['column']['layer_config'])
# @app.route('/notify')
# def notify():
#     client_id = session.get('client_id')
#     print(f"Client {client_id}")
#     return Response(generate_notifications(client_id), content_type='text/event-stream')


if __name__ == '__main__':
    app.config.from_object('config.config.DevConfig')
    print('load config')
    # app.secret_key = 'dev'
    app.run(host='192.168.0.143', debug=True, port=8080)

    # print(secure_filename('2022-11-18-17-16temp-大梁.txt'))
