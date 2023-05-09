from multiprocessing import allow_connection_pickling
import os
from time import sleep
from urllib import response
from flask import Flask, request, redirect, url_for, render_template,send_from_directory,session,g, Response, stream_with_context,jsonify
from flask_mail import Mail, Message
from flask_session import Session
from werkzeug.utils import secure_filename
from main import main_functionV3, main_col_function,storefile,Output_Config
import functools
import json
import time
from datetime import timedelta
from auth import createPhoneCode,sendPhoneMessage
from beam_count import count_beam_multiprocessing
from column_count import count_column_main,count_column_multiprocessing
app = Flask(__name__)
app.config.from_object('config.config.Config')
Session(app)
mail= Mail(app)
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

def send_error_response(warning_message:str):
    print(warning_message)
    response = Response()
    response.status_code = 200
    response.data = json.dumps({'validate':f'{warning_message}'})
    response.content_type = 'application/json'
    return response

@app.route('/compare_beam', methods=['POST'])
# @login_required
def upload_file():
    if request.method == 'POST':
        try:
            # if session.get('post_status','accept') == 'disable':raise TabError
            # if time.time() - session['last_post_time'] < 60:
            #     raise ConnectionRefusedError
            # else:
            #     session['last_post_time'] = time.time()

            uploaded_beams = request.files.getlist("file1")
            uploaded_plans = request.files.getlist("file2")
            uploaded_columns = request.files.getlist("file_col")
            email_address = request.form['email_address']
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
            layer_config = {
                'text_layer':text_layer,
                'block_layer':block_layer,
                'floor_layer':floor_layer,
                'big_beam_layer':big_beam_layer,
                'big_beam_text_layer':big_beam_text_layer,
                'sml_beam_layer':sml_beam_layer,
                'size_layer':size_layer,
                'sml_beam_text_layer':sml_beam_text_layer
            }

            xs_col = request.form.get('xs-col')
            xs_beam = request.form.get('xs-beam')
            sizing = request.form.get('sizing')
            mline_scaling = request.form.get('mline_scaling')

            beam_ok = False
            plan_ok = False
            column_ok = False
            filenames = []
            project_name = time.strftime("%Y-%m-%d-%H-%M", time.localtime())+project_name
            print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} start {project_name}')
            progress_file = f'{app.config["OUTPUT_FOLDER"]}/{project_name}_progress'
            if len(uploaded_beams) > 1: dwg_type = 'muti'
            Output_Config(project_name=project_name,layer_config=layer_config,file_new_directory=app.config['OUTPUT_FOLDER'])
            for uploaded_beam in uploaded_beams:
                if uploaded_beam and allowed_file(uploaded_beam.filename) and xs_beam:
                    beam_ok, beam_new_file,input_beam_file = storefile(uploaded_beam,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                    # filename_beam = secure_filename(uploaded_beam.filename)
                    # beam_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_beam}'))
                    # print(input_beam_file)
                    beam_file.append(input_beam_file)
            for uploaded_column in uploaded_columns:
                if uploaded_column and allowed_file(uploaded_column.filename) and xs_col:
                    column_ok, column_new_file,input_column_file = storefile(uploaded_column,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                    # filename_column = secure_filename(uploaded_column.filename)
                    # column_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_column}'))
                    column_file.append(input_column_file)
            for uploaded_plan in uploaded_plans:
                if uploaded_plan and allowed_file(uploaded_plan.filename):
                    plan_ok, plan_new_file,input_plan_file = storefile(uploaded_plan,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],project_name)
                    # filename_plan = secure_filename(uploaded_plan.filename)
                    # plan_file.append(os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_plan}'))
                    plan_file.append(input_plan_file)
                    col_plan_new_file = f'{os.path.splitext(plan_new_file)[0]}_column.dwg'
                    print(col_plan_new_file)
                    # col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-column-{uploaded_plan.filename}')
            if beam_ok and len(plan_file)==1:filenames.append(os.path.split(plan_new_file)[1])
            if len(beam_file)==1:filenames.append(os.path.split(beam_new_file)[1])
            if len(column_file)==1:filenames.append(os.path.split(column_new_file)[1])
            if column_ok and len(plan_file)==1:filenames.append(os.path.split(col_plan_new_file)[1])
            # return
                # filename_plan = secure_filename(uploaded_plan.filename)
                # plan_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_plan}')
                # plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-{filename_plan}')
                # col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_COL_MARKON-{filename_plan}')
                # uploaded_plan.save(plan_file)
                # plan_ok = True
            if beam_ok and plan_ok:
                # main function
                # txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{beam_type}.txt')
                # sb_txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{sbeam_type}.txt')
                output_file = main_functionV3(beam_filenames = beam_file,
                                plan_filenames = plan_file,
                                beam_new_filename = beam_new_file,
                                plan_new_filename = plan_new_file,
                                output_directory = app.config['OUTPUT_FOLDER'],
                                project_name = project_name,
                                layer_config = layer_config,
                                progress_file = progress_file,
                                sizing = sizing,
                                mline_scaling = mline_scaling)
                filenames_beam = output_file
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
            if(email_address):
                try:
                    print(f'send_email:{email_address}, filenames:{filenames}')
                    sendResult(email_address,filenames,"配筋圖核對結果")
                except Exception as e: 
                    print(e)
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate':f'完成，請至輸出結果查看'})
            response.content_type = 'application/json'
            time.sleep(1)
        except ConnectionRefusedError:
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate':f'發送請求過於頻繁，請稍等'})
            response.content_type = 'application/json'
        except Exception as ex:
            print(ex)
            response = Response()
            response.status_code = 200
            response.data = json.dumps({'validate':f'發生錯誤'})
            response.content_type = 'application/json'
        print(f'{email_address}:{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} end {project_name}')
        return response
    return 400 
            # return render_template('tool1_result.html', filenames=filenames)
    # return render_template('tool1.html')

@app.route('/results')
@login_required
def result_page():
    filenames = session.get('filenames',[])
    count_filenames = session.get('count_filenames',[])
    # print(session.get('filenames',[]))
    # if filenames is None or len(filenames)==0:
    #     return render_template('tool1_result.html', filenames=[])
    # else:
    return render_template('tool1_result.html', filenames=filenames,count_filenames = count_filenames)

@app.route('/results/<filename>/',methods=['GET','POST'])
def result_file(filename):
    if(not filename in session.get('filenames',[]) and not filename in session.get('count_filenames',[])):return redirect('/')
    response = send_from_directory(app.config['OUTPUT_FOLDER'],
                               filename, as_attachment = True)
    response.cache_control.max_age = 0
    return response
@app.route('/demo/<filename>/',methods=['GET','POST'])
def demo_file(filename):
    response = send_from_directory(app.config['DEMO_FOLDER'],
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

@app.route('/tool2', methods=['GET','POST'])
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

@app.route('/admin_login', methods=['POST'])
def admin_login():
    user_code = request.form.get('user_code')
    if user_code == "wp32s%v9jhh!n+5i":
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
    pass
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['user_agree'] = 'agree'
        
        return redirect(url_for('home'))
    return render_template('statement.html', template_folder='./')

@app.route('/count_beam',methods=['POST'])
def count_beam():
    try:
        beam_filename = ''
        beam_filenames = []
        uploaded_xlsx = request.files['file_floor_xlsx']
        uploaded_beams = request.files.getlist("file_beam")
        project_name = request.form['project_name']
        email_address = request.form['email_address']
        print(request.form['company'])
        template_name = request.form['company']
        excel_filename = ''
        excel_filename_rcad = ''
        # project_name = time.strftime("%Y-%m-%d-%H-%M", time.localtime())+project_name
        beam_filename = ''
        temp_file = ''
        rebar_txt = ''
        rebar_excel = ''
        # rebar_input_file = os.path.join(app.config['OUTPUT_FOLDER'],rebar_file)
        if len(uploaded_beams) == 0:
            response.status_code = 404
            response.data = json.dumps({'validate':f'未上傳檔案'})
            response.content_type = 'application/json'
            return response
        for uploaded_beam in uploaded_beams:
            beam_ok, beam_new_file,input_beam_file = storefile(uploaded_beam,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # beam_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_beam.filename)}')
            beam_filename = input_beam_file
            beam_filenames.append(beam_filename)
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-temp.pkl')
            print(f'beam_filename:{beam_filename},temp_file:{temp_file}')
        if uploaded_xlsx:
            xlsx_ok, xlsx_new_file,input_xlsx_file = storefile(uploaded_xlsx,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # xlsx_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_xlsx.filename)}')
            xlsx_filename = input_xlsx_file
            print(f'xlsx_filename:{xlsx_filename}')
        layer_config = {
            'rebar_data_layer':request.form['rebar_data_layer'].split('\r\n'), # 箭頭和鋼筋文字的塗層
            'rebar_layer':request.form['rebar_layer'].split('\r\n'), # 鋼筋和箍筋的線的塗層
            'tie_text_layer':request.form['tie_text_layer'].split('\r\n'), # 箍筋文字圖層
            'block_layer':request.form['block_layer'].split('\r\n'), # 框框的圖層
            'beam_text_layer' :request.form['beam_text_layer'].split('\r\n'), # 梁的字串圖層
            'bounding_block_layer':request.form['bounding_block_layer'].split('\r\n'),
            'rc_block_layer':request.form['rc_block_layer'].split('\r\n'),
            's_dim_layer':['S-DIM']
            }
        print(layer_config)
        if beam_filename != '' and temp_file != '' and beam_ok:
            # rebar_txt,rebar_txt_floor,rebar_excel,rebar_dwg =count_beam_main(beam_filename=beam_filename,layer_config=layer_config,temp_file=temp_file,
            #                                                                     output_folder=app.config['OUTPUT_FOLDER'],project_name=project_name,template_name=template_name)
            excel_filename,excel_filename_rcad,output_dwg_list =  count_beam_multiprocessing(beam_filenames=beam_filenames,layer_config=layer_config,temp_file=temp_file,
                                                            project_name=project_name,output_folder=app.config['OUTPUT_FOLDER'],
                                                            template_name=template_name,floor_parameter_xlsx=xlsx_filename)
            # output_dwg_list = ['P2022-06A 岡山大鵬九村社宅12FB2_20230410_170229_Markon.dwg']
            if 'count_filenames' in session:
                session['count_filenames'].extend([excel_filename,excel_filename_rcad])
                session['count_filenames'].extend(output_dwg_list)
            else:
                session['count_filenames'] = [excel_filename,excel_filename_rcad]
                session['count_filenames'].extend(output_dwg_list)
        if(email_address):
            try:
                sendResult(email_address,[excel_filename,excel_filename_rcad],"梁配筋圖數量計算結果")
                sendResult(email_address,output_dwg_list,"梁配筋圖數量計算結果")
                print(f'send_email:{email_address}, filenames:{session["count_filenames"]}')
            except:
                pass
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate':f'計算完成，請至輸出結果查看'})
        response.content_type = 'application/json'
        # print(request.form['project_name'])
        time.sleep(1)
    except Exception as ex:
        print(ex)
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate':f'發生錯誤'})
        response.content_type = 'application/json'
    return response
@app.route('/count_column',methods=['POST'])
def count_column():
    try:
        uploaded_columns = request.files.getlist("file_column")
        project_name = request.form['project_name']
        email_address = request.form['email_address']
        template_name = request.form['companyColumn']
        uploaded_xlsx = request.files['file_floor_xlsx']
        column_filename = ''
        column_filenames = []
        column_excel = ''
        column_ok = False
        if len(uploaded_columns) == 0:
            response.status_code = 404
            response.data = json.dumps({'validate':f'未上傳檔案'})
            response.content_type = 'application/json'
            return response
        for uploaded_column in uploaded_columns:
            column_ok, column_new_file,input_column_file = storefile(uploaded_column,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # column_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_column.filename)}')
            column_filename = input_column_file
            column_filenames.append(column_filename)
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-temp.pkl')
            print(f'column_filename:{column_filename},temp_file:{temp_file}')
        if uploaded_xlsx:
            xlsx_ok, xlsx_new_file,input_xlsx_file = storefile(uploaded_xlsx,app.config['UPLOAD_FOLDER'],app.config['OUTPUT_FOLDER'],f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}')
            # xlsx_filename = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{time.strftime("%Y-%m-%d-%H-%M", time.localtime())}-{secure_filename(uploaded_xlsx.filename)}')
            xlsx_filename = input_xlsx_file
            print(f'xlsx_filename:{xlsx_filename}')
        layer_config = {
            'text_layer':request.form['column_text_layer'].split('\r\n'),
            'line_layer':request.form['column_line_layer'].split('\r\n'),
            'rebar_text_layer':request.form['column_rebar_text_layer'].split('\r\n'), # 箭頭和鋼筋文字的塗層
            'rebar_layer':request.form['column_rebar_layer'].split('\r\n'), # 鋼筋和箍筋的線的塗層
            'tie_text_layer':request.form['column_tie_text_layer'].split('\r\n'), # 箍筋文字圖層
            'tie_layer':request.form['column_tie_layer'].split('\r\n'), # 箍筋文字圖層
            'block_layer':request.form['column_block_layer'].split('\r\n'), # 框框的圖層
            'column_rc_layer':request.form['column_rc_layer'].split('\r\n'), #斷面圖層
        }
        print(layer_config)
        if len(column_filenames) != 0 and temp_file != '' and column_ok:
            # column_excel = count_column_main(column_filename=column_filename,layer_config= layer_config,temp_file= temp_file,
            #                                  output_folder=app.config['OUTPUT_FOLDER'],project_name=project_name,template_name=template_name,floor_parameter_xlsx=xlsx_filename)
            column_excel =count_column_multiprocessing(column_filenames=column_filenames,layer_config=layer_config,temp_file=temp_file,
                                                        output_folder=app.config['OUTPUT_FOLDER'],project_name=project_name,
                                                        template_name=template_name,floor_parameter_xlsx=xlsx_filename)
            if 'count_filenames' in session:
                session['count_filenames'].extend([column_excel])
            else:
                session['count_filenames'] = [column_excel]
        if(email_address):
            try:
                sendResult(email_address,[column_excel],"梁配筋圖數量計算結果")
                print(f'send_email:{email_address}, filenames:{session["count_filenames"]}')
            except:
                pass
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate':f'{layer_config}'})
        response.content_type = 'application/json'
        return response
    except Exception as ex:
        print(ex)
        response = Response()
        response.status_code = 200
        response.data = json.dumps({'validate':f'發生錯誤'})
        response.content_type = 'application/json'
    return response
    
# @app.route('/send_email',methods=['POST'])
def sendResult(recipients:str,filenames:list,mail_title:str):
    output_folder = app.config['OUTPUT_FOLDER']
    # recipients = "elements.users29@gmail.com"
    # filenames = ["temp-0110_Markon.dwg","temp-0110_20230110_160947_rebar.txt","temp-0110_20230110_160947_rebar_floor.txt","temp-0110_20230110_160949_Count.xlsx"]
    with app.app_context():
        msg = Message(mail_title,recipients=[recipients])
        for filename in filenames:
            # filename = os.path.join(output_folder,filename)
            if('.txt' in filename):content_type = "text/plain"
            if('.xlsx' in filename):content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if('.dwg' in filename):content_type = "application/x-dwg"
            with app.open_resource(os.path.join(output_folder,filename)) as fp:
                msg.attach(filename=filename,disposition="attachment",content_type=content_type,data=fp.read())
        mail.send(msg)
    return 200
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
    print('load config')
    # app.secret_key = 'dev'
    app.run(host = '192.168.0.143',debug=True,port=8080)

    # print(secure_filename('2022-11-18-17-16temp-大梁.txt'))