from multiprocessing import allow_connection_pickling
import os
from time import sleep
from flask import Flask, request, redirect, url_for, render_template,send_from_directory
from werkzeug.utils import secure_filename
from main import main_functionV3, main_col_function

import time
app = Flask(__name__, template_folder='./')
UPLOAD_FOLDER = 'C:/Users/Vince/Desktop/BeamQC/INPUT'
OUTPUT_FOLDER = 'C:/Users/Vince/Desktop/BeamQC/OUTPUT'
ALLOWED_EXTENSIONS = set(['dwg'])
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 60MB

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/tool1', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_beam = request.files["file1"]
        uploaded_plan = request.files["file2"]
        uploaded_column = request.files["file_col"]
        beam_type = '大梁'
        sbeam_type = '小梁'
        beam_file =''
        plan_file = ''
        txt_file =''
        project_name = request.form['project_name']
        text_col_layer = request.form['text_col_layer']
        line_layer = request.form['line_layer']
        text_layer = request.form['text_layer']
        block_layer = request.form['block_layer']
        floor_layer = request.form['floor_layer']
        big_beam_layer = request.form['big_beam_layer']
        sml_beam_layer = request.form['sml_beam_layer']
        col_layer = request.form['col_layer']
        size_layer = request.form['size_layer']
        explode_plan = request.form.get('explode_plan')
        explode_beam = request.form.get('explode_beam')
        explode_col = request.form.get('explode_col')
        xs_col = request.form.get('xs-col')
        xs_beam = request.form.get('xs-beam')
        beam_ok = False
        plan_ok = False
        column_ok = False
        filenames = []
        if uploaded_beam and allowed_file(uploaded_beam.filename) and xs_beam:
            filename_beam = secure_filename(uploaded_beam.filename)
            beam_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_beam}')
            beam_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-{filename_beam}')
            uploaded_beam.save(beam_file)
            beam_ok = True
        if uploaded_column and allowed_file(uploaded_column.filename) and xs_col:
            filename_column = secure_filename(uploaded_column.filename)
            column_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_column}')
            column_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-{filename_column}')
            uploaded_column.save(column_file)
            column_ok = True
        if uploaded_plan and allowed_file(uploaded_plan.filename):
            filename_plan = secure_filename(uploaded_plan.filename)
            plan_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{project_name}-{filename_plan}')
            plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-{filename_plan}')
            # col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_COL_MARKON-{filename_plan}')
            uploaded_plan.save(plan_file)
            plan_ok = True
        if beam_ok and plan_ok:
            # main function
            txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{beam_type}.txt')
            sb_txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-{sbeam_type}.txt')
            os.system(f'python plan_to_beam.py {beam_file} {plan_file} {beam_new_file} {plan_new_file} {txt_file} {sb_txt_file} {text_layer} {block_layer} {floor_layer} {big_beam_layer} {sml_beam_layer} {size_layer} {project_name} {explode_plan} {explode_beam}')
            # main_functionV3(beam_file,plan_file,beam_new_file,plan_new_file,txt_file,sb_txt_file,block_layer,project_name,explode)
            filenames_beam = [f'{project_name}-{beam_type}.txt',f'{project_name}-{sbeam_type}.txt',
                                f'{project_name}_MARKON-{filename_beam}',f'{project_name}_MARKON-{filename_plan}']
            filenames.extend(filenames_beam) 
        if column_ok and plan_ok:
            # main function
            txt_file = os.path.join(app.config['OUTPUT_FOLDER'],f'{project_name}-column.txt')
            col_plan_new_file = os.path.join(app.config['OUTPUT_FOLDER'], f'{project_name}_MARKON-column-{filename_plan}')
            os.system(f'python plan_to_col.py {column_file} {plan_file} {column_new_file} {col_plan_new_file} {txt_file} {text_col_layer} {line_layer} {block_layer} {floor_layer} {col_layer} {project_name} {explode_plan} {explode_col}')
            # main_col_function(column_file,plan_file,column_new_file,col_plan_new_file,txt_file,block_layer,project_name,explode)
            filenames_column = [f'{project_name}-column.txt',
                                f'{project_name}_MARKON-{filename_column}',
                                f'{project_name}_MARKON-column-{filename_plan}']
            filenames.extend(filenames_column)
        if column_ok or beam_ok:
            return render_template('tool1_result.html', filenames=filenames)
    return render_template('tool1.html')

@app.route('/results/<filename>/')
def result_file(filename):
    response = send_from_directory(app.config['OUTPUT_FOLDER'],
                               filename, as_attachment = True)
    response.cache_control.max_age = 0
    return response

@app.route('/tool1_README')
def tool2():
    return render_template('tool1_README.html')

@app.route('/tool3')
def tool3():
    return render_template('tool3.html')

@app.route('/tool4')
def tool4():
    return render_template('tool4.html')

@app.route('/tool5')
def tool5():
    return render_template('tool5.html')