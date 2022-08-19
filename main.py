from multiprocessing.spawn import prepare
import time
import multiprocessing
# from AutocadApi import read_plan,read_beam,write_beam,write_plan,error
from plan_to_beam import read_plan,read_beam,write_beam,write_plan,write_result_log
from werkzeug.utils import secure_filename
import os
import plan_to_col
from datetime import datetime

def main_functionV3(beam_filenames,plan_filenames,beam_new_filename,plan_new_filename,big_file,sml_file,text_layer,block_layer,floor_layer,size_layer,big_beam_layer,big_beam_text_layer,sml_beam_layer,sml_beam_text_layer,task_name,progress_file,sizing,mline_scaling):
    start = time.time()

    plan_file = './result/plan.txt' # plan.txt的路徑
    beam_file = './result/beam.txt' # beam.txt的路徑
    excel_file = './result/result_log.xlsx' # result_log.xlsx的路徑
    
    date = time.strftime("%Y-%m-%d", time.localtime())

    multiprocessing.freeze_support()    
    pool = multiprocessing.Pool()

    res_plan =[]
    res_beam = []
    set_plan = set()
    dic_plan = {}
    set_beam = set()
    dic_beam = {}

    for plan_filename in plan_filenames:
        res_plan.append(pool.apply_async(read_plan, (plan_filename, plan_new_filename, big_file, sml_file, floor_layer, big_beam_layer, big_beam_text_layer, sml_beam_layer, sml_beam_text_layer, block_layer, size_layer, plan_file, progress_file, sizing, mline_scaling, date)))
    
    for beam_filename in beam_filenames:
        res_beam.append(pool.apply_async(read_beam, (beam_filename, text_layer, beam_file, progress_file, sizing)))
    
    plan_drawing = 0
    if len(plan_filenames) == 1:
        plan_drawing = 1
    beam_drawing = 0
    if len(beam_filenames) == 1:
        beam_drawing = 1

    for plan in res_plan:
        if plan:
            plan = plan.get()
            set_plan = set_plan | plan[0]
            if plan_drawing:
                dic_plan = plan[1]
        else:
            end = time.time()
            write_result_log(excel_file, task_name, '', '', '', '', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return

    for beam in res_beam:
        if beam:
            beam = beam.get()
            set_beam = set_beam | beam[0]
            if beam_drawing:
                dic_beam = beam[1]
        else:
            end = time.time()
            write_result_log(excel_file, task_name, '', '', '', '', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return

    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, plan_drawing, progress_file, sizing, mline_scaling)
    beam_result = write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, beam_drawing, progress_file, sizing)

    end = time.time()
    write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')

def main_col_function(col_filenames,plan_filenames,col_new_filename,plan_new_filename,result_file,text_layer,line_layer,block_layer,floor_layer,col_layer,task_name,progress_file):
    start = time.time()

    plan_file = './result/col_plan.txt' # plan.txt的路徑
    col_file = './result/col.txt' # col.txt的路徑
    excel_file = './result/result_log_col.xlsx' # result_log.xlsx的路徑

    date = time.strftime("%Y-%m-%d", time.localtime())

    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()
    res_plan =[]
    res_col = []
    set_plan = set()
    dic_plan = {}
    set_col = set()
    dic_col = {}
    for plan_filename in plan_filenames:
        res_plan.append(pool.apply_async(plan_to_col.read_plan, (plan_filename, floor_layer, col_layer, block_layer, plan_file, progress_file)))
    for col_filename in col_filenames:
        res_col.append(pool.apply_async(plan_to_col.read_col, (col_filename, text_layer, line_layer, col_file, progress_file)))

    plan_drawing = 0
    if len(plan_filenames) == 1:
        plan_drawing = 1
    col_drawing = 0
    if len(col_filenames) == 1:
        col_drawing = 1

    for plan in res_plan:
        if plan:
            plan = plan.get()
            set_plan = set_plan | plan[0]
            if plan_drawing:
                dic_plan = plan[1]
        else:
            end = time.time()
            plan_to_col.write_result_log(excel_file,task_name,'','', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return 

    for col in res_col:
        if col:
            col = col.get()
            set_col = set_col | col[0]
            if col_drawing:
                dic_col = col[1]
        else:
            end = time.time()
            plan_to_col.write_result_log(excel_file,task_name,'','', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return 
                
    plan_result = plan_to_col.write_plan(plan_filename, plan_new_filename, set_plan, set_col, dic_plan, result_file, date, plan_drawing, progress_file)
    col_result = plan_to_col.write_col(col_filename, col_new_filename, set_plan, set_col, dic_col, result_file, date,col_drawing, progress_file)

    end = time.time()
    plan_to_col.write_result_log(excel_file,task_name,plan_result,col_result, f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    return 

def storefile(file,file_directory,file_new_directory,project_name):
    filename_beam = secure_filename(file.filename)
    save_file = os.path.join(file_directory, f'{project_name}-{filename_beam}')
    file_new_name = os.path.join(file_new_directory, f'{project_name}_MARKON-{filename_beam}')
    file.save(save_file)
    file_ok = True
    return file_ok , file_new_name
