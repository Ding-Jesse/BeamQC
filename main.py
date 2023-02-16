from multiprocessing.spawn import prepare
import time
import multiprocessing
# from AutocadApi import read_plan,read_beam,write_beam,write_plan,error
from plan_to_beam import read_plan,read_beam,write_beam,write_plan,write_result_log
from werkzeug.utils import secure_filename
import os
import plan_to_col
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment,Font,PatternFill
from openpyxl.worksheet.worksheet import Worksheet

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
        plan = plan.get()
        if plan:
            set_plan = set_plan | plan[0]
            if plan_drawing:
                dic_plan = plan[1]
        else:
            end = time.time()
            write_result_log(excel_file, task_name, '', '', '', '', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return

    for beam in res_beam:
        beam = beam.get()
        if beam:
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
        plan = plan.get()
        if plan:
            set_plan = set_plan | plan[0]
            if plan_drawing:
                dic_plan = plan[1]
        else:
            end = time.time()
            plan_to_col.write_result_log(excel_file,task_name,'','', f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
            return 

    for col in res_col:
        col = col.get()
        if col:
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

def OutputExcel(df:pd.DataFrame,file_path,sheet_name,auto_fit_columns=[],auto_fit_rows=[],columns_list=[],rows_list=[]):
    if os.path.exists(file_path):
        book = load_workbook(file_path)
        writer = pd.ExcelWriter(file_path, engine='openpyxl') 
        writer.book = book
        # sheet = book[sheet_name]
        # sheet.column_dimensions['A'] =ColumnDimension(sheet,'L',bestFit=True)
    else:
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter') 
    df.to_excel(writer,sheet_name=sheet_name)
    writer.save()

    book = load_workbook(file_path)
    writer = pd.ExcelWriter(file_path, engine='openpyxl') 
    writer.book = book
    if os.path.exists(file_path) and len(auto_fit_columns) >0:
        AutoFit_Columns(book[sheet_name],auto_fit_columns,auto_fit_rows)
    if os.path.exists(file_path) and len(columns_list) >0:
        Decorate_Worksheet(book[sheet_name],columns_list,rows_list)
    writer.save()
    return file_path

def Decorate_Worksheet(sheet:Worksheet,columns_list:list,rows_list:list):
    for i in columns_list:
        for j in rows_list:
            sheet.cell(j,i).alignment = Alignment(vertical='center',wrap_text=True,horizontal='center')
            sheet.cell(j,i).font = Font(name='Calibri')
            if sheet.cell(j,i).value == 'NG.':sheet.cell(j,i).fill = PatternFill("solid",start_color='00FF0000')

def AutoFit_Columns(sheet:Worksheet,auto_fit_columns:list,auto_fit_rows:list):
    for i in auto_fit_columns:
        sheet.column_dimensions[get_column_letter(i)].width = 80
    for i in auto_fit_rows:
        sheet.row_dimensions[i].height = 20
    for i in auto_fit_rows:
        for j in auto_fit_columns:
            sheet.cell(i,j).alignment = Alignment(wrap_text=True,vertical='center',horizontal='center')

if __name__ == '__main__':
    from item.column import Column
    from item.beam import Beam
    from item.floor import Floor
    c = Column()
    b = Beam('1',1,1)
    f = Floor('1')