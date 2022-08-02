from multiprocessing.spawn import prepare
import time
import multiprocessing
# from AutocadApi import read_plan,read_beam,write_beam,write_plan,error
from plan_to_beam import read_plan,read_beam,write_beam,write_plan,error, write_result_log
from werkzeug.utils import secure_filename
import os
import plan_to_col
from datetime import datetime

def main_functionV3(beam_filenames,plan_filenames,beam_new_filename,plan_new_filename,big_file ,sml_file,block_layer,task_name,explode):
    start = time.time()
    # task_name = 'task3'
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # plan_filename = "K:/100_Users/EI 202208 Bamboo/BeamQC/task3/XS-PLAN.dwg" # XS-PLAN的路徑
    # beam_filename = "K:/100_Users/EI 202208 Bamboo/BeamQC/task3/XS-BEAM.dwg" # XS-BEAM的路徑
    # plan_new_filename = f"K:/100_Users/EI 202208 Bamboo/BeamQC/task3/{task_name}-XS-PLAN_new.dwg" # XS-PLAN_new的路徑
    # beam_new_filename = f"K:/100_Users/EI 202208 Bamboo/BeamQC/task3/{task_name}-XS-BEAM_new.dwg" # XS-BEAM_new的路徑
    plan_file = './result/beam_plan.txt' # plan.txt的路徑
    beam_file = './result/beam.txt' # beam.txt的路徑
    excel_file = './result/result_log.xlsx' # result_log.xlsx的路徑
    size_layer = 'S-TEXT'
    # big_file = final_filename # 大梁結果
    # sml_file = sb_final_filename # 小梁結果

    date = time.strftime("%Y-%m-%d", time.localtime())
    
    # 在plan裡面自訂圖層
    floor_layer = "S-TITLE" # 樓層字串的圖層
    beam_layer = ["S-TEXTG", "S-TEXTB"] # beam的圖層，因為有兩個以上，所以用list來存
    # block_layer = "DEFPOINTS" # 框框的圖層
    # explode = 0 # 需不需要提前炸圖塊

    # 在beam裡面自訂圖層
    text_layer = "S-RC"
    multiprocessing.freeze_support()    
    pool = multiprocessing.Pool()
    print('Start Reading')
    res_plan =[]
    res_beam = []
    set_plan = set()
    dic_plan = {}
    set_beam = set()
    dic_beam = {}
    for plan_filename in plan_filenames:
        res_plan.append(pool.apply_async(read_plan, (plan_filename,floor_layer, beam_layer, block_layer, size_layer, plan_file, explode)))
    for beam_filename in beam_filenames:
        res_beam.append(pool.apply_async(read_beam, (beam_filename, text_layer, beam_file,0)))
    
    plan_drawing = 0
    if len(plan_filenames) == 1:
        plan_drawing = 1
    beam_drawing = 0
    if len(beam_filenames) == 1:
        beam_drawing = 1

    for plan in res_plan:
        plan = plan.get()
        set_plan = set_plan | plan[0]
        if plan_drawing:
            dic_plan = plan[1]

    for beam in res_beam:
        beam = beam.get()
        set_beam = set_beam | beam[0]
        if beam_drawing:
            dic_beam = beam[1]

    print('Start Writing')
    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date,plan_drawing)
    beam_result = write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date,beam_drawing)

    end = time.time()
    write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    # write_result_log(excel_file,'','','','','','',time.strftime("%Y-%m-%d %H:%M", time.localtime()),'')

def main_col_function(col_filenames,plan_filenames,col_new_filename,plan_new_filename,result_file,block_layer,task_name, explode):
    
    start = time.time()
    # task_name = 'task13'
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # plan_filename = "C:/Users/Vince/Desktop/BeamQC/data/task13/XS-PLAN.dwg" # XS-PLAN的路徑
    # col_filename = "C:/Users/Vince/Desktop/BeamQC/data/task13/XS-COL.dwg" # XS-COL的路徑
    # plan_new_filename = f"C:/Users/Vince/Desktop/BeamQC/data/task13/{task_name}-XS-PLAN_new.dwg" # XS-PLAN_new的路徑
    # col_new_filename = f"C:/Users/Vince/Desktop/BeamQC/data/task13/{task_name}-XS-COL_new.dwg" # XS-COL_new的路徑
    plan_file = './result/col_plan.txt' # plan.txt的路徑
    col_file = './result/col.txt' # col.txt的路徑
    excel_file = './result/result_log_col.xlsx' # result_log.xlsx的路徑
    # result_file = f"C:/Users/Vince/Desktop/BeamQC/data/task13/{task_name}-柱配筋.txt" # 柱配筋結果

    date = time.strftime("%Y-%m-%d", time.localtime())
    
    # 在plan裡面自訂圖層
    floor_layer = "S-TITLE" # 樓層字串的圖層
    col_layer = "S-TEXTC" # col的圖層
    # block_layer = "DwFm" # 圖框的圖層
    # explode = 0 # 需不需要提前炸圖塊

    # 在col裡面自訂圖層
    text_layer = "S-TEXT" # 文字的圖層
    line_layer = "S-STUD" # 線的圖層
    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()
    print('Start Reading')
    res_plan =[]
    res_col = []
    set_plan = set()
    dic_plan = {}
    set_col = set()
    dic_col = {}
    for plan_filename in plan_filenames:
        res_plan.append(pool.apply_async(plan_to_col.read_plan, (plan_filename, floor_layer, col_layer, block_layer, plan_file, explode)))
    for col_filename in col_filenames:
        res_col.append(pool.apply_async(plan_to_col.read_col, (col_filename, text_layer, line_layer, col_file, explode)))

    plan_drawing = 0
    if len(plan_filenames) == 1:
        plan_drawing = 1
    col_drawing = 0
    if len(col_filenames) == 1:
        col_drawing = 1

    for plan in res_plan:
        plan = plan.get()
        set_plan = set_plan | plan[0]
        if plan_drawing:
            dic_plan = plan[1]

    for col in res_col:
        col = col.get()
        set_col = set_col | col[0]
        if col_drawing:
            dic_col = col[1]
    
    plan_result = plan_to_col.write_plan(plan_filename, plan_new_filename, set_plan, set_col, dic_plan, result_file, date, plan_drawing)
    col_result = plan_to_col.write_col(col_filename, col_new_filename, set_plan, set_col, dic_col, result_file, date,col_drawing)

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
