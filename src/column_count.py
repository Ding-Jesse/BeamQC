from __future__ import annotations
import re
import win32com.client
import pythoncom
import time
import copy
import numpy as np
import pandas as pd
import os
import src.save_temp_file as save_temp_file
from typing import Literal
from collections import defaultdict
from src.logger import setup_custom_logger
from utils.demand import calculate_column_beam_joint_shear
from item.pdf import create_scan_pdf
from item.rebar import readRebarExcel
from item.floor import Floor, read_parameter_df, summary_floor_rebar, summary_floor_column_rebar_ratio
from multiprocessing.pool import ThreadPool as Pool
from src.main import OutputExcel
from src.column_scan import (column_check, create_column_scan,
                             output_detail_scan_report, output_ng_ratio)
from src.beam_count import vtPnt, read_parameter_json
from src.plan_count import sort_plan_count
from item.beam import Beam
from item.column import Column
from src.plan_to_beam import (turn_floor_to_float, turn_floor_to_string,
                              turn_floor_to_list, floor_exist, vtFloat)


slash_pattern = r'(.+)F[~|-](.+)F'  # ~
commom_pattern = r'(,|、)'
multi = True

global main_logger


def progress(message):  # 把進度印到log裡面，在app.py會對這個檔案做事
    main_logger.info(message)


def error(message):  # 把錯誤印到log裡面，在app.py會對這個檔案做事
    main_logger.error(message)


def read_column_cad(column_filename):
    # layer_list = [layer for key,layer in layer_config.items()]
    # line_layer = layer_config['line_layer']
    error_count = 0
    pythoncom.CoInitialize()
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_column = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_column error in step 1: {e}, error_count = {error_count}.')

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 30:
        try:
            doc_column = wincad_column.Documents.Open(column_filename)
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f' {os.path.basename(column_filename)} read_column error in step 2: {e}, error_count = {error_count}.')

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 30:
        try:
            msp_column = doc_column.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_column error in step 3: {e}, error_count = {error_count}.')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 30:
        try:
            doc_column.Close(SaveChanges=False)
        except:
            pass
        return False
    # flag = 0
    # # Step 4. 解鎖所有圖層 -> 不然不能刪東西
    # while not flag and error_count <= 10:
    #     try:
    #         layer_count = doc_column.Layers.count

    #         for x in range(layer_count):
    #             layer = doc_column.Layers.Item(x)
    #             layer.Lock = False
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         msp_column  = doc_column.Modelspace
    #         error(f'read_col error in step 4: {e}, error_count = {error_count}.')
    # progress('柱配筋圖讀取進度 4/10', progress_file)

    # Step 5. 遍歷所有物件 -> 炸圖塊
    # if burst_layer_list:
    #     print('burst')
    #     flag = 0
    #     while not flag and error_count <= 10:
    #         try:
    #             count = 0
    #             total = msp_column.Count
    #             # layer_list = [text_layer, line_layer]
    #             for object in msp_column:
    #                 count += 1
    #                 if object.EntityName == "AcDbBlockReference" and object.Layer in burst_layer_list:
    #                     object.Explode()
    #             flag = 1
    #         except Exception as e:
    #             error_count += 1
    #             msp_column = doc_column.Modelspace
    #             time.sleep(5)
    #             error(
    #                 f'read_col error in step 5: {e}, error_count = {error_count}.')
        # progress('柱配筋圖讀取進度 5/10')
    progress(f'{column_filename} : Finish')
    # Step 6. 重新匯入modelspace
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         msp_column = doc_column.Modelspace
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_col error in step 6: {e}, error_count = {error_count}.')
    return msp_column, doc_column


def sort_col_cad(msp_column,
                 doc_column,
                 layer_config: dict,
                 temp_file: str):
    layer_config = {key: value for key, value in layer_config.items()}
    text_layer = list(layer_config['text_layer'])
    line_layer = list(layer_config['line_layer'])
    rebar_text_layer = list(layer_config['rebar_text_layer'])
    rebar_layer = list(layer_config['rebar_layer'])
    tie_layer = list(layer_config['tie_layer'])
    tie_text_layer = list(layer_config['tie_text_layer'])
    column_rc_layer = list(layer_config['column_rc_layer'])
    burst_layer_list = list(layer_config['burst_layer_list'])
    coor_to_floor_set = set()  # set(coor, floor)
    coor_to_col_set = set()  # set(coor, col)
    coor_to_size_set = set()  # set(coor, size)
    coor_to_floor_slash_set = set()
    coor_to_floor_line_list = []  # (橫線y座標, start, end)
    coor_to_col_line_list = []  # (縱線x座標, start, end)
    coor_to_rebar_text_list = []
    coor_to_rebar_list = []
    coor_to_tie_text_list = []
    coor_to_tie_list = []
    coor_to_section_list = []
    error_count = 0
    while error_count <= 30:
        try:
            total = msp_column.Count
            break
        except:
            error_count += 1
    # try:
    progress(
        f'柱配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
    for count, msp_object in enumerate(msp_column):
        object_list = []
        error_count = 0
        if count % 1000 == 0:
            progress(f'柱配筋圖已讀取{count}/{total}個物件')
        while error_count <= 3 and not object_list:
            try:
                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference" and msp_object.Layer in burst_layer_list:
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read error in step 7: {ex}, error_count = {error_count}.')
        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                if object.Layer in tie_layer:
                    if object.ObjectName == "AcDbPolyline" or\
                            object.ObjectName == "AcDb2dPolyline" or object.ObjectName == "AcDbLine":
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_tie_list.append((coor1, coor2))
                if object.Layer in tie_text_layer:
                    if object.ObjectName == "AcDbText":
                        if not hasattr(object, 'GetBoundingBox'):
                            coor1 = (round(object.InsertionPoint[0], 2), round(
                                object.InsertionPoint[1], 2))
                            coor2 = (round(object.InsertionPoint[0], 2), round(
                                object.InsertionPoint[1], 2))
                        else:
                            coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                                object.GetBoundingBox()[0][1], 2))
                            coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                                object.GetBoundingBox()[1][1], 2))
                        coor_to_tie_text_list.append(
                            ((coor1, coor2), object.TextString))
                if object.Layer in rebar_text_layer and object.ObjectName == "AcDbText":
                    if re.search(r'\d+.#\d+', object.TextString):
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_rebar_text_list.append(
                            ((coor1, coor2), object.TextString))
                if object.Layer in rebar_layer:
                    if object.ObjectName == "AcDbCircle":
                        coor1 = (round(object.Center[0], 2), round(
                            object.Center[1], 2))
                        coor_to_rebar_list.append((coor1, 'circle'))
                    if object.ObjectName == "AcDbEllipse":
                        coor1 = (round(object.Center[0], 2), round(
                            object.Center[1], 2))
                        coor_to_rebar_list.append((coor1, 'ellipse'))
                    # if object.ObjectName == "AcDbPolyline":
                    #     coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    #     coor_to_rebar_list.append(coor1)
                    if object.ObjectName == "AcDbBlockReference":
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor_to_rebar_list.append((coor1, object.Name))

                if object.Layer in text_layer and object.ObjectName == "AcDbText":
                    if (object.TextString[0] == 'C' or 'EC' in object.TextString) \
                            and len(object.TextString) <= 7:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        match_obj = re.search(slash_pattern, object.TextString)
                        if match_obj:
                            suffix_index = re.search(
                                r'(\D*)\d+(\D*)', match_obj.group(1))
                            first_column = re.findall(
                                r'\d+', match_obj.group(1))
                            last_column = re.findall(
                                r'\d+', match_obj.group(2))
                            if first_column and last_column:
                                for column_number in range(int(first_column[0]), int(last_column[0]) + 1):
                                    temp_string = f'{suffix_index.group(1)}{column_number}{suffix_index.group(2)}'
                                    coor_to_col_set.add(
                                        ((coor1, coor2), temp_string))
                            else:
                                coor_to_col_set.add(
                                    ((coor1, coor2), object.TextString))
                        elif re.search(commom_pattern, object.TextString):
                            sep = re.search(
                                commom_pattern, object.TextString).group(1)
                            for column_text in object.TextString.split(sep):
                                coor_to_col_set.add(
                                    ((coor1, coor2), column_text))
                        else:
                            coor_to_col_set.add(
                                ((coor1, coor2), object.TextString))

                    elif 'x' in object.TextString or \
                            'X' in object.TextString:
                        size = object.TextString.replace('X', 'x')
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_size_set.add(((coor1, coor2), size))
                    elif ('F' in object.TextString or
                          'B' in object.TextString or
                          'R' in object.TextString or
                          re.match(r'\d+', object.TextString)) and \
                            'O' not in object.TextString:  # 可能有樓層
                        floor = object.TextString
                        if '_' in floor:  # 可能有B_6F表示B棟的6F
                            floor = floor.split('_')[1]
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        if re.search(slash_pattern, floor):
                            match_obj = re.search(slash_pattern, floor)
                            first_floor = int(
                                turn_floor_to_float(match_obj.group(1)))
                            last_floor = int(
                                turn_floor_to_float(match_obj.group(2)))
                            for floor_float in range(first_floor, last_floor + 1):
                                coor_to_floor_set.add(
                                    ((coor1, coor2), turn_floor_to_string(floor_float)))
                        elif re.search(commom_pattern, floor):
                            sep = re.search(
                                commom_pattern, object.TextString).group()
                            for floor_float in floor.split(sep):
                                if isinstance(floor_float, float) or isinstance(floor_float, int):
                                    coor_to_floor_set.add(
                                        ((coor1, coor2), turn_floor_to_string(floor_float)))
                        else:
                            if turn_floor_to_float(floor):
                                coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                                    object.GetBoundingBox()[0][1], 2))
                                coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                                    object.GetBoundingBox()[1][1], 2))
                                floor = turn_floor_to_float(floor)
                                floor = turn_floor_to_string(floor)
                                coor_to_floor_set.add(((coor1, coor2), floor))
                    elif object.TextString in ['~', '-']:
                        # for slash text not combine with floor text
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_floor_slash_set.add(((coor1, coor2), floor))

                elif object.Layer in line_layer:
                    # consider line as polyline
                    if object.ObjectName == "AcDbPolyline" and len(object.Coordinates) == 8:
                        for k in range(0, len(object.Coordinates), 2):
                            coor1 = (round(object.Coordinates[k], 2), round(
                                object.Coordinates[k+1], 2))
                            if k + 1 == len(object.Coordinates) - 1:
                                coor2 = (round(object.Coordinates[0], 2), round(
                                    object.Coordinates[1], 2))
                            else:
                                coor2 = (round(object.Coordinates[k + 2], 2), round(
                                    object.Coordinates[k + 3], 2))
                            if abs(coor1[0] - coor2[0]) < 5:
                                coor_to_col_line_list.append(
                                    (coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1])))
                            elif abs(coor1[1] - coor2[1]) < 5:
                                coor_to_floor_line_list.append(
                                    (coor1[1], min(coor1[0], coor2[0]), max(coor1[0], coor2[0])))
                    else:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        if abs(coor1[0] - coor2[0]) < 5:
                            coor_to_col_line_list.append(
                                (coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1])))
                        elif abs(coor1[1] - coor2[1]) < 5:
                            coor_to_floor_line_list.append(
                                (coor1[1], min(coor1[0], coor2[0]), max(coor1[0], coor2[0])))
                if object.Layer in column_rc_layer:
                    if object.ObjectName == "AcDbPolyline":
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_section_list.append((coor1, coor2))
                continue
            except Exception as ex:
                error(ex)
                error_count += 1
                object_list.append(object)
                time.sleep(2)
        coor_to_col_line_list.sort(key=lambda x: x[0])
        coor_to_floor_line_list.sort(key=lambda x: x[0])
        # except Exception as e:
        #     error_count += 1
        #     time.sleep(5)
        #     error(f'read_col error in step 7: {e}, error_count = {error_count}.')
    # if multi:
    #     return{'coor_to_col_set':coor_to_col_set,
    #             'coor_to_size_set':coor_to_size_set,
    #             'coor_to_floor_set': coor_to_floor_set,
    #             'coor_to_col_line_list':coor_to_col_line_list,
    #             'coor_to_floor_line_list':coor_to_floor_line_list,
    #             'coor_to_rebar_text_list':coor_to_rebar_text_list,
    #             'coor_to_rebar_list':coor_to_rebar_list,
    #             'coor_to_tie_text_list':coor_to_tie_text_list,
    #             'coor_to_tie_list':coor_to_tie_list,
    #             'coor_to_section_list':coor_to_section_list
    #         }
    save_temp_file.save_pkl({'coor_to_col_set': coor_to_col_set,
                             'coor_to_size_set': coor_to_size_set,
                             'coor_to_floor_set': coor_to_floor_set,
                             'coor_to_col_line_list': coor_to_col_line_list,
                             'coor_to_floor_line_list': coor_to_floor_line_list,
                             'coor_to_rebar_text_list': coor_to_rebar_text_list,
                             'coor_to_rebar_list': coor_to_rebar_list,
                             'coor_to_tie_text_list': coor_to_tie_text_list,
                             'coor_to_tie_list': coor_to_tie_list,
                             'coor_to_section_list': coor_to_section_list,
                             'coor_to_floor_slash_set': coor_to_floor_slash_set
                             }, temp_file)
    try:
        doc_column.Close(SaveChanges=False)
    except:
        error('Cant Close Dwg File')


def regex_col_name(coor_to_col_set: set):
    for coor in list(coor_to_col_set):
        text_coor = coor[0]
        txt = coor[1]
        match_obj = re.search(slash_pattern, txt)
        if match_obj:
            suffix_index = re.search(
                r'(\D*)\d+(\D*)', match_obj.group(1))
            first_column = re.findall(
                r'\d+', match_obj.group(1))
            last_column = re.findall(
                r'\d+', match_obj.group(2))
            if first_column and last_column:
                coor_to_col_set.remove(coor)
                for column_number in range(int(first_column[0]), int(last_column[0]) + 1):
                    temp_string = f'{suffix_index.group(1)}{column_number}{suffix_index.group(2)}'
                    coor_to_col_set.add(
                        (text_coor, temp_string))

        elif re.search(commom_pattern, txt):
            coor_to_col_set.remove(coor)
            sep = re.search(
                commom_pattern, txt).group(1)
            for column_text in txt.split(sep):
                coor_to_col_set.add(
                    (text_coor, column_text))


def cal_column_rebar(data={},
                     rebar_excel_path='',
                     line_order=1,
                     size_type: Literal['text', 'section'] = 'text',
                     **kwargs):
    # output_txt =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar.txt')
    # output_txt_2 =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar_floor.txt')
    # excel_filename = (
    #     f'{output_folder}/'
    #     f'{project_name}_'
    #     f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
    #     f'Count.xlsx'
    # )
    if not data:
        return
    coor_to_col_set = data['coor_to_col_set']
    coor_to_size_set = data['coor_to_size_set']
    coor_to_floor_set = data['coor_to_floor_set']
    coor_to_col_line_list = data['coor_to_col_line_list']
    coor_to_floor_line_list = data['coor_to_floor_line_list']
    coor_to_rebar_text_list = data['coor_to_rebar_text_list']
    coor_to_rebar_list = data['coor_to_rebar_list']
    coor_to_tie_text_list = data['coor_to_tie_text_list']
    coor_to_tie_list = data['coor_to_tie_list']
    coor_to_section_list = data['coor_to_section_list']
    coor_to_floor_slash_set = set()
    if 'coor_to_floor_slash_set' in data:
        coor_to_floor_slash_set = data['coor_to_floor_slash_set']
    readRebarExcel(file_path=rebar_excel_path)
    progress('結合格線與柱編號')

    regex_col_name(coor_to_col_set)
    new_coor_to_col_line_list = concat_col_to_grid(
        coor_to_col_set=coor_to_col_set, coor_to_col_line_list=coor_to_col_line_list)

    progress('結合格線與樓層編號')
    parameter_df: pd.DataFrame = read_parameter_df(
        rebar_excel_path, '柱參數表')
    parameter_df.set_index(['樓層'], inplace=True)
    floor_seq_list = list(parameter_df.index)
    new_coor_to_floor_line_list = concat_floor_to_grid(coor_to_floor_set=coor_to_floor_set,
                                                       coor_to_floor_line_list=coor_to_floor_line_list,
                                                       coor_to_floor_slash_set=coor_to_floor_slash_set,
                                                       floor_list=floor_seq_list,
                                                       line_order=line_order)
    # draw_grid_line(new_coor_to_floor_line_list=new_coor_to_floor_line_list,new_coor_to_col_line_list=new_coor_to_col_line_list,msp_beam=msp_column,doc_beam=doc_column)

    if size_type == 'text':
        output_column_list = concat_name_to_col_floor(coor_to_size_set=coor_to_size_set,
                                                      new_coor_to_col_line_list=new_coor_to_col_line_list,
                                                      new_coor_to_floor_line_list=new_coor_to_floor_line_list)
    progress('獲取斷面大小資訊')
    if size_type == 'section':
        output_column_list = get_size_from_section(new_coor_to_col_line_list=new_coor_to_col_line_list,
                                                   new_coor_to_floor_line_list=new_coor_to_floor_line_list,
                                                   coor_to_section_list=coor_to_section_list,
                                                   coor_to_size_set=coor_to_size_set)
    progress('結合柱編號與柱主筋')

    combine_col_rebar(column_list=output_column_list, coor_to_rebar_list=coor_to_rebar_list,
                      coor_to_rebar_text_list=coor_to_rebar_text_list)
    progress('結合柱編號與柱箍筋')
    combine_col_tie(column_list=output_column_list,
                    coor_to_tie_list=coor_to_tie_list,
                    coor_to_tie_text_list=coor_to_tie_text_list,
                    **kwargs)
    return output_column_list


def output_grid_dwg(data, msp_column, doc_column):
    coor_to_col_set = data['coor_to_col_set']
    coor_to_floor_set = data['coor_to_floor_set']
    coor_to_col_line_list = data['coor_to_col_line_list']
    coor_to_floor_line_list = data['coor_to_floor_line_list']
    new_coor_to_col_line_list = concat_col_to_grid(
        coor_to_col_set=coor_to_col_set, coor_to_col_line_list=coor_to_col_line_list)
    new_coor_to_floor_line_list = concat_floor_to_grid(
        coor_to_floor_set=coor_to_floor_set, coor_to_floor_line_list=coor_to_floor_line_list)
    draw_grid_line(new_coor_to_floor_line_list=new_coor_to_floor_line_list,
                   new_coor_to_col_line_list=new_coor_to_col_line_list,
                   msp_beam=msp_column,
                   doc_beam=doc_column)


def cal_column_in_plan(column_list: list[Column],
                       plan_filename: str,
                       plan_layer_config: dict,
                       plan_pkl: str = ''):
    plan_floor_count = sort_plan_count(plan_filename=plan_filename,
                                       layer_config=plan_layer_config,
                                       plan_pkl=plan_pkl)
    for column in column_list:
        if column.floor in plan_floor_count:
            if column.serial in plan_floor_count[column.floor]:
                column.plan_count = plan_floor_count[column.floor][column.serial]
                continue
        column.plan_count = 1


def modify_column_measure(column_list: list[Column]):
    for column in column_list:
        column.x_size /= 10
        column.y_size /= 10
        for tie in column.tie:
            tie.change_spacing(tie.spacing / 10)


def create_report(output_column_list: list[Column],
                  floor_parameter_xlsx='',
                  output_folder='',
                  project_name='',
                  plan_filename='',
                  plan_layer_config=None,
                  measure_type: str = 'cm',
                  plan_pkl: str = ''):
    excel_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Count.xlsx'
    )
    pdf_report = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'柱_report.pdf'
    )

    pdf_report_appendix = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'柱_appendix.pdf'
    )
    if plan_filename or plan_pkl:
        cal_column_in_plan(column_list=output_column_list,
                           plan_filename=plan_filename,
                           plan_layer_config=plan_layer_config,
                           plan_pkl=plan_pkl)

    floor_list = floor_parameter(column_list=output_column_list,
                                 floor_parameter_xlsx=floor_parameter_xlsx)
    progress('整理樓層與柱編號')
    sort_floor_column(floor_list=floor_list, column_list=output_column_list)
    progress('統計樓層柱鋼筋')
    rebar_df, concrete_df, coupler_df, formwork_df, detail_report = summary_floor_rebar(
        floor_list=floor_list, item_type='column', measure_type=measure_type)
    progress('讀取SCAN項目')
    cs_list = create_column_scan()
    progress('柱斷面檢核')
    scan_df = column_check(column_list=output_column_list,
                           column_scan_list=cs_list)
    progress('統整柱鋼筋比')
    header_list, ratio_dict = summary_floor_column_rebar_ratio(
        floor_list=floor_list)
    progress('產生柱EXCEL報表')
    column_df = output_col_excel(column_list=output_column_list,
                                 output_folder=output_folder,
                                 project_name=project_name)
    progress('產生柱詳細檢核報表')
    ng_df = output_detail_scan_report(column_list=output_column_list)
    column_ng_df, sum_df = output_ng_ratio(scan_df)
    progress('產生報表')
    OutputExcel(df_list=[scan_df],
                file_path=excel_filename,
                sheet_name='柱檢核表',
                auto_fit_columns=[1],
                auto_fit_rows=[1],
                columns_list=range(2, len(scan_df.columns)+2), rows_list=range(2, len(scan_df.index)+2))
    OutputExcel(df_list=[rebar_df],
                file_path=excel_filename, sheet_name='鋼筋統計表')
    OutputExcel(df_list=[concrete_df],
                file_path=excel_filename, sheet_name='混凝土統計表')
    OutputExcel(df_list=[coupler_df],
                file_path=excel_filename, sheet_name='續接器統計表')
    OutputExcel(df_list=[column_df],
                file_path=excel_filename, sheet_name='柱統計表')
    OutputExcel(df_list=[formwork_df],
                file_path=excel_filename, sheet_name='模板統計表')
    OutputExcel(df_list=[ng_df], file_path=excel_filename, sheet_name='詳細檢核表')
    progress('產生PDF報表')
    create_scan_pdf(scan_list=cs_list,
                    scan_df=ng_df.copy(),
                    concrete_df=concrete_df,
                    formwork_df=formwork_df,
                    rebar_df=rebar_df,
                    ng_sum_df=sum_df,
                    beam_ng_df=column_ng_df,
                    project_prop={
                        "專案名稱:": f'{project_name}_柱',
                        "測試日期:": time.strftime("%Y/%m/%d %H:%M:%S", time.localtime())
                    },
                    pdf_filename=pdf_report,
                    header_list=header_list,
                    ratio_dict=ratio_dict,
                    report_type='column',
                    item_name='柱',
                    detail_report=detail_report,
                    appendix=pdf_report_appendix)
    return excel_filename, pdf_report, pdf_report_appendix


def create_column_joint_report(column_beam_df: pd.DataFrame, column_list: list[Column], beam_list: list[Beam]):
    excel_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Joint.xlsx'
    )
    if column_beam_df and beam_list:
        joint_df = calculate_column_beam_joint_shear(
            column_list=column_list, beam_list=beam_list, column_beam_df=column_beam_df)
        OutputExcel(df_list=[joint_df],
                    file_path=excel_filename,
                    sheet_name='梁柱接頭檢核表',
                    auto_fit_columns=[1],
                    auto_fit_rows=[1],
                    columns_list=range(2, len(joint_df.columns)+2), rows_list=range(2, len(joint_df.index)+2))
    return excel_filename


def concat_grid_line(line_list: list, start_line: list, overlap: function):
    while True:
        temp_coor = (start_line[1], start_line[2])
        temp_line_list = [l for l in line_list if l[0]
                          == start_line[0] and overlap(l, start_line)]
        if len(temp_line_list) == 0:
            break
        new_line_top = max(temp_line_list, key=lambda l: l[2])[2]
        new_line_bot = min(temp_line_list, key=lambda l: l[1])[1]
        if start_line[2] == new_line_top and start_line[1] == new_line_bot:
            break
        # start_line[2] = new_line_top
        # start_line[1] = new_line_bot
        start_line = (start_line[0], new_line_bot, new_line_top)
    return start_line


def concat_col_to_grid(coor_to_col_set: set, coor_to_col_line_list: list):
    def _overlap(l1, l2):
        return (l2[1] - l1[2])*(l2[2] - l1[1]) <= 0
    # Step 8. 完成col_to_line_set 格式:(col, left, right, up)
    # coor_to_col_set:((coor1,coor2),string)
    # coor_to_col_line_list:[(coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1]))] y向格線
    new_coor_to_col_line_list = []
    # coor_to_col_line_list = [l for l in coor_to_col_line_list if (l[1] == )]
    for element in coor_to_col_set:
        coor = element[0]
        col = element[1]
        left_temp_list = [l for l in coor_to_col_line_list if l[1] <=
                          coor[0][1] and coor[0][1] <= l[2] and l[0] <= coor[0][0]]
        right_temp_list = [l for l in coor_to_col_line_list if l[1]
                           <= coor[0][1] and coor[0][1] <= l[2] and l[0] >= coor[0][0]]
        left_closet_line = [0, float("inf"), float("-inf")]
        right_closet_line = [0, float("inf"), float("-inf")]
        if len(left_temp_list) > 0:
            left_closet_line = list(
                min(left_temp_list, key=lambda l: abs(coor[0][0] - l[0])))
            left_closet_line = concat_grid_line(
                coor_to_col_line_list, left_closet_line, _overlap)
        if len(right_temp_list) > 0:
            right_closet_line = list(
                min(right_temp_list, key=lambda l: abs(coor[0][0] - l[0])))
            right_closet_line = concat_grid_line(
                coor_to_col_line_list, right_closet_line, _overlap)
        top = max(left_closet_line[2], right_closet_line[2])
        bot = min(left_closet_line[1], right_closet_line[1])
        new_coor_to_col_line_list.append(
            (col, (left_closet_line[0], right_closet_line[0], bot, top)))
    return new_coor_to_col_line_list


def concat_floor_to_grid(coor_to_floor_set: set,
                         coor_to_floor_line_list: list,
                         coor_to_floor_slash_set: set,
                         floor_list: list = [],
                         line_order: float = 0):
    '''
    line order means the n th (0-base) horz-line in the bottom of the row
    '''
    def _overlap(l1, l2):
        return (l2[1] - l1[2])*(l2[2] - l1[1]) <= 0

    def expand_floors(text):
        floor_float_list = [turn_floor_to_float(floor) for floor in floor_list]
        # not equal to -1000 (FB)
        Bmax = 0  # 地下最深到幾層(不包括FB不包括FB)
        Fmax = 0  # 正常樓最高到幾層
        Rmax = 0  # R開頭最高到幾層(不包括PRF)
        try:
            Bmax = min([f for f in floor_float_list if f < 0 and f != -1000])
        except:
            pass
        try:
            Fmax = max([f for f in floor_float_list if 1000 > f > 0])
        except:
            pass
        # not equal to PRF (FB)
        try:
            Rmax = max([f for f in floor_float_list if f > 1000 and f != 2000])
        except:
            pass
        return turn_floor_to_list(text, Bmax, Fmax=Fmax, Rmax=Rmax)
    new_coor_to_floor_line_list = []
    record_floor_line_dict = defaultdict(list)
    for element in coor_to_floor_set:

        coor = element[0]
        floor = element[1]

        if floor not in floor_list:
            continue
        top_temp_list = [l for l in coor_to_floor_line_list if l[1]
                         <= coor[0][0] and coor[0][0] <= l[2] and l[0] >= coor[0][1]]
        bot_temp_list = [l for l in coor_to_floor_line_list if l[1]
                         <= coor[0][0] and coor[0][0] <= l[2] and l[0] <= coor[0][1]]
        top_closet_line = [0, float("inf"), float("-inf")]
        bot_closet_line = [0, float("inf"), float("-inf")]
        # find closet line
        if line_order:
            if len(top_temp_list) > 0:
                top_closet_line = list(
                    min(top_temp_list, key=lambda l: abs(coor[0][1] - l[0])))
                top_closet_line = concat_grid_line(
                    coor_to_floor_line_list, top_closet_line, _overlap)
            if len(bot_temp_list) > line_order:
                bot_temp_list.sort(key=lambda l: abs(coor[0][1] - l[0]))
                bot_closet_line = bot_temp_list[line_order]
                bot_closet_line = concat_grid_line(
                    coor_to_floor_line_list, bot_closet_line, _overlap)
        else:
            if len(top_temp_list) > 0:
                top_closet_line = list(
                    min(top_temp_list, key=lambda l: abs(coor[0][1] - l[0])))
                top_closet_line = concat_grid_line(
                    coor_to_floor_line_list, top_closet_line, _overlap)
            if len(bot_temp_list) > 0:
                bot_closet_line = list(
                    min(bot_temp_list, key=lambda l: abs(coor[0][1] - l[0])))
                bot_closet_line = concat_grid_line(
                    coor_to_floor_line_list, bot_closet_line, _overlap)
        # find no belone line
        right = max(top_closet_line[2], bot_closet_line[2])
        left = min(top_closet_line[1], bot_closet_line[1])
        new_coor_to_floor_line_list.append(
            (floor, (left, right, bot_closet_line[0], top_closet_line[0])))
        record_floor_line_dict[(
            left, right, bot_closet_line[0], top_closet_line[0])].append(floor)
    if coor_to_floor_slash_set:
        for floor_line, floors in record_floor_line_dict.items():
            slash_list = [(coor, slash) for (coor, slash) in coor_to_floor_slash_set if
                          (floor_line[0] - coor[0][0]) * (floor_line[1] - coor[0][0]) < 0 and
                          (floor_line[2] - coor[0][1]) * (floor_line[3] - coor[0][1]) < 0]
            if not slash_list:
                continue
            # only support for two floors with slash , not ok with four and tw slash
            if len(floors) != 2:
                continue
            new_floors = expand_floors(f'{floors[0]}~{floors[1]}')

            for i in range(1, len(new_floors) - 1):
                new_coor_to_floor_line_list.append(
                    (new_floors[i], floor_line))

    return new_coor_to_floor_line_list


def draw_grid_line(new_coor_to_floor_line_list: list, new_coor_to_col_line_list: list, msp_beam: object, doc_beam: object):
    output_dwg = os.path.join(
        output_folder, f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_Markon.dwg')
    for grid in new_coor_to_floor_line_list:
        coor_list1 = [grid[1][0], grid[1][2], 0, grid[1][1], grid[1]
                      [2], 0, grid[1][1], grid[1][3], 0, grid[1][0], grid[1][3], 0]
        # coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
        points1 = vtFloat(coor_list1)
        line1 = msp_beam.AddPolyline(points1)
        text1 = msp_beam.AddMText(vtPnt(
            (grid[1][0] + grid[1][1])/2 - 25, (grid[1][2] + grid[1][3])/2), 10, grid[0])
        text1.Height = 50
        line1.SetWidth(0, 5, 5)
        line1.color = 101
    for grid in new_coor_to_col_line_list:
        coor_list1 = [grid[1][0], grid[1][2], 0, grid[1][1], grid[1]
                      [2], 0, grid[1][1], grid[1][3], 0, grid[1][0], grid[1][3], 0]
        # coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
        points1 = vtFloat(coor_list1)
        line1 = msp_beam.AddPolyline(points1)
        text1 = msp_beam.AddMText(
            vtPnt((grid[1][0] + grid[1][1])/2, (grid[1][2] + grid[1][3])/2), 10, grid[0])
        text1.Height = 25
        line1.SetWidth(0, 5, 5)
        line1.color = 201
    doc_beam.SaveAs(output_dwg)
    doc_beam.Close(SaveChanges=True)
    # points2 = vtFloat(coor_list2)


def _ingrid(size_coor, grid_coor):
    pt_x = size_coor[0]
    pt_y = size_coor[1]
    if len(grid_coor) == 0:
        return False
    if (pt_x - grid_coor[0])*(pt_x - grid_coor[1]) < 0 and (pt_y - grid_coor[2])*(pt_y - grid_coor[3]) < 0:
        return True
    return False


def concat_name_to_col_floor(coor_to_size_set: set,
                             new_coor_to_col_line_list: list,
                             new_coor_to_floor_line_list: list):

    output_column_list: list[Column]
    output_column_list = []
    for element in coor_to_size_set:
        new_column = Column()
        coor = element[0]
        size = element[1]
        new_column.set_size(size)
        col = [c for c in new_coor_to_col_line_list if _ingrid(
            size_coor=coor[0], grid_coor=c[1])]
        floor = [f for f in new_coor_to_floor_line_list if _ingrid(
            size_coor=coor[0], grid_coor=f[1])]
        if col and floor:
            new_column.set_border(col[0][1], floor[0][1])
            new_column.set_column_border(
                col[0][1][0], floor[0][1], border_type="Table")
        if len(col) > 0:
            new_column.serial = col[0][0]
        if len(floor) > 0:
            new_column.floor = floor[0][0]
        if len(col) > 1:
            progress(
                f'{size}:{coor} => {list(map(lambda c:c[0],col))}')
            new_column.multi_column.extend(list(map(lambda c: c[0], col[1:])))
            # print(f'{size}:{coor} => {list(map(lambda c:c[0],col))}')
        if len(floor) > 1:
            progress(
                f'{size}:{coor} => {list(map(lambda c:c[0],floor))}')
            # print(f'{size}:{coor} => {list(map(lambda c:c[0],floor))}')
            new_column.multi_floor.extend(list(map(lambda c: c[0], floor[1:])))
        if new_column.serial != '':
            output_column_list.append(new_column)
    return output_column_list


def get_size_from_section(new_coor_to_col_line_list: list,
                          new_coor_to_floor_line_list: list,
                          coor_to_section_list: list,
                          coor_to_size_set: set,):
    output_column_list: list[Column]
    output_column_list = []
    size_text = ''
    for coor1, coor2 in coor_to_section_list:
        new_column = Column()
        x_size = abs(coor1[0] - coor2[0])
        y_size = abs(coor1[1] - coor2[1])
        size = f'{x_size}x{y_size}'
        col = [c for c in new_coor_to_col_line_list if _ingrid(
            size_coor=coor1, grid_coor=c[1])]
        floor = [f for f in new_coor_to_floor_line_list if _ingrid(
            size_coor=coor1, grid_coor=f[1])]
        if col and floor:
            new_column.set_border(col[0][1], floor[0][1])
            new_column.set_column_border(coor1, coor2)
            size_text = [
                s for s in coor_to_size_set if new_column.in_grid(coor=s[0][0])]
            if size_text:
                # print(f'CAD:{size}  TEXT:{size_text[0][1]}')
                size = size_text[0][1]
            new_column.set_size(size)
        if len(col) > 0:
            new_column.serial = col[0][0]
        if len(floor) > 0:
            new_column.floor = floor[0][0]
        if len(col) > 1:
            progress(
                f'{size}:{coor1} => {list(map(lambda c:c[0],col))}')
            # print(f'{size}:{coor1} => {list(map(lambda c:c[0],col))}')
            new_column.multi_column.extend(list(map(lambda c: c[0], col[0:])))
        if len(floor) > 1:
            progress(
                f'{size}:{coor1} => {list(map(lambda c:c[0],floor))}')
            # print(f'{size}:{coor1} => {list(map(lambda c:c[0],floor))}')
            new_column.multi_floor.extend(list(map(lambda c: c[0], floor[1:])))
        if len([c for c in output_column_list if c.floor == new_column.floor and c.serial == new_column.serial]) > 0:
            progress(
                f'{new_column.floor}{new_column.serial} is exists')
            # print(f'{new_column.floor}{new_column.serial} is exists')
            continue
        if new_column.serial != '':
            output_column_list.append(new_column)
    return output_column_list


def combine_col_rebar(column_list: list[Column], coor_to_rebar_list: list, coor_to_rebar_text_list: list):
    count = 0
    for coor, rebar_text in [c for c in coor_to_rebar_text_list if '@' not in c[1]]:
        column = [c for c in column_list if c.in_grid(coor=coor[0])]
        if len(column) > 0:
            if (coor[0], rebar_text) in column[0].multi_rebar_text:
                # print(f'{coor[0]}:{rebar_text} is exists')
                continue
            column[0].multi_rebar_text.append((coor[0], rebar_text))
        count += 1
        if count % 100 == 0:
            progress(f'{count} / {len(coor_to_rebar_text_list)}')
            # if column[0].rebar_text == '':
            #     column[0].rebar_text = rebar_text
            #     column[0].rebar_text_coor = coor[0]
            #     column[0].multi_rebar_text.append((coor[0],rebar_text))
            # else:
            #     column[0].multi_rebar_text.append((coor[0],rebar_text))
    count = 0
    progress('組合柱斷面與鋼筋')
    for column in column_list:
        count += 1
        rebars = [
            rebar for rebar in coor_to_rebar_list if column.in_column_section(rebar[0])]
        if count % 100 == 0:
            progress(f'{count} / {len(column_list)}')
        for rebar in rebars:
            column.add_rebar_coor(rebar)
            coor_to_rebar_list.remove(rebar)
    # for rebar in coor_to_rebar_list:
        # column = [c for c in column_list if c.in_column_section(rebar[0])]
        # if len(column) > 0:
        #     column[0].add_rebar_coor(rebar)
    for column in column_list:
        column.sort_rebar()

        # print(f'{column.floor}:{column.serial} x:{column.x_row} y:{column.y_row}')


def combine_col_tie(column_list: list[Column],
                    coor_to_tie_text_list: list,
                    coor_to_tie_list: list,
                    **kwargs):
    count = 0

    def distance(coor1, coor2):
        return ((coor1[0] - coor2[0])**2 + (coor1[1] - coor2[1])**2)**0.5

    progress('組合柱斷面與繫筋')
    # column_list = [c for c in column_list if c.floor ==
    #                '11F' and c.serial == 'C10']
    for column in column_list:
        ties = [tie for tie in coor_to_tie_list if column.in_column_section(
            coor=tie[0]) and column.in_column_section(coor=tie[1]) and distance(*tie) > 40]
        for tie in ties:
            column.add_tie(tie)
            coor_to_tie_list.remove(tie)
        count += 1
        if count % 100 == 0:
            progress(f'{count} / {len(column_list)}')
    # for tie in coor_to_tie_list:
    #     column = [c for c in column_list if c.in_grid(
    #         coor=tie[0]) and c.in_grid(coor=tie[1])]
    #     if len(column) > 0:
    #         column[0].add_tie(tie)
    count = 0
    progress('組合柱斷面與箍筋文字')
    for coor, tie_text in coor_to_tie_text_list:
        column = [c for c in column_list if c.in_grid(
            coor=coor[1]) and c.in_grid(coor=coor[1])]
        if len(column) > 0:
            column[0].add_tie_text(coor=coor, text=tie_text)
        count += 1
        if count % 100 == 0:
            progress(f'{count} / {len(coor_to_tie_text_list)}')
    for column in column_list:
        column.sort_tie(**kwargs)
        # print(f'{column.floor}:{column.serial} x:{column.x_tie} y:{column.y_tie} tie:{column.tie_dict}')


def output_col_excel(column_list: list[Column], output_folder: str, project_name: str):
    header_info_1 = [('樓層', ''), ('柱編號', ''), ('X向 柱寬', 'cm'), ('Y向 柱寬', 'cm')]
    header_rebar = [('柱主筋', '主筋'), ('柱主筋', 'X向支數'), ('柱主筋', 'Y向支數'),
                    ('柱箍筋', '圍束區'), ('柱箍筋', '非圍束區'), ('柱箍筋', 'X向繫筋'), ('柱箍筋', 'Y向繫筋')]
    header_second_rebar = [('次柱主筋', '主筋'), ('次柱主筋', 'X向支數'), ('次柱主筋', 'Y向支數')]
    sorted(column_list, key=lambda c: c.serial)
    header = pd.MultiIndex.from_tuples(header_info_1 + header_rebar)
    column_df = pd.DataFrame(
        np.empty([len(column_list), len(header)], dtype='<U16'), columns=header)
    row = 0
    for c in column_list:
        try:
            if c.serial == '' or c.floor == '':
                continue
            column_df.at[row, ('樓層', '')] = c.floor
            column_df.at[row, ('柱編號', '')] = c.serial
            column_df.at[row, ('X向 柱寬', 'cm')] = c.x_size
            column_df.at[row, ('Y向 柱寬', 'cm')] = c.y_size
            if len(c.total_rebar) > 0:
                column_df.at[row, ('柱主筋', '主筋')] = c.total_rebar[0][0].text
                column_df.at[row, ('柱主筋', 'X向支數')
                             ] = c.x_dict[c.total_rebar[0][0].size]
                column_df.at[row, ('柱主筋', 'Y向支數')
                             ] = c.y_dict[c.total_rebar[0][0].size]
            if len(c.total_rebar) == 2:
                column_df.at[row, ('次柱主筋', '主筋')] = c.total_rebar[1][0].text
                column_df.at[row, ('次柱主筋', 'X向支數')
                             ] = c.x_dict[c.total_rebar[1][0].size]
                column_df.at[row, ('次柱主筋', 'Y向支數')
                             ] = c.y_dict[c.total_rebar[1][0].size]
            if c.tie_dict:
                column_df.at[row, ('柱箍筋', '圍束區')] = str(c.confine_tie)
                column_df.at[row, ('柱箍筋', '非圍束區')] = str(c.middle_tie)
                column_df.at[row, ('柱箍筋', 'X向繫筋')] = c.x_tie
                column_df.at[row, ('柱箍筋', 'Y向繫筋')] = c.y_tie
        except:
            progress(f'{c.floor} {c.serial} 資料有誤')
        row += 1

    # output_column_list = sorted(output_column_list,key=lambda c:c.serial)
    # column_df.sort_values(by=[('柱編號', '')],ascending=True,inplace=True)
    return column_df


def floor_parameter(column_list: list[Column], floor_parameter_xlsx: str):
    floor_list: list[Floor]
    floor_list = []
    parameter_df = read_parameter_df(floor_parameter_xlsx, '柱參數表')
    parameter_df.set_index(['樓層'], inplace=True)
    floor_seq_list = parameter_df.index
    for c in column_list[:]:
        # for floor in c.multi_floor:
        for column_name in c.multi_column:
            if c.serial == column_name:
                continue
            new_c = copy.deepcopy(c)
            # new_c.floor = floor
            new_c.serial = column_name
            # new_c.multi_floor = []
            column_list.append(new_c)
    for c in column_list[:]:
        for floor in c.multi_floor:
            if c.floor == floor:
                continue
            new_c = copy.deepcopy(c)
            new_c.floor = floor
            # new_c.serial = column_name
            column_list.append(new_c)

    for floor_name in parameter_df.index:
        temp_floor = Floor(str(floor_name))
        floor_list.append(temp_floor)
        temp_floor.set_column_prop(parameter_df.loc[floor_name])
        temp_floor.add_column(
            [c for c in column_list if c.floor == temp_floor.floor_name])

    return floor_list


def sort_floor_column(floor_list: list[Floor], column_list: list[Column]):
    def match_column(col: Column, col_list: list[Column], pos: str):
        temp_list = [c for c in col_list if c.serial == col.serial]
        if temp_list and pos == 'up':
            col.up_column = temp_list[0]
        if temp_list and pos == 'bot':
            col.bot_column = temp_list[0]
    floor_seq = list(map(lambda f: f.floor_name, floor_list))
    list(map(lambda c: c.set_seq(floor_seq), column_list))
    for i in range(0, len(floor_list) - 1):
        temp_list = floor_list[i].column_list
        bot_list = floor_list[i + 1].column_list
        list(map(lambda c: match_column(c, bot_list, 'bot'), temp_list))
    for i in range(1, len(floor_list)):
        temp_list = floor_list[i].column_list
        up_list = floor_list[i - 1].column_list
        list(map(lambda c: match_column(c, up_list, 'up'), temp_list))
    column_list.sort(key=lambda c: (c.serial, c.seq))


def count_column_multiprocessing(column_filenames: list[str],
                                 layer_config: dict,
                                 temp_file: list[str],
                                 output_folder='',
                                 project_name='',
                                 template_name='',
                                 floor_parameter_xlsx='',
                                 progress_file='',
                                 plan_filename='',
                                 plan_layer_config=None,
                                 client_id="temp",
                                 measure_type="cm"):
    def read_col_multi(column_filename, temp_file):
        msp_column, doc_column = read_column_cad(
            column_filename=column_filename)
        sort_col_cad(msp_column=msp_column,
                     doc_column=doc_column,
                     layer_config=layer_config,
                     temp_file=temp_file,
                     progress_file=progress_file)
        output_column_list = cal_column_rebar(data=save_temp_file.read_temp(temp_file),
                                              rebar_excel_path=floor_parameter_xlsx,
                                              progress_file=progress_file)
        return output_column_list
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    start = time.time()  # 開始測量執行時間
    with Pool() as p:
        jobs = []
        column_list = []
        for i, filename in enumerate(column_filenames):
            temp_new = os.path.splitext(temp_file)[0]
            column_temp = f'{temp_new}-{i}.pkl'
            jobs.append(p.apply_async(read_col_multi, (filename, column_temp)))
        for job in jobs:
            output_column_list = job.get()
            column_list.extend(output_column_list)
    save_temp_file.save_pkl(output_column_list, f'{temp_new}-column_list.pkl')
    excel_filename, pdf_report = create_report(output_column_list=column_list,
                                               floor_parameter_xlsx=floor_parameter_xlsx,
                                               output_folder=output_folder,
                                               project_name=project_name,
                                               progress_file=progress_file,
                                               plan_filename=plan_filename,
                                               plan_layer_config=plan_layer_config,
                                               measure_type=measure_type)

    end = time.time()
    print("執行時間：%f 秒" % (end - start))
    return os.path.basename(excel_filename), os.path.basename(pdf_report), f'{temp_new}-column_list.pkl'


def count_column_multifiles(project_name: str,
                            column_filenames: list[str],
                            floor_parameter_xlsx: str,
                            output_folder: str,
                            pkl_file_folder: str,
                            layer_config,
                            line_order,
                            size_type,
                            plan_pkl: str = '',
                            plan_layer_config: str = '',
                            plan_filename: str = '',
                            measure_type="cm",
                            client_id="temp",
                            **kwargs):

    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)

    all_column_list = []
    excel_filename = ''
    output_file_list = []

    now_time = time.strftime("%Y%m%d_%H%M%S")

    if plan_pkl:
        plan = {'plan_pkl': plan_pkl}
    elif plan_filename and plan_layer_config:
        plan = {'plan_filename': plan_filename,
                'plan_layer_config': plan_layer_config}
    else:
        plan = {}

    if kwargs.get('column_pkl', ''):
        all_column_list = save_temp_file.read_temp(
            kwargs.get('column_pkl', ""))

        result_pkl = kwargs.get('column_pkl', '')
    else:
        if kwargs.get('pkl', []):
            for i, filename in enumerate(kwargs.get('pkl', [])):
                try:
                    column_list = cal_column_rebar(data=save_temp_file.read_temp(filename),
                                                   rebar_excel_path=floor_parameter_xlsx,
                                                   line_order=line_order,
                                                   size_type=size_type,
                                                   **kwargs)

                    all_column_list.extend(column_list)
                except Exception as ex:
                    print(f'{filename} error {ex}')
                    save_temp_file.save_pkl(
                        all_column_list, f'{pkl_file_folder}/{project_name}-{now_time}-column-object-{i}.pkl')
        else:
            for i, filename in enumerate(column_filenames):

                base_filename_without_ext = os.path.splitext(
                    os.path.basename(filename))[0]

                try:

                    tmp_file = f'{pkl_file_folder}/{project_name}-{now_time}-{base_filename_without_ext}-column-data-{i}.pkl'

                    if not os.path.exists(tmp_file):
                        msp_column, doc_column = read_column_cad(filename)
                        sort_col_cad(msp_column=msp_column,
                                     doc_column=doc_column,
                                     layer_config=layer_config,
                                     temp_file=tmp_file)

                    column_list = cal_column_rebar(data=save_temp_file.read_temp(tmp_file),
                                                   rebar_excel_path=floor_parameter_xlsx,
                                                   line_order=line_order,
                                                   size_type=size_type,
                                                   **kwargs)

                    all_column_list.extend(column_list)

                except Exception as ex:
                    print(f'{filename} error {ex}')
                    save_temp_file.save_pkl(
                        all_column_list, f'{pkl_file_folder}/{project_name}-{now_time}-column-object-{i}.pkl')

        result_pkl = f'{pkl_file_folder}/{project_name}-{now_time}-column-object-all.pkl'

        save_temp_file.save_pkl(
            all_column_list, tmp_file=result_pkl)

    if all_column_list:
        excel_filename, pdf_report, pdf_report_appendix = create_report(output_column_list=all_column_list,
                                                                        output_folder=output_folder,
                                                                        project_name=project_name,
                                                                        floor_parameter_xlsx=floor_parameter_xlsx,
                                                                        measure_type=measure_type,
                                                                        **plan)

        return [os.path.basename(excel_filename), os.path.basename(pdf_report), os.path.basename(pdf_report_appendix)], result_pkl

# def count_column_main(column_filename, layer_config, temp_file='temp_1221_1F.pkl', output_folder='', project_name='', template_name='', floor_parameter_xlsx=''):
#     start = time.time()
#     msp_column, doc_column = read_column_cad(column_filename=column_filename)
#     sort_col_cad(msp_beam=msp_column, doc_column=doc_column,
#                  layer_config=layer_config, temp_file=temp_file)
#     output_column_list = cal_column_rebar(
#         data=save_temp_file.read_temp(temp_file))
#     output_excel = create_report(output_column_list=output_column_list, output_folder=output_folder,
#                                  project_name=project_name, floor_parameter_xlsx=floor_parameter_xlsx)
#     # output_dwg = draw_rebar_line(class_beam_list=class_beam_list,msp_beam=msp_column,doc_beam=doc_column,output_folder=output_folder,project_name=project_name)
#     print(f'Total Time:{time.time() - start}')
#     return os.path.basename(output_excel)


if __name__ == '__main__':
    from os import listdir
    from os.path import isfile, join

    parameter = read_parameter_json('Elements')['column']
    # parameter['']
    count_column_multifiles(
        project_name='Test',
        column_filenames=[
            # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-002_柱配筋圖-1.dwg',
            # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-003_柱配筋圖-2.dwg',
            # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-004_柱配筋圖-3.dwg',
            # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-005_柱配筋圖-4.dwg',
            # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-006_柱配筋圖-5.dwg',
            r'D:\Desktop\BeamQC\TEST\2025-0113\富樂群-2025-01-13-10-44-XS-COL.dwg'],
        # r'D:\Desktop\BeamQC\TEST\2024-1021\柱\S3-008_柱配筋圖-7.dwg'],
        floor_parameter_xlsx=r'TEST\2025-0113\富樂群-2025-01-13-09-39-floor.xlsx',
        output_folder=r'D:\Desktop\BeamQC\TEST\2025-0113',
        pkl_file_folder=r'D:\Desktop\BeamQC\TEST\2025-0113',
        # pkl=[r'TEST\2025-0113\Test-20250113_133819-富樂群-2025-01-13-10-44-XS-COL-column-data-0.pkl'],
        **parameter
    )
    # sys.argv[1] # XS-COL的路徑
    # col_filename = r'D:\Desktop\BeamQC\TEST\2024-0830\柱\11002_S3101_C1棟柱配筋圖.dwg'
    # column_filenames = [
    #     # sys.argv[1] # XS-COL的路徑
    #     r'D:\Desktop\BeamQC\TEST\2024-0830\柱\11002_S3109_C1棟柱配筋圖.dwg',
    #     # r'D:\Desktop\BeamQC\TEST\2023-0324\岡山\XS-COL(南基地).dwg',#sys.argv[1] # XS-COL的路徑
    #     # r'D:\Desktop\BeamQC\TEST\INPUT\1-2023-02-15-15-23--XS-COL-3.dwg',#sys.argv[1] # XS-COL的路徑
    #     # r'D:\Desktop\BeamQC\TEST\INPUT\1-2023-02-15-15-23--XS-COL-4.dwg'#sys.argv[1] # XS-COL的路徑
    # ]
    # floor_parameter_xlsx = r'D:\Desktop\BeamQC\TEST\2024-0923\P2022-04A 國安社宅二期暨三期22FB4-2024-09-24-16-02-floor_1.xlsx'
    # output_folder = r'D:\Desktop\BeamQC\TEST\2024-0923'
    # project_name = '2024-0924'
    # plan_filename = r'D:\Desktop\BeamQC\TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-XS-PLAN.dwg'
    # plan_layer_config = {
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],
    #     'name_text_layer': ['S-TEXTG', 'S-TEXTB', 'S-TEXTC'],
    #     'floor_text_layer': ['S-TITLE']
    # }
    # plan_layer_config = {
    #     'block_layer': ['DwFm'],
    #     'name_text_layer': ['BTXT', 'CTXT', 'BTXT_S_'],
    #     'floor_text_layer': ['TEXT1']
    # }
    # layer_config = {
    #     'text_layer': ['TABLE', 'SIZE'],
    #     'line_layer': ['TABLE'],
    #     'rebar_text_layer': ['NBAR'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['RBAR'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['NBAR'],  # 箍筋文字圖層
    #     'tie_layer': ['RBAR'],  # 箍筋文字圖層
    #     'block_layer': ['DwFm'],  # 框框的圖層
    #     'column_rc_layer': ['OLINE']  # 斷面圖層
    # }
    # DrawRC
    # entity_type = {
    #     'rebar_layer': ['AcDbPolyline'],
    #     'rebar_data_layer': ['AcDbMText'],
    #     'rebar_data_leader_layer': ['AcDbLeader'],
    #     'tie_text_layer': ['AcDbText']
    # }
    # RCAD
    # layer_config = {
    #     'text_layer': ['文字-柱線名稱', '文字-樓群名稱', '文字-斷面尺寸'],
    #     'line_layer': ['GirdInner', 'GirdBoundary'],
    #     'rebar_text_layer': ['文字-主筋根數'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['主筋斷面', '#8', '#10'],  # 鋼筋和箍筋的線的塗層
    #     # 箍筋文字圖層
    #     'tie_text_layer': ['文字-剪力筋 中央區', '文字-剪力筋-BC', '文字-剪力筋-圍束區', '文字-剪力筋'],
    #     'tie_layer': ['箍筋線'],  # 箍筋文字圖層
    #     'block_layer': ['0', '_noprint'],  # 框框的圖層
    #     'column_rc_layer': ['柱斷面線']  # 斷面圖層
    # }
    # Elements
    # layer_config = {
    #     'text_layer': ['S-TEXT'],
    #     'line_layer': ['S-TABLE'],
    #     'rebar_text_layer': ['S-TEXT'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['S-REINFD'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['S-TEXT'],  # 箍筋文字圖層
    #     'tie_layer': ['S-REINF'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],  # 框框的圖層
    #     'column_rc_layer': ['S-RC']  # 斷面圖層
    # }
    # JJP
    # layer_config = {
    #     'text_layer': ['TITLE', 'BARNOTE'],
    #     'line_layer': ['TABLE'],
    #     'rebar_text_layer': ['BARNOTE'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['BARA'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['BARNOTE'],  # 箍筋文字圖層
    #     'tie_layer': ['BARS'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS', 'XREF'],  # 框框的圖層
    #     'column_rc_layer': ['OLINE'],  # 斷面圖層
    #     'burst_layer_list': ['XREF']
    # }
    # 永峻
    # layer_config = {
    #     'text_layer': ['TEXT-1', 'CT-1', '手改'],
    #     'line_layer': ['CT-1'],
    #     'rebar_text_layer': ['TEXT-1', '手改'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['CT-3', '手改'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['TEXT-1'],  # 箍筋文字圖層
    #     'tie_layer': ['CT-2', '手改'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS', 'XREF'],  # 框框的圖層
    #     'column_rc_layer': ['CT-1'],  # 斷面圖層
    #     'burst_layer_list': ['']
    # }
    # main_logger = setup_custom_logger(__name__, client_id=project_name)
    # msp_column = None
    # doc_column = None
    # all_column_list = []

    # mypath = r'D:\Desktop\BeamQC\TEST\2024-0830\柱'
    # onlyfiles = [os.path.join(mypath, f) for f in listdir(mypath) if isfile(
    #     join(mypath, f)) and os.path.splitext(f)[1] == ".dwg"]

    # tmp_file = f'{output_folder}/P2022-04A 國安社宅二期暨三期22FB4-2024-09-24-16-02-temp-0.pkl'

    # column_list = cal_column_rebar(data=save_temp_file.read_temp(tmp_file),
    #                                rebar_excel_path=floor_parameter_xlsx)

    # for i, filename in enumerate(onlyfiles):
    #     try:
    #         print(f'run {filename}')

    #         tmp_file = f'{output_folder}/0904-column-test-{i}.pkl'

    #         if not os.path.exists(tmp_file):
    #             msp_column, doc_column = read_column_cad(filename)
    #             sort_col_cad(msp_column=msp_column,
    #                          doc_column=doc_column,
    #                          layer_config=layer_config,
    #                          temp_file=tmp_file,
    #                          progress_file=r'result\tmp')

    #         column_list = cal_column_rebar(data=save_temp_file.read_temp(tmp_file),
    #                                        rebar_excel_path=floor_parameter_xlsx,
    #                                        progress_file=r'result\tmp')

    #         all_column_list.extend(column_list)

    #     except Exception:
    #         print(f'{filename} error')
    #         save_temp_file.save_pkl(
    #             all_column_list, f'{output_folder}/column_list-temp-{i}.pkl')

    # save_temp_file.save_pkl(
    #     column_list, f'{output_folder}/column-3.pkl')

    # output_grid_dwg(data=save_temp_file.read_temp(r'D:\Desktop\BeamQC\TEST\2024-0819\column.pkl'),
    #                 msp_column=msp_column,
    #                 doc_column=doc_column)
    # print(save_temp_file.read_temp(
    #     r'D:\Desktop\BeamQC\TEST\INPUT\test-2023-02-15-15-41-temp-0.pkl'))
    # temp = []
    # for i in range(0, 12):
    #     column_list = cal_column_rebar(data=save_temp_file.read_temp(f'D:/Desktop/BeamQC/TEST/2024-0829/柱/column-data-{i}.pkl'),
    #                                    rebar_excel_path=floor_parameter_xlsx,
    #                                    progress_file=r'result\tmp')
    #     temp.extend(column_list)

    # column_list = save_temp_file.read_temp(
    #     f'{output_folder}/column-3.pkl')
    # save_temp_file.save_pkl(
    #     temp, r'D:\Desktop\BeamQC\TEST\2024-0829\柱\column_list-2.pkl')

    # create_report(output_column_list=column_list,
    #               output_folder=output_folder,
    #               project_name=project_name,
    #               floor_parameter_xlsx=floor_parameter_xlsx,
    #               progress_file=r'result\tmp',
    #               plan_pkl=r'D:\Desktop\BeamQC\TEST\2024-0923\P2022-04A 國安社宅二期暨三期22FB4-2024-09-23-11-32-XS-PLAN_plan_count_set.pkl',
    #               plan_layer_config=plan_layer_config,
    #               measure_type='cm')

    # count_column_multiprocessing(column_filenames=onlyfiles,
    #                              layer_config=layer_config,
    #                              temp_file=r'TEST\2024-0829\柱\column-data.pkl',
    #                              output_folder=output_folder,
    #                              project_name=project_name,
    #                              floor_parameter_xlsx=floor_parameter_xlsx,
    #                              measure_type='mm')
