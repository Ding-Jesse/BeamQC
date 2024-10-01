import win32com.client
import pythoncom
import re
import time
from math import inf
import multiprocessing
import os
import pandas as pd
import sys
import src.save_temp_file as save_temp_file
from typing import Literal
from gzip import READ
from multiprocessing.spawn import prepare
from tkinter import HIDDEN
from numpy import Inf, object_
from openpyxl import load_workbook
from itertools import groupby
from functools import cmp_to_key
from src.plan_to_beam import turn_floor_to_float, turn_floor_to_string, turn_floor_to_list, floor_exist, vtFloat, mycmp
from src.logger import setup_custom_logger
from collections import defaultdict
from utils.algorithm import match_points, convert_mm_to_cm, find_all_matching_patterns, extract_dimensions
weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡', '、']
special_char_pattern = r'[、|,]'
global main_logger
# def read_plan(plan_filename, floor_layer, col_layer, block_layer, result_filename, progress_file):


def error(error_message):
    '''
    把錯誤訊息印到error.log裡面
    '''
    global main_logger
    main_logger.error(error_message)


def progress(message):
    '''把進度印到progress裡面，在app.py會對這個檔案做事'''
    global main_logger
    main_logger.info(message)


def get_distance(coor1, coor2):
    if isinstance(coor1[0], tuple) and isinstance(coor2[0], tuple):
        return abs(coor1[0][0]-coor2[0][0]) + abs(coor1[0][1]-coor2[0][1])
    if isinstance(coor1, tuple) and isinstance(coor2, tuple):
        return abs(coor1[0]-coor2[0]) + abs(coor1[1]-coor2[1])
    return 10000


def read_plan(plan_filename: str, layer_config: dict):
    def _cal_ratio(pt1, pt2):
        if abs(pt1[1]-pt2[1]) == 0:
            return 1000
        return abs(pt1[0]-pt2[0])/abs(pt1[1]-pt2[1])
    floor_layer = layer_config['floor_layer']
    col_layer = layer_config['col_layer']
    block_layer = layer_config['block_layer']
    column_block_layer = layer_config['column_block_layer']
    size_layer = layer_config['size_layer']
    table_line_layer = layer_config['table_line_layer']
    column_block_object_type = ['AcDbBlockReference', "AcDbPolyline"]
    error_count = 0
    text_entity_name = ["AcDbText", 'AcDbAttribute']
    progress('開始讀取平面圖(核對項目: 柱配筋對應)')
    # Step 1. 打開應用程式
    wincad_plan = None
    while wincad_plan is None and error_count <= 10:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 1: {ex}, error_count = {error_count}.')
    progress('平面圖讀取進度 1/11')

    # Step 2. 匯入檔案
    doc_plan = None
    while doc_plan is None and wincad_plan and error_count <= 10:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)

        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 2: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 2/11')

    # Step 3. 匯入modelspace
    msp_plan = None
    while msp_plan is None and doc_plan and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 3: {ex}, error_count = {error_count}.')
    progress('平面圖讀取進度 3/11')

    # Step 4 解鎖所有圖層 -> 不然不能刪東西
    flag = 0
    while not flag and error_count <= 10:
        try:
            layer_count = doc_plan.Layers.count
            for x in range(layer_count):
                layer = doc_plan.Layers.Item(x)
                layer.Lock = False
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 4: {ex}, error_count = {error_count}.')
    progress('平面圖讀取進度 4/11')

    # Step 7. 遍歷所有物件 -> 完成 coor_to_floor_set, coor_to_col_set, block_coor_list
    coor_to_floor_set = set()  # set (字串的coor, floor)
    coor_to_col_set = set()  # set (coor, col)
    coor_to_size_set = set()  # set (coor, size)
    table_to_size_set = []  # set (coor, size)
    block_coor_list = []  # 存取方框最左下角的點座標
    table_line_list = []
    column_block_entity = []
    progress('正在遍歷平面圖上所有物件並篩選出有效信息，運行時間取決於平面圖大小，請耐心等候...')
    used_layer_list = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name
    total = msp_plan.Count
    progress(
        f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
    count = 0
    for msp_object in msp_plan:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'平面圖已讀取{count}/{total}個物件')
        while error_count <= 3 and not object_list:
            try:
                if msp_object.Layer not in used_layer_list:
                    break
                # print(f'{msp_object.Layer}:{msp_object.EntityName}')
                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference":
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read_plan error in step 7-1: {ex}, error_count = {error_count}.')

        while error_count <= 3 and object_list:
            count += 1
            if count % 1000 == 0:
                progress(f'平面圖已讀取{count}/{total}個物件')
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer
                # 取floor的字串 -> 抓括號內的字串 (Ex. '十層至十四層結構平面圖(10F~14F)' -> '10F~14F')
                # 若此處報錯，可能原因: 1. 沒有括號, 2. 待補
                # if object.Layer == floor_layer and object.ObjectName == "AcDbText":
                #     print(object.TextString)
                #  and object.InsertionPoint[1] >= 0
                if object_layer in floor_layer and object.ObjectName in text_entity_name and '(' in object.TextString:
                    floor = object.TextString
                    floor = re.search(r'\(([^)]+)', floor).group(1)  # 取括號內的樓層數
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))  # 不取概數的話後面抓座標會出問題，例如兩個樓層在同一格
                    no_chinese = False
                    for ch in floor:  # 待修正
                        if ch == 'F' or ch.isdigit():
                            no_chinese = True
                            break
                    if floor != '' and no_chinese:
                        coor_to_floor_set.add((coor, floor))
                    else:
                        error(
                            'read_plan error in step 7: floor is an empty string. ')
                # 取col的字串
                if object_layer in col_layer and \
                    object.ObjectName in text_entity_name and \
                        object.TextString != '' and \
                (object.TextString[0] == 'C' or (('¡æ' in object.TextString or '⊥' in object.TextString) and 'C' in object.TextString)) and \
                'S' not in object.TextString:
                    col = f"C{object.TextString.split('C')[1].strip()}"
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    coor_to_col_set.add(((coor1, coor2), col))
                    # print(f'{(coor1, coor2)},{col}')
                # 取size
                if object_layer in col_layer and object.ObjectName in text_entity_name and '(' in object.TextString:
                    if re.match(r'\((.+)[X|x](.+)\)', object.TextString):
                        size_group = re.search(
                            r'\((.+)[X|x](.+)\)', object.TextString)
                        size = f'{size_group.group(1)}x{size_group.group(2)}'
                    else:
                        size = (object.TextString.split(
                            '(')[1]).split(')')[0]  # 取括號內東西即可
                        size = size.replace('X', 'x')
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    if 'x' in size:
                        coor_to_size_set.add(((coor1, coor2), size.strip()))
                    else:
                        coor_to_size_set.add(((coor1, coor2), '另詳'))

                # 找框框，完成block_coor_list，格式為((0.0, 0.0), (14275.54, 10824.61))
                # 此處不會報錯
                if object_layer in block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    block_coor_list.append((coor1, coor2))
                    if _cal_ratio(coor1, coor2) >= 1/5 and _cal_ratio(coor1, coor2) <= 5:  # 避免雜訊影響框框
                        block_coor_list.append((coor1, coor2))

                if object_layer in size_layer and object.EntityName in text_entity_name:
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    table_to_size_set.append(
                        (coor1, object.TextString.strip()))

                if object_layer in table_line_layer and (object.EntityName == "AcDbLine" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    table_line_list.append((coor1, coor2))
                flag = 1

                if object_layer in column_block_layer and object.EntityName in column_block_object_type:
                    # column_block = ColumnBlock(object.GetBoundingBox()[
                    #                            0], object.GetBoundingBox()[1])
                    column_block_entity.append((object.GetBoundingBox()[0],
                                                object.GetBoundingBox()[1]))
            except Exception as ex:
                object_list.append(object)
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 7: {ex}, error_count = {error_count}.')
    progress('平面圖讀取進度 7/11')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    error_count = 0
    while error_count < 3:
        try:
            doc_plan.Close(SaveChanges=False)
            break
        except:
            time.sleep(3)
            error_count += 1
        # return False
    return {
        'coor_to_floor_set': coor_to_floor_set,
        'coor_to_col_set': coor_to_col_set,
        'coor_to_size_set': coor_to_size_set,
        'block_coor_list': block_coor_list,
        'table_to_size_list': table_to_size_set,
        'table_line_list': table_line_list,
        'column_block_entity': column_block_entity
    }


def sort_plan(plan_data: dict,
              drawing_unit: Literal['cm', 'mm'] = 'cm'):
    # Step 8. 透過 coor_to_floor_set 以及 block_coor_list 完成 floor_to_coor_set，格式為(floor, block左下角和右上角的coor)
    # 此處不會報錯，沒在框框裡就直接扔了
    block_coor_list = plan_data['block_coor_list']
    coor_to_floor_set = plan_data['coor_to_floor_set']
    coor_to_col_set = plan_data['coor_to_col_set']
    coor_to_size_set = plan_data['coor_to_size_set']
    table_to_size_list = plan_data['table_to_size_list']
    table_line_list = plan_data['table_line_list']
    column_block_entity = plan_data['column_block_entity']
    floor_to_coor_set = set()
    floor_table_size_dict: dict[str, dict] = dict()

    for string_coor, floor in coor_to_floor_set:  # set (字串的coor, floor)
        for block_coor in block_coor_list:
            x_diff_left = string_coor[0] - block_coor[0][0]  # 和左下角的diff
            y_diff_left = string_coor[1] - block_coor[0][1]
            x_diff_right = string_coor[0] - block_coor[1][0]  # 和右上角的diff
            y_diff_right = string_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:  # 要在框框裡面才算
                floor_to_coor_set.add((floor, block_coor, string_coor))
    for floor, block_coor, floor_string_coor in floor_to_coor_set:
        floor_table_size_dict[floor] = {}

        floor_table_size = [(string_coor, text if drawing_unit == 'cm' else convert_mm_to_cm(text))
                            for string_coor, text in table_to_size_list
                            if string_coor[0] - block_coor[0][0] > 0 and
                            string_coor[1] - block_coor[0][1] > 0 and
                            string_coor[0] - block_coor[1][0] < 0 and
                            string_coor[1] - block_coor[1][1] < 0]
        floor_table_line = [(coor1, coor2) for coor1, coor2 in table_line_list
                            if coor1[0] - block_coor[0][0] > 0 and
                            coor1[1] - block_coor[0][1] > 0 and
                            coor2[0] - block_coor[1][0] < 0 and
                            coor2[1] - block_coor[1][1] < 0 and
                            coor1[1] == coor2[1]]
        # floor_table_size.sort(key=lambda x: x[0][1])
        # match_string_list = [(k, list(g))
        #                      for k, g in groupby(floor_table_size, lambda x: x[0][1])]
        for string_coor, text in floor_table_size:
            upper_bound = [(0, 0), (0, 0)]
            lower_bound = [(0, 0), (0, 0)]
            if 'Cn' in text:
                upper_line = [
                    line for line in floor_table_line if line[0][1] >= string_coor[1]]
                lower_line = [
                    line for line in floor_table_line if line[0][1] <= string_coor[1]]
                if upper_line:
                    upper_bound = min(upper_line, key=lambda l: l[0][1])
                if lower_line:
                    lower_bound = max(lower_line, key=lambda l: l[0][1])
                sizes = [t for s_coor, t in floor_table_size if re.match(
                    r'\d+[x|X]\d+', t) and upper_bound[0][1] >= s_coor[1] >= lower_bound[0][1]]
                if sizes:
                    real_size = re.match(r'\d+[x|X]\d+', sizes[0]).group()
                    floor_table_size_dict[floor].update({'Cn': real_size})
            elif re.match(r'C\d+', text):
                upper_line = [
                    line for line in floor_table_line if line[0][1] <= string_coor[1]]
                lower_line = [
                    line for line in floor_table_line if line[0][1] >= string_coor[1]]
                if upper_line:
                    upper_bound = min(upper_line, key=lambda l: l[0][1])
                if lower_line:
                    lower_bound = max(lower_line, key=lambda l: l[0][1])
                sizes = [t for s_coor, t in floor_table_size if re.match(
                    r'\d+[x|X]\d+', t) and upper_bound[0][1] >= s_coor[1] >= lower_bound[0][1]]
                if sizes:
                    real_size = re.match(r'\d+[x|X]\d+', sizes[0]).group()
                    floor_table_size_dict[floor].update({text: real_size})

    progress('平面圖讀取進度 8/11')

    # Step 9. 算出Bmax, Fmax, Rmax
    # 此處可能報錯的地方在於turn_floor_to_float，但函式本身return false時就會報錯，所以此處不另外再報錯
    Bmax = 0  # 地下最深到幾層(不包括FB不包括FB)
    Fmax = 0  # 正常樓最高到幾層
    Rmax = 0  # R開頭最高到幾層(不包括PRF)
    for y in floor_to_coor_set:
        floor = y[0]
        tmp_floor_list = []
        to_bool = False
        for char in weird_to_list:
            if char in floor:
                to_char = char
                start = floor.split(to_char)[0]
                end = floor.split(to_char)[1]
                if not (turn_floor_to_float(start)) or not turn_floor_to_float(end):
                    for temp in re.split(r'\W+', floor):
                        tmp_floor_list.append(temp)
                else:
                    tmp_floor_list.append(turn_floor_to_float(start))
                    tmp_floor_list.append(turn_floor_to_float(end))
                to_bool = True
                break
        if not to_bool:
            comma_char = ','
            for char in weird_comma_list:
                if char in floor:
                    comma_char = char
                    break
            comma = floor.count(comma_char)
            for i in range(comma + 1):
                tmp_floor_list.append(
                    turn_floor_to_float(floor.split(comma_char)[i]))

        for x in tmp_floor_list:
            if x < 0 and x < Bmax and x != -1000:
                Bmax = x
            elif x > 0 and x < 1000 and x > Fmax:
                Fmax = x
            elif x > 1000 and x != 2000:
                Rmax = x
    progress('平面圖讀取進度 9/11')

    # 先把不合理的樓層踢掉
    new_floor_to_coor_list = []
    for x in floor_to_coor_set:
        floor_name = x[0]
        block_coor = x[1]
        floor_string_coor = x[2]
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_floor_to_coor_list.append(
                (floor_list, block_coor, floor_string_coor))

    floor_to_coor_set = new_floor_to_coor_list

    new_coor_to_floor_list = []
    for x in coor_to_floor_set:
        string_coor = x[0]
        floor_name = x[1]
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_coor_to_floor_list.append((string_coor, floor_list))

    coor_to_floor_set = new_coor_to_floor_list

    # Step 10. 完成col_size_coor_set，格式: set(col, size, the coor of big_block(left, right, up, down))
    col_size_coor_set = set()
    for x in coor_to_col_set:
        col_coor = x[0][0]
        col_full_coor = x[0]
        col_name = x[1]
        min_diff = 1000 if drawing_unit == 'cm' else 10000
        match_size = ''
        match_size_coor = ''

        closest_size = min(coor_to_size_set, key=lambda y: abs(
            col_coor[0] - y[0][0][0]) + abs(col_coor[1] - y[0][0][1]))
        if abs(col_coor[0] - closest_size[0][0][0]) + abs(col_coor[1] - closest_size[0][0][1]) < min_diff:

            match_size = closest_size[1]
            match_size_coor = closest_size[0]

            if drawing_unit == 'mm':
                match_size = convert_mm_to_cm(match_size)

            left = min(col_full_coor[0][0], match_size_coor[0][0])
            right = max(col_full_coor[1][0], match_size_coor[1][0])
            up = max(col_full_coor[1][1], match_size_coor[1][1])
            down = min(col_full_coor[0][1], match_size_coor[0][1])

            col_size_coor_set.add(
                (col_name, match_size, (left, right, up, down)))
        else:
            col_size_coor_set.add(
                (col_name, '', (col_full_coor[0][0], col_full_coor[1][0], col_full_coor[1][1], col_full_coor[0][1])))
    progress('平面圖讀取進度 10/11')
    # # DEBUG: 檢查col跟size有沒有被圈在一起，或者被亂圈到其他地方
    # for x in col_size_coor_set:
    #     coor = x[2]
    #     coor_list = [coor[0] - 20, coor[3] - 20, 0, coor[1] + 20, coor[3] - 20, 0, coor[1] + 20, coor[2] + 20, 0, coor[0] - 20, coor[2] + 20, 0, coor[0] - 20, coor[3] - 20, 0]
    #     points = vtFloat(coor_list)
    #     pointobj = msp_plan.AddPolyline(points)
    #     for i in range(4):
    #         pointobj.SetWidth(i, 10, 10)
    # return

    # Step 11. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補
    set_plan = set()  # set元素為 (樓層, col, size)
    dic_plan = {}  # 透過(樓層, col, size)去找col跟size的整體座標
    # set(col, size, the coor of big_block(left, right, up, down))
    floor_to_col_string_dict = defaultdict(list)
    for x in col_size_coor_set:
        col_coor = (x[2][0], x[2][3])  # 取左下角即可
        full_coor = x[2]  # 左下跟右上都有
        col_name = x[0]
        col_size = x[1]
        min_floor = None

        block_list = [(floor_list, block_coor, string_coor) for floor_list, block_coor, string_coor in floor_to_coor_set if
                      (col_coor[0] - block_coor[0][0]) * (col_coor[0] - block_coor[1][0]) < 0 and
                      (col_coor[1] - block_coor[0][1]) * (col_coor[1] - block_coor[1][1]) < 0]
        if len(block_list) > 1:
            block_list = [
                block for block in block_list if block[2][1] < col_coor[1]]

        if len(block_list) > 1:
            min_floor = min(
                block_list, key=lambda x: get_distance(x[2], col_coor))[0]
        if len(block_list) == 1:
            min_floor = block_list[0][0]

        floor_list = min_floor

        if floor_list is None:
            continue
        for floor in floor_list:
            if col_size == '' and floor in floor_table_size_dict:
                if col_name in floor_table_size_dict[floor]:
                    col_size = floor_table_size_dict[floor][col_name]
                elif 'Cn' in floor_table_size_dict[floor]:
                    col_size = floor_table_size_dict[floor]['Cn']
            floor_to_col_string_dict[floor].append(
                (col_coor, col_name, col_size))
            set_plan.add((floor, col_name, col_size))
            dic_plan[(floor, col_name, col_size)] = full_coor
        else:
            error('read_plan error in step 11: min_floor cannot be found.')

    progress('平面圖讀取進度 11/11')
    # match plan column block and string
    floor_column_block_entity_dict = defaultdict(list)
    for column_block in column_block_entity:
        start_pt = column_block[0]
        end_pt = column_block[1]

        min_floor = None

        block_list = [(floor_list, block_coor, string_coor) for floor_list, block_coor, string_coor in floor_to_coor_set if
                      (start_pt[0] - block_coor[0][0]) * (start_pt[0] - block_coor[1][0]) < 0 and
                      (start_pt[1] - block_coor[0][1]) * (start_pt[1] - block_coor[1][1]) < 0]
        if len(block_list) > 1:
            block_list = [
                block for block in block_list if block[2][1] < start_pt[1]]

        if len(block_list) > 1:
            min_floor = min(
                block_list, key=lambda x: get_distance(x[2], start_pt))[0]
        if len(block_list) == 1:
            min_floor = block_list[0][0]
        if min_floor is None:
            continue
        for floor in min_floor:
            mid_pt = ((start_pt[0] + end_pt[0]) / 2,
                      (start_pt[1] + end_pt[1]) / 2)
            x_block_length = round(abs(start_pt[0] - end_pt[0]), 2)
            y_block_length = round(abs(start_pt[1] - end_pt[1]), 2)

            if drawing_unit == 'mm':
                x_block_length /= 10.0
                y_block_length /= 10.0

            floor_column_block_entity_dict[floor].append(
                (mid_pt, (start_pt, end_pt), (x_block_length, y_block_length)))

    col_string_match_col_block = []
    for floor, col_block_items in floor_column_block_entity_dict.items():
        if floor not in floor_to_col_string_dict:
            continue
        col_items = floor_to_col_string_dict[floor]
        match_result, _ = match_points([col_block[0] for col_block in col_block_items],
                                       [col_string[0] for col_string in col_items])
        for i, j in match_result:
            col_string_match_col_block.append(
                (floor, col_block_items[i], col_items[j]))

    block_error_list = []
    for floor, col_block, col_string in col_string_match_col_block:
        mid_pt, (start_pt, end_pt), (x_block_length,
                                     y_block_length) = col_block
        col_coor, col_name, col_size = col_string

        if len(re.findall(r"\d+", col_size)) < 2:
            continue

        match_group = re.findall(r"\d+", col_size)
        x_length = int(match_group[0])
        y_length = int(match_group[1])

        if int(x_block_length) != x_length or int(y_block_length) != y_length:
            block_error_list.append(
                (floor, col_name, col_size, f'{x_block_length}x{y_block_length}'))

    progress('平面圖讀取完畢。')

    return (set_plan, dic_plan, block_error_list, col_string_match_col_block)

# def read_col(col_filename, text_layer, line_layer, result_filename, progress_file):


def read_col(col_filename: str, layer_config: dict):
    '''
    Read AutoCAD XS-COL info.
    '''
    text_layer = layer_config['text_layer']
    line_layer = layer_config['line_layer']
    error_count = 0
    progress('開始讀取柱配筋圖')
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_col = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_col error in step 1: {ex}, error_count = {error_count}.')
    progress('柱配筋圖讀取進度 1/10')

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_col = wincad_col.Documents.Open(col_filename)
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_col error in step 2: {ex}, error_count = {error_count}.')
    progress('柱配筋圖讀取進度 2/10')

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_col = doc_col.Modelspace
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_col error in step 3: {ex}, error_count = {error_count}.')
    progress('柱配筋圖讀取進度 3/10')

    # Step 4. 解鎖所有圖層 -> 不然不能刪東西
    flag = 0
    while not flag and error_count <= 10:
        try:
            layer_count = doc_col.Layers.count

            for x in range(layer_count):
                layer = doc_col.Layers.Item(x)
                layer.Lock = False
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            msp_col = doc_col.Modelspace
            error(
                f'read_col error in step 4: {e}, error_count = {error_count}.')
    progress('柱配筋圖讀取進度 4/10')

    progress('柱配筋圖讀取進度 5/10')

    # Step 7. 遍歷所有物件 -> 完成一堆座標對應的set跟list
    progress('正在遍歷柱配筋圖上所有物件並篩選出有效信息，運行時間取決於柱配筋圖大小，請耐心等候')
    coor_to_floor_set = set()  # set(coor, floor)
    coor_to_col_set = set()  # set(coor, col)
    coor_to_size_set = set()  # set(coor, size)
    coor_to_floor_line_list = []  # (橫線y座標, start, end)
    coor_to_col_line_list = []  # (縱線x座標, start, end)
    flag = 0
    total = msp_col.Count
    count = 0
    progress(
        f'柱配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
    used_layer_list = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name
    for msp_object in msp_col:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'柱配筋圖已讀取{count}/{total}個物件')
        while error_count <= 3 and not object_list:
            try:
                if msp_object.Layer not in used_layer_list:
                    break
                # print(f'{msp_object.Layer}:{msp_object.EntityName}')
                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference":
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read_col error in step 7-1: {ex}, error_count = {error_count}.')
        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer
                if object_layer in text_layer and object.ObjectName == "AcDbText":
                    if object.TextString[0] == 'C' and len(object.TextString) <= 7:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_col_set.add(
                            ((coor1, coor2), object.TextString))

                    elif 'x' in object.TextString or 'X' in object.TextString:
                        size = object.TextString.replace('X', 'x').strip()
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_size_set.add(((coor1, coor2), size))
                    elif ('F' in object.TextString or 'B' in object.TextString or 'R' in object.TextString) and 'O' not in object.TextString:  # 可能有樓層
                        floor = object.TextString.strip()
                        if '_' in floor:  # 可能有B_6F表示B棟的6F
                            floor = floor.split('_')[1]
                        if turn_floor_to_float(floor):
                            coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                                object.GetBoundingBox()[0][1], 2))
                            coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                                object.GetBoundingBox()[1][1], 2))
                            floor = turn_floor_to_float(floor)
                            floor = turn_floor_to_string(floor)
                            coor_to_floor_set.add(((coor1, coor2), floor))

                elif object_layer in line_layer:
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    if coor1[0] == coor2[0]:
                        coor_to_col_line_list.append(
                            (coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1])))
                    elif coor1[1] == coor2[1]:
                        coor_to_floor_line_list.append(
                            (coor1[1], min(coor1[0], coor2[0]), max(coor1[0], coor2[0])))
            except Exception as ex:
                object_list.append(object)
                error_count += 1
                time.sleep(5)
                error(
                    f'read_col error in step 7: {ex}, error_count = {error_count}.')
    coor_to_col_line_list.sort(key=lambda x: x[0])
    coor_to_floor_line_list.sort(key=lambda x: x[0])
    progress('柱配筋圖讀取進度 7/10')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    error_count = 0
    while error_count < 3:
        try:
            doc_col.Close(SaveChanges=False)
            break
        except:
            time.sleep(3)
            error_count += 1

    return {
        'coor_to_col_set': coor_to_col_set,
        'coor_to_col_line_list': coor_to_col_line_list,
        'coor_to_floor_set': coor_to_floor_set,
        'coor_to_floor_line_list': coor_to_floor_line_list,
        'coor_to_size_set': coor_to_size_set
    }


def sort_col(col_data: dict,
             drawing_unit: Literal['cm', 'mm'] = 'cm',
             bottom_line_offset=1,
             exclude_string: list[str] = []):
    # Step 8. 完成col_to_line_set 格式:(col, left, right, up)
    '''

    '''
    coor_to_col_set = col_data['coor_to_col_set']
    coor_to_col_line_list = col_data['coor_to_col_line_list']
    coor_to_floor_set = col_data['coor_to_floor_set']
    coor_to_floor_line_list = col_data['coor_to_floor_line_list']
    coor_to_size_set = col_data['coor_to_size_set']

    col_to_line_set = set()
    for x in coor_to_col_set:
        coor = x[0]
        col = x[1]
        new_coor_to_col_line_list = []
        for y in coor_to_col_line_list:  # (縱線x座標, start, end)
            if y[1] <= coor[0][1] <= y[2]:  # 先看y座標有沒有被夾住
                new_coor_to_col_line_list.append(y)
        tmp_set = set(new_coor_to_col_line_list)
        new_coor_to_col_line_list = list(tmp_set)
        new_coor_to_col_line_list.sort(key=lambda x: x[0])
        for y in range(len(new_coor_to_col_line_list)-1):  # 再看x座標被哪兩條線夾住
            if new_coor_to_col_line_list[y][0] < coor[0][0] < new_coor_to_col_line_list[y+1][0]:
                col_to_line_set.add(
                    (col, new_coor_to_col_line_list[y][0], new_coor_to_col_line_list[y+1][0], coor[1][1]))
    progress('柱配筋圖讀取進度 8/10')

    # Step 9. 完成floor_to_line_set 格式:(floor, down, up, left)
    floor_to_line_set = set()
    for x in coor_to_floor_set:
        coor = x[0]
        floor = x[1]
        new_coor_to_floor_line_list = []
        for y in coor_to_floor_line_list:  # (橫線y座標, start, end)
            if y[1] <= coor[0][0] <= y[2]:  # 先看x座標有沒有被夾住
                failed = 0  # 查看是否有重複
                for z in new_coor_to_floor_line_list:
                    if z[0] == y[0]:
                        failed = 1
                        break
                if failed == 0:
                    new_coor_to_floor_line_list.append(y)
        tmp_set = set(new_coor_to_floor_line_list)
        new_coor_to_floor_line_list = list(tmp_set)
        new_coor_to_floor_line_list.sort(key=lambda x: x[0])
        try:
            # 再看y座標被哪兩條線夾住，下面那條要往下平移一格
            for y, line in enumerate(new_coor_to_floor_line_list):
                if new_coor_to_floor_line_list[y][0] < coor[0][1] < new_coor_to_floor_line_list[y+1][0]:
                    floor_to_line_set.add(
                        (floor, new_coor_to_floor_line_list[y - bottom_line_offset][0],
                         new_coor_to_floor_line_list[y+1][0], coor[0][0]))
        except:
            pass
    progress('柱配筋圖讀取進度 9/10')

    # Step 10. 完成set_col和dic_col
    dic_col = {}
    set_col = set()
    for x in coor_to_size_set:
        coor = x[0]
        size = x[1]

        if find_all_matching_patterns(text=size, patterns=exclude_string):
            continue

        size = extract_dimensions(size)

        if drawing_unit == 'mm':
            size = convert_mm_to_cm(size)
        min_floor = ''
        min_floor_coor = ''
        min_floor_diff = 10000
        min_col: str = ''
        min_col_coor = ''
        min_col_diff = 10000
        for y in floor_to_line_set:  # (floor, down, up, left)
            if y[1] <= coor[1][1] <= y[2] and coor[1][0] - y[3] >= 0 and coor[1][0] - y[3] <= min_floor_diff:
                min_floor = y[0]
                min_floor_coor = (y[1], y[2])
                min_floor_diff = coor[1][0] - y[3]
        for y in col_to_line_set:  # (col, left, right, up)
            if y[1] <= coor[1][0] <= y[2] and y[3] - coor[1][1] >= 0 and y[3] - coor[1][1] <= min_col_diff:
                min_col = y[0]
                min_col_coor = (y[1], y[2])
                min_col_diff = y[3] - coor[1][1]
        if min_floor != '' and min_col != '':
            if re.findall(special_char_pattern, min_col):
                cols = re.split(special_char_pattern, min_col)
                for col in cols:
                    set_col.add((min_floor, col, size))
                    dic_col[(min_floor, col, size)] = (
                        min_col_coor[0], min_col_coor[1], min_floor_coor[1], min_floor_coor[0])
            elif '-' in min_col:
                try:
                    start = int((min_col.split('-')[0]).split('C')[1])
                    end = int((min_col.split('-')[1]).split('C')[1])
                    for i in range(start, end + 1):
                        set_col.add((min_floor, f'C{i}', size))
                        dic_col[(min_floor, f'C{i}', size)] = (
                            min_col_coor[0], min_col_coor[1], min_floor_coor[1], min_floor_coor[0])  # (left, right, up, down)
                except Exception:  # CW3-1 之類的不是區間
                    set_col.add((min_floor, min_col, size))
                    dic_col[(min_floor, min_col, size)] = (min_col_coor[0], min_col_coor[1],
                                                           min_floor_coor[1], min_floor_coor[0])  # (left, right, up, down)
            else:
                set_col.add((min_floor, min_col, size))
                dic_col[(min_floor, min_col, size)] = (min_col_coor[0], min_col_coor[1],
                                                       min_floor_coor[1], min_floor_coor[0])  # (left, right, up, down)

    # doc_col.Close(SaveChanges=False)
    progress('柱配筋圖讀取進度 10/10')
    progress('柱配筋圖讀取完成。')
    # col.txt單純debug用，不想多新增檔案可以註解掉
    # with open(result_filename, "w", encoding='utf-8') as f:
    #     f.write("in col: \n")
    #     l = list(set_col)
    #     l.sort()
    #     for x in l:
    #         f.write(f'{x}\n')

    return (set_col, dic_col)


# 完成 in plan but not in col 的部分並在圖上mark有問題的部分
def write_plan(plan_filename,
               plan_new_filename,
               set_plan,
               set_col,
               dic_plan,
               date,
               drawing,
               block_match,
               client_id):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    error_count = 0
    progress("開始標註平面圖(核對項目: 柱配筋)及輸出核對結果至'column.txt'。")
    pythoncom.CoInitialize()
    set1 = set_plan - set_col
    list1 = list(set1)
    list1.sort()
    set2 = set_col - set_plan
    list2 = list(set2)
    list2.sort()
    # f = open(result_filename, "w", encoding='utf-8')

    # f.write("in plan but not in col: \n")
    # f.write(f"------------------------------\n")
    if drawing:
        # Step 1. 開啟應用程式
        flag = 0
        while not flag and error_count <= 10:
            try:
                wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_plan error in step 1, {e}, error_count = {error_count}.')
        progress('平面圖標註進度 1/5')
        # Step 2. 匯入檔案
        flag = 0
        while not flag and error_count <= 10:
            try:
                doc_plan = wincad_plan.Documents.Open(plan_filename)
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_plan error in step 2, {e}, error_count = {error_count}.')
        progress('平面圖標註進度 2/5')
        # Step 3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag and error_count <= 10:
            try:
                msp_plan = doc_plan.Modelspace
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_plan error in step 3, {e}, error_count = {error_count}.')
        time.sleep(5)
        progress('平面圖標註進度 3/5')
        # Step 4. 設定mark的圖層
        flag = 0
        while not flag and error_count <= 10:
            try:
                layer_plan = doc_plan.Layers.Add(f"S-CLOUD_{date}")
                doc_plan.ActiveLayer = layer_plan
                layer_plan.color = 10
                layer_plan.Linetype = "Continuous"
                layer_plan.Lineweight = 0.5
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_plan error in step 4, {e}, error_count = {error_count}.')
        progress('平面圖標註進度 4/5')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_plan.Close(SaveChanges=False)
        except:
            pass
        return False

    # Step 5. 完成in plan but not in col，畫圖，以及計算錯誤率
    error_num = 0
    error_list = []
    for x in list1:
        if x[0] != 'FBF':
            wrong_data = 0
            for y in list2:
                if x[0] == y[0] and x[1] == y[1]:
                    error_list.append((x, 0, y[2]))
                    # f.write(f'{x}: 尺寸有誤，在XS-COL那邊是{y[2]}\n')
                    wrong_data = 1
                    break
            if not wrong_data:
                error_list.append((x, 1))
                # f.write(f'{x}: 找不到這根柱子\n')

            error_num += 1

            if drawing:
                coor = dic_plan[x]
                coor_list = [coor[0] - 20, coor[3] - 20, 0, coor[1] + 20, coor[3] - 20, 0, coor[1] +
                             20, coor[2] + 20, 0, coor[0] - 20, coor[2] + 20, 0, coor[0] - 20, coor[3] - 20, 0]
                points = vtFloat(coor_list)
                pointobj = msp_plan.AddPolyline(points)
                for i in range(4):
                    pointobj.SetWidth(i, 10, 10)
    if drawing:
        for block in block_match:
            floor, col_block_item, col_item = block
            block_pt = col_block_item[0]
            col_pt = col_item[0]
            lineObj = msp_plan.AddPolyline(
                vtFloat([block_pt[0], block_pt[1], 0, col_pt[0], col_pt[1], 0]))
            lineObj.setWidth(0, 3, 3)

    error_list = sorted(error_list, key=cmp_to_key(mycmp))
    # for x in error_list:
    #     if x[1] == 0:
    #         f.write(f'{x[0]}: 尺寸有誤，在XS-COL那邊是{x[2]}\n')
    #     else:
    #         f.write(f'{x[0]}: 找不到這根柱子\n')

    if drawing:
        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)

    # count = 0
    # for x in set_plan:
    #     count += 1

    # 計算錯誤率可能會噴錯，因為分母為0
    # try:
    #     rate = round(error_num / count * 100, 2)
    #     f.write(f'error rate = {rate} %\n')
    # except:
    #     rate = 'unfinish'
    #     error(f'write_plan error in step 5, there are no col in plan.txt?')

    # f.close()
    progress('平面圖標註進度 5/5')
    progress("標註平面圖(核對項目: 柱配筋)及輸出核對結果。")
    return error_list


# 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
def write_col(col_filename,
              col_new_filename,
              set_plan,
              set_col,
              dic_col,
              date,
              drawing,
              client_id):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    error_count = 0
    progress("開始標註柱配筋圖及輸出核對結果至'column.txt'。")
    pythoncom.CoInitialize()
    set1 = set_plan - set_col
    list1 = list(set1)
    list1.sort()
    set2 = set_col - set_plan
    list2 = list(set2)
    list2.sort()

    # f = open(result_filename, "a", encoding='utf-8')
    # f.write("in col but not in plan: \n")
    if drawing:
        # Step 1. 開啟應用程式
        flag = 0
        while not flag and error_count <= 10:
            try:
                wincad_col = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_col error in step 1, {ex}, error_count = {error_count}.')
        progress('柱配筋圖標註進度 1/5')
        # Step 2. 匯入檔案
        flag = 0
        while not flag and error_count <= 10:
            try:
                doc_col = wincad_col.Documents.Open(col_filename)
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_col error in step 2, {ex}, error_count = {error_count}.')
        progress('柱配筋圖標註進度 2/5')
        # Step 3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag and error_count <= 10:
            try:
                msp_col = doc_col.Modelspace
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_col error in step 3, {ex}, error_count = {error_count}.')
        time.sleep(5)
        progress('柱配筋圖標註進度 3/5')
        # Step 4. 設定mark的圖層
        flag = 0
        while not flag and error_count <= 10:
            try:
                layer_col = doc_col.Layers.Add(f"S-CLOUD_{date}")
                doc_col.ActiveLayer = layer_col
                layer_col.color = 10
                layer_col.Linetype = "Continuous"
                layer_col.Lineweight = 0.5
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_col error in step 4, {e}, error_count = {error_count}.')
        progress('柱配筋圖標註進度 4/5')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_col.Close(SaveChanges=False)
        except:
            pass
        return False

    # Step 5. 完成in plan but not in col，畫圖，以及計算錯誤率
    error_num = 0
    error_list = []
    for x in list2:
        wrong_data = 0
        for y in list1:
            if x[0] == y[0] and x[1] == y[1]:
                error_list.append((x, 0, y[2]))
                # f.write(f'{x}: 尺寸有誤，在XS-COL那邊是{y[2]}\n')
                wrong_data = 1
                break
        if not wrong_data:
            error_list.append((x, 1))
            # f.write(f'{x}: 找不到這根柱子\n')

        error_num += 1

        if drawing:
            coor = dic_col[x]
            coor_list = [coor[0], coor[3], 0, coor[1], coor[3], 0, coor[1],
                         coor[2], 0, coor[0], coor[2], 0, coor[0], coor[3], 0]
            points = vtFloat(coor_list)
            pointobj = msp_col.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)

    error_list = sorted(error_list, key=cmp_to_key(mycmp))
    # for x in error_list:
    #     if x[1] == 0:
    #         f.write(f'{x[0]}: 尺寸有誤，在XS-PLAN那邊是{x[2]}\n')
    #     else:
    #         f.write(f'{x[0]}: 找不到這根柱子\n')

    if drawing:
        doc_col.SaveAs(col_new_filename)
        doc_col.Close(SaveChanges=True)

    count = 0

    for x in set_col:
        count += 1

    # 計算錯誤率可能會噴錯，因為分母為0
    # try:
    #     rate = round(error_num / count * 100, 2)
    #     f.write(f'error rate = {rate} %\n')
    # except Exception as ex:
    #     rate = 'unfinish'
    #     error(
    #         f'write_col error in step 5, there are no col in col.txt? ex:{ex}')

    # f.close()
    progress('柱配筋圖標註進度 5/5')
    progress("標註柱配筋圖及輸出核對結果至'column.txt'完成。")
    return error_list


def write_result_log(plan_error_list, col_error_list, set_plan, set_col, block_error_list, block_match_list):
    '''
    (floor,name,size)
    '''
    def string_to_ord(s: str):
        temp = 0
        for i, c in enumerate(s[::-1]):
            temp += ord(c) * (10 ** i)
        return temp

    def error_cmp(a, b):
        if isinstance(a[0], tuple):
            a = a[0]
        if isinstance(b[0], tuple):
            b = b[0]
        if turn_floor_to_float(a[0]) > turn_floor_to_float(b[0]):
            return 1
        elif turn_floor_to_float(a[0]) == turn_floor_to_float(b[0]):
            if string_to_ord(a[1]) > string_to_ord(b[1]):
                return 1
            else:
                return -1
        else:
            return -1

    plan_result_dict = {
        'size error': [],
        'not found': [],
        'block error': [],
        'block item': [],
        'item': [],
    }
    col_result_dict = {
        'size error': [],
        'not found': [],
        'block error': [],
        'no size': [],
        'item': [],
    }
    plan_error_list = sorted(plan_error_list, key=cmp_to_key(error_cmp))
    block_error_list = sorted(block_error_list, key=cmp_to_key(error_cmp))
    for error_result in plan_error_list:
        floor, name, size = error_result[0]
        if error_result[1] == 0:
            xs_col_size = error_result[2]
            plan_result_dict['size error'].append(
                (floor, name, size, f'尺寸有誤，在XS-COL那邊是{xs_col_size}'))
        if error_result[1] == 1:
            plan_result_dict['not found'].append(
                (floor, name, size, f'在XS-COL找不到這根柱子'))

    for block_error in block_error_list:
        col_floor, col_name, col_size, block = block_error
        plan_result_dict['block error'].append(
            (col_floor, col_name, col_size, f'圖塊尺寸為{block}'))
    plan_result_dict['summary'] = {}
    plan_result_dict['item'] = sorted(list(set_plan))
    plan_result_dict['block item'] = sorted(block_match_list)
    if len(plan_result_dict["item"]) == 0:
        error_ratio = 100
    else:
        error_ratio = round((len(plan_result_dict["size error"]) + len(plan_result_dict["not found"]) + len(
            plan_result_dict["block error"]))/len(plan_result_dict["item"]), 2)*100
    plan_result_dict['summary'].update({
        '尺寸錯誤': len(plan_result_dict['size error']),
        '缺少配筋': len(plan_result_dict['not found']),
        '圖塊尺寸錯誤': len(plan_result_dict['block error']),
        '總數': len(plan_result_dict['item']),
        '錯誤率': f'{error_ratio} %'
    })

    for error_result in col_error_list:
        floor, name, size = error_result[0]
        if error_result[1] == 0:
            xs_col_size = error_result[2]
            col_result_dict['size error'].append(
                (floor, name, size, f'尺寸有誤，在XS-PLAN那邊是{xs_col_size}'))
        if error_result[1] == 1:
            col_result_dict['not found'].append(
                (floor, name, size, f'在XS-PLAN找不到這根柱子'))
    col_result_dict['item'] = sorted(list(set_col))
    if len(col_result_dict["item"]) == 0:
        error_ratio = 100
    else:
        error_ratio = round((len(col_result_dict["size error"]) + len(
            col_result_dict["not found"])) / len(col_result_dict["item"]), 2)*100
    col_result_dict['summary'] = {}
    col_result_dict['summary'].update({
        '尺寸錯誤': len(col_result_dict['size error']),
        '缺少配筋': len(col_result_dict['not found']),
        '總數': len(col_result_dict['item']),
        '錯誤率': f'{error_ratio} %'
    })
    output_excel_data = {
        'XS-PLAN 統整': [pd.DataFrame(plan_result_dict['summary'], index=[0])],
        'XS-PLAN 結果': [pd.DataFrame(plan_result_dict['size error'], columns=['樓層', '編號', '尺寸', '錯誤']),
                       pd.DataFrame(plan_result_dict['not found'], columns=[
                           '樓層', '編號', '尺寸', '錯誤']),
                       pd.DataFrame(plan_result_dict['block error'], columns=['樓層', '編號', '尺寸', '錯誤'])],
        'XS-COL 統整': [pd.DataFrame(col_result_dict['summary'], index=[0])],
        'XS-COL 結果': [pd.DataFrame(col_result_dict['size error'], columns=['樓層', '編號', '尺寸', '錯誤']),
                      pd.DataFrame(col_result_dict['not found'], columns=['樓層', '編號', '尺寸', '錯誤'])],
        'XS-PLAN 詳細內容': [pd.DataFrame(plan_result_dict['item'], columns=['floor', 'serial', 'size'])],
        'XS-PLAN 圖塊核對詳細內容': [pd.DataFrame(plan_result_dict['block item'], columns=['floor', 'serial', 'size'])],
        'XS-COL 詳細內容': [pd.DataFrame(col_result_dict['item'], columns=['floor', 'serial', 'size'])]
    }

    return plan_result_dict, col_result_dict, output_excel_data


def run_plan(plan_filename: str,
             layer_config: dict,
             client_id: str,
             drawing_unit: str = 'cm',
             pkl: str = ""):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    if pkl == "":
        plan_col_set = read_plan(plan_filename=plan_filename,
                                 layer_config=layer_config)
        save_temp_file.save_pkl(
            data=plan_col_set, tmp_file=f'{os.path.splitext(plan_filename)[0]}_plan_to_col.pkl')
    else:
        plan_col_set = save_temp_file.read_temp(
            tmp_file=pkl)
    set_plan, dic_plan, \
        block_error_list, column_block_match_list = sort_plan(plan_data=plan_col_set,
                                                              drawing_unit=drawing_unit)
    # output_txt = f'{os.path.splitext(plan_new_filename)[0]}_result.txt'
    return (set_plan, dic_plan, block_error_list, column_block_match_list)


def run_col(col_filename: str,
            layer_config: dict,
            client_id: str,
            drawing_unit='cm',
            bottom_line_offset=1,
            exclude_string=[],
            pkl: str = ""):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    if pkl == "":
        floor_col_set = read_col(col_filename=col_filename,
                                 layer_config=layer_config)
        save_temp_file.save_pkl(
            data=floor_col_set, tmp_file=f'{os.path.splitext(col_filename)[0]}_col_set.pkl')
    else:
        floor_col_set = save_temp_file.read_temp(tmp_file=pkl)
    set_col, dic_col = sort_col(col_data=floor_col_set,
                                drawing_unit=drawing_unit,
                                bottom_line_offset=bottom_line_offset,
                                exclude_string=exclude_string)
    return (set_col, dic_col)


error_file = './result/error_log.txt'  # error_log.txt的路徑

if __name__ == '__main__':
    plan_dwg_file = r'D:\Desktop\BeamQC\TEST\2024-0528\2024-05-28-11-50_temp-XS-PLAN.dwg'
    col_dwg_file = r'D:\Desktop\BeamQC\TEST\2024-0528\2024-05-28-11-50_temp-A.dwg'
    plan_new_filename = r'D:\Desktop\BeamQC\TEST\2024-0528\2024-05-28-11-50_temp-XS-PLAN_new.dwg'
    col_new_filename = r'D:\Desktop\BeamQC\TEST\2024-0528\2024-05-28-11-50_temp-A_new.dwg'
    layer_config = {
        'text_layer': ['S-TEXT'],
        'line_layer': ['S-TABLE'],
        'block_layer': ['0', 'DwFm', 'DEFPOINTS'],
        'floor_layer': ['S-TITLE'],
        'col_layer': ['S-TEXTC'],
        'size_layer': ['S-TEXT'],
        'table_line_layer': ['S-TABLE'],
        'column_block_layer': ['S-COL']
    }
    set_plan, dic_plan, block_error_list, block_match_result_list = run_plan(
        plan_filename=r'D:\Desktop\BeamQC\TEST\2024-0930\XS-PLAN.dwg',
        layer_config=layer_config,
        client_id='0930-temp_col',
        pkl=r"TEST\2024-0930\XS-PLAN_plan_to_col.pkl",
        drawing_unit='mm'
    )
    set_col, dic_col = run_col(
        col_filename=r'D:\Desktop\BeamQC\TEST\2024-0930\XS-COL.dwg',
        layer_config=layer_config,
        client_id='0930-temp-col',
        pkl=r'',
        bottom_line_offset=2,
        exclude_string=['mm']
    )

    date = time.strftime("%Y-%m-%d", time.localtime())
    plan_result = write_plan(plan_dwg_file,
                             plan_new_filename,
                             set_plan,
                             set_col,
                             dic_plan,
                             date,
                             drawing=False,
                             block_match=block_match_result_list,
                             client_id='temp')
    col_result = write_col(col_dwg_file,
                           col_new_filename,
                           set_plan,
                           set_col,
                           dic_col,
                           date,
                           drawing=False,
                           client_id='temp')
    plan_result_dict, col_result_dict, excel_data = write_result_log(plan_error_list=plan_result,
                                                                     col_error_list=col_result,
                                                                     set_plan=set_plan,
                                                                     set_col=set_col,
                                                                     block_error_list=block_error_list,
                                                                     block_match_list=block_match_result_list
                                                                     )
    from main import OutputExcel
    output_folder = r'TEST\2024-0930'
    for sheet_name, df_list in excel_data.items():
        OutputExcel(df_list=df_list,
                    df_spacing=1,
                    file_path=os.path.join(output_folder, 'test.xlsx'),
                    sheet_name=sheet_name)
