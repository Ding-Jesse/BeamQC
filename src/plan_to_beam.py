import win32com.client
import pythoncom
import re
import time
import multiprocessing
import os
import pandas as pd
from functools import cmp_to_key
import src.save_temp_file as save_temp_file
import json
import sys
from gzip import READ
from io import TextIOWrapper
from math import inf, sqrt
from multiprocessing.spawn import prepare
from tabnanny import check
from tkinter import HIDDEN
from numpy import object_
from openpyxl import load_workbook
from collections import Counter
from src.logger import setup_custom_logger
from utils.algorithm import match_points, for_loop_min_match
from typing import Literal
from utils.algorithm import convert_mm_to_cm

weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡B']
beam_head1 = ['B', 'b', 'G', 'g']
beam_head2 = ['FB', 'FG', 'FW', 'Fb', 'Fg', 'CB', 'CG', 'cb']
global main_logger


def turn_floor_to_float(floor):
    '''把字串變成小數 (因為1MF = 1.5, 所以不能用整數)'''

    if ' ' in floor:  # 不小心有空格要把空格拔掉
        floor = floor.replace(' ', '')

    if floor == 'FB':  # FB 直接變-1000層
        floor = str(-1000)

    if floor == 'PRF' or floor == 'PR' or floor == 'PF':  # PRF 直接變2000層
        floor = str(2000)

    if floor == 'RF':  # RF = R1F
        floor = str(1001)

    if 'F' in floor:  # 有F要把F拔掉
        floor = floor.replace("F", "")

    try:  # 轉int可能會失敗
        if 'B' in floor:  # 有B直接變負整數
            floor = str(-int(floor.replace("B", "")))

        if 'R' in floor:  # 有R直接+1000
            floor = str(int(floor.replace("R", "")) + 1000)

        if 'M' in floor:  # 半層以0.5表示
            floor = str(int(floor.replace("M", "")) + 0.5)
    except Exception:
        pass
        # print(f'{floor}: cannot trans to int')

    try:
        floor = float(floor)
        return floor
    except Exception:
        # error(f'turn_floor_to_float error: {floor} cannot be turned to float.')
        return False


def turn_floor_to_string(floor):
    '''
    把數字變回字串
    '''
    if floor == -1000:
        floor = 'FBF'  # 因為beam的部分字尾非F會自動補F，所以在diff的時候要一致

    elif floor > -1000 and floor < 0:
        floor = f'B{int(-floor)}F'

    elif floor > 0 and floor < 1000:
        if floor * 2 % 2 == 0:  # 整數*2之後會是2的倍數
            floor = f'{int(floor)}F'
        else:  # 如果有.5的話，*2之後會是奇數
            floor = f'{int(floor - 0.5)}MF'

    elif floor > 1000 and floor < 2000:
        floor = f'R{int(floor - 1000)}F'

    elif floor == 2000:
        floor = 'PRF'

    else:
        error(
            f'turn_floor_to_string error: {floor} cannot be turned to string.')
        return False

    return floor


def turn_floor_to_list(floor: str, Bmax, Fmax, Rmax):
    '''
    將多重樓層如(2F-RF)轉為[2F....RF]
    '''
    floor_list = []
    to_bool = False
    for char in weird_to_list:
        if char in floor:
            to_char = char
            start = floor.split(to_char)[0]
            end = floor.split(to_char)[1]
            if not turn_floor_to_float(start) or not turn_floor_to_float(end):
                for temp in re.split(r'\W+', floor):
                    floor_list.append(temp)
            else:
                try:
                    start = int(turn_floor_to_float(start))
                    end = int(turn_floor_to_float(end))
                    if start > end:
                        tmp = start
                        start = end
                        end = tmp
                    for i in range(start, end + 1):
                        if floor_exist(i, Bmax, Fmax, Rmax):
                            floor_list.append(turn_floor_to_string(i))
                except Exception:
                    error(
                        f'turn_floor_to_list error: {floor} cannot be turned to list.')
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
            new_floor = floor.split(comma_char)[i]
            new_floor = turn_floor_to_float(new_floor)
            new_floor = turn_floor_to_string(new_floor)
            if new_floor:
                floor_list.append(new_floor)
            else:
                error(
                    f'turn_floor_to_list error: {floor} cannot be turned to list.')

    return floor_list


def floor_exist(i, Bmax, Fmax, Rmax):
    '''
    判斷是否為空號，例如B2F-PRF會從-2跑到2000，但顯然區間裡面的值不可能都合法
    '''
    if i == -1000 or i == 2000:
        return True

    elif i >= Bmax and i < 0:
        return True

    elif i > 0 and i <= Fmax:
        return True

    elif i > 1000 and i <= Rmax:
        return True

    return False


def vtFloat(l):
    '''
    要把點座標組成的list轉成autocad看得懂的樣子
    '''
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)


def error(error_message):
    '''
    把錯誤訊息印到error.log裡面
    '''
    try:
        global main_logger
        main_logger.error(error_message)
    except:
        print(error_message)
    # f = open(error_file, 'a', encoding='utf-8')
    # localtime = time.asctime(time.localtime(time.time()))
    # f.write(f'{localtime} | {error_message}\n')
    # f.close()
    # return


def progress(message):
    '''把進度印到progress裡面，在app.py會對這個檔案做事'''
    global main_logger
    try:
        main_logger.info(message)
    except NameError:
        print(message)
    # f = open(progress_file, 'a', encoding='utf-8')
    # f.write(f'{message}\n')
    # f.close()
    # return

# 可以先看完 write_plan 跟 write_beam 整理 txt 的規則再看這個函式在幹嘛
# 自定義排序規則


def mycmp(a, b):  # a, b 皆為 tuple , 可能是 ((floor, beam), 0, correct) 或 ((floor, beam), 1)
    if a[1] == b[1]:  # err_message 一樣，比樓層
        if turn_floor_to_float(a[0][0]) > turn_floor_to_float(b[0][0]):
            return 1
        elif turn_floor_to_float(a[0][0]) == turn_floor_to_float(b[0][0]):
            if a[0][1] >= b[0][1]:
                return 1
            else:
                return -1
        else:
            return -1
    else:
        if a[1] == 0:
            return 1
        else:
            return -1


def activate_cad(filname):
    error_count = 0
    # Step 13-1. 開啟應用程式
    wincad_instance = None
    while wincad_instance is None and error_count <= 10:
        try:
            wincad_instance = win32com.client.Dispatch("AutoCAD.Application")
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'open cad error in step 13-1, {ex}, error_count = {error_count}')

    # Step 13-2. 匯入檔案
    document = None
    while wincad_instance and document is None and error_count <= 10:
        try:
            document = wincad_instance.Documents.Open(filname)
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read {filname} error in step 13-2, {ex}, error_count = {error_count}')

    # Step 13-3. 載入modelspace(還要畫圖)
    model_space = None
    while document and model_space is None and error_count <= 10:
        try:
            model_space = document.Modelspace
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read {filname} modelspace error in step 13-3, {ex}, error_count = {error_count}')
    return document, model_space


def read_plan(plan_filename, layer_config: dict, sizing, mline_scaling):
    def _cal_ratio(pt1, pt2):
        if abs(pt1[1]-pt2[1]) == 0:
            return 1000
        return abs(pt1[0]-pt2[0])/abs(pt1[1]-pt2[1])
    floor_layer = layer_config['floor_layer']
    beam_layer = layer_config['big_beam_layer'] + \
        layer_config['sml_beam_layer']
    beam_text_layer = layer_config['big_beam_text_layer'] + \
        layer_config['sml_beam_text_layer']
    block_layer = layer_config['block_layer']
    size_layer = layer_config['size_layer']
    text_object_type = ['AcDbAttribute', "AcDbText", "AcDbMLeader"]

    error_count = 0

    doc_plan, msp_plan = activate_cad(plan_filename)
    # Step 4 解鎖所有圖層 -> 不然不能刪東西
    while doc_plan and error_count <= 10:
        try:
            layer_count = doc_plan.Layers.count
            for x in range(layer_count):
                layer = doc_plan.Layers.Item(x)
                layer.Lock = False
            break
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 4: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 4/13')

    # Step 7. 遍歷所有物件 -> 完成各種我們要的set跟list

    progress('正在遍歷平面圖上的物件並篩選出有效信息，運行時間取決於平面圖大小，請耐心等候')

    coor_to_floor_set = set()  # set (字串的coor, floor)，Ex. 求'1F'這個字串的座標在哪
    # set (coor, [beam, size])，Ex. 求'B1-6'這個字串的座標在哪，如果後面有括號的話，順便紀錄尺寸，否則size = ''
    coor_to_beam_set = set()
    block_coor_list = []  # 存取方框最左下角的點座標
    none_concat_size_text_list = []
    # for sizing
    coor_to_size_beam = set()  # set (coor, size_beam)，Ex. 紀錄表格中'Bn'這個字串的座標
    coor_to_size_string = set()  # set (coor, size_string)，Ex. 紀錄表格中'25x50'這個字串的座標

    # for mline_scaling
    # set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale)
    beam_direction_mid_scale_set = set()

    # while not flag and error_count <= 10:
    #     try:
    count = 0
    total = msp_plan.Count
    used_layer_list = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name
    progress(
        f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')

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
                    break
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read_plan error in step 7-1: {ex}, error_count = {error_count}.')

        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer
                # 找size
                if sizing or mline_scaling:
                    if object_layer in size_layer and \
                            object.EntityName in text_object_type and \
                        object.TextString != '' and \
                            object.GetBoundingBox()[0][1] >= 0:
                        coor = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        if 'FGn' in object.TextString:
                            coor_to_size_beam.add((coor, 'FG'))
                        if 'FBn' in object.TextString:
                            coor_to_size_beam.add((coor, 'FB'))
                            continue
                        if 'FWB' in object.TextString:
                            coor_to_size_beam.add((coor, 'FWB'))
                        if 'Fbn' in object.TextString:
                            coor_to_size_beam.add((coor, 'Fb'))
                            continue
                        if 'DB' in object.TextString:
                            coor_to_size_beam.add((coor, 'DB'))  # 車道梁特別處理
                            continue
                        if 'dbn' in object.TextString:
                            coor_to_size_beam.add((coor, 'db'))  # 車道梁特別處理
                            continue
                        if 'Gn' in object.TextString and 'FGn' not in object.TextString:
                            if 'C' in object.TextString and ('(' in object.TextString or object.TextString.count('Gn') >= 2):
                                coor_to_size_beam.add((coor, 'CG'))
                                coor_to_size_beam.add((coor, 'G'))
                            elif 'C' in object.TextString and '(' not in object.TextString:
                                coor_to_size_beam.add((coor, 'CG'))
                            else:
                                coor_to_size_beam.add((coor, 'G'))
                        if 'Bn' in object.TextString and 'W' not in object.TextString and \
                                'D' not in object.TextString and 'FBn' not in object.TextString:
                            if 'C' in object.TextString and ('(' in object.TextString or object.TextString.count('Bn') >= 2):
                                coor_to_size_beam.add((coor, 'CB'))
                                coor_to_size_beam.add((coor, 'B'))
                            elif 'C' in object.TextString and '(' not in object.TextString:
                                coor_to_size_beam.add((coor, 'CB'))
                            else:
                                coor_to_size_beam.add((coor, 'B'))
                        if 'bn' in object.TextString:
                            if 'c' in object.TextString and ('(' in object.TextString or object.TextString.count('bn') >= 2):
                                coor_to_size_beam.add((coor, 'cb'))
                                coor_to_size_beam.add((coor, 'b'))
                            elif 'c' in object.TextString and '(' not in object.TextString:
                                coor_to_size_beam.add((coor, 'cb'))
                            else:
                                coor_to_size_beam.add((coor, 'b'))
                        if 'g' in object.TextString:
                            comma_char = ','
                            for char in weird_comma_list:
                                if char in object.TextString:
                                    comma_char = char
                            comma = object.TextString.count(comma_char)
                            for i in range(comma + 1):
                                beam = object.TextString.split(comma_char)[i]
                                if 'g' in beam and beam.split('g')[1].isdigit():
                                    if 'c' in beam and '(' in beam:
                                        coor_to_size_beam.add(
                                            (coor, beam.split(')')[1]))
                                        coor_to_size_beam.add(
                                            (coor, f"c{beam.split(')')[1]}"))
                                    else:
                                        coor_to_size_beam.add((coor, beam))
                                elif 'g' in beam and beam.split('g')[1] != '' and beam.split('g')[1][0] == 'n':
                                    if 'c' in object.TextString and '(' in object.TextString:
                                        coor_to_size_beam.add((coor, 'cg'))
                                        coor_to_size_beam.add((coor, 'g'))
                                    elif 'c' in object.TextString and '(' not in object.TextString:
                                        coor_to_size_beam.add((coor, 'cg'))
                                    else:
                                        coor_to_size_beam.add((coor, 'g'))
                        if 'x' in object.TextString or \
                            'X' in object.TextString or \
                                'x' in object.TextString:
                            string = (object.TextString.replace(
                                ' ', '')).replace('X', 'x').strip()
                            try:
                                first = string.split('x')[0]
                                second = string.split('x')[1]
                                if float(first) and float(second):
                                    coor_to_size_string.add((coor, string))
                            except Exception as ex:
                                pass

                # 找複線
                if mline_scaling:
                    if object_layer in beam_layer and object.ObjectName == "AcDbMline":
                        start = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        end = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        x_diff = abs(start[0] - end[0])
                        y_diff = abs(start[1] - end[1])
                        mid = ((start[0] + end[0]) / 2,
                               (start[1] + end[1]) / 2)
                        if x_diff + y_diff > 100:  # 防超短的東東
                            if x_diff < y_diff:  # 算直的, 1
                                beam_direction_mid_scale_set.add(
                                    (object_layer, 1, mid, round(abs(object.MLineScale), 2)))
                            else:  # 算橫的, 0
                                beam_direction_mid_scale_set.add(
                                    (object_layer, 0, mid, round(abs(object.MLineScale), 2)))
                # 取floor的字串 -> 抓括號內的字串 (Ex. '十層至十四層結構平面圖(10F~14F)' -> '10F~14F')
                # 若此處報錯，可能原因: 1. 沒有括號, 2. 有其他括號在鬧(ex. )
                if object_layer in floor_layer and \
                    object.EntityName in text_object_type and \
                        '(' in object.TextString:
                    floor = object.TextString.strip()
                    floor = re.search(r'\(([^)]+)', floor).group(1)  # 取括號內的樓層數
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))  # 不取概數的話後面抓座標會出問題，例如兩個樓層在同一格
                    no_chinese = False
                    for ch in floor:  # 待修正
                        if ch == 'B' or ch == 'F' or ch == 'R' or ch.isdigit():
                            no_chinese = True
                            continue
                    if floor != '' and no_chinese:
                        coor_to_floor_set.add((coor, floor))
                    else:
                        error(
                            'read_plan error in step 7: floor is an empty string or it is Chinese. ')
                    continue
                # 取beam的字串
                # 此處會錯的地方在於可能會有沒遇過的怪怪comma，但報應不會在這裡產生，會直接反映到結果
                # use search will cause b7-1(50x70) become a none concat size
                if object_layer in beam_text_layer and object.EntityName in text_object_type:
                    matches = re.match(r'\(\d+(x|X)\d+\)', object.TextString)
                    if matches:
                        beam = matches.group()
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))

                        size = re.search(r'(\d+(x|X)\d+)', beam)
                        if size:
                            none_concat_size_text_list.append(
                                ((coor1, coor2), size.group(0).strip()))
                            continue
                        else:
                            progress(object.TextString)

                if object_layer in beam_text_layer and object.EntityName in text_object_type\
                        and object.TextString != '' and (object.TextString[0] in beam_head1 or object.TextString[0:2] in beam_head2):

                    beam: str = object.TextString
                    beam = beam.replace(' ', '')  # 有空格要把空格拔掉

                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    size = ''
                    matches = re.match(
                        r'(.+(\(|\uff08).+(\)|\uff09))', object.TextString)
                    if matches:
                        beam_size = re.search(r'((\d+)[x|X](\d+))', beam)
                        beam_name = re.search(r'(.+\b(?=\(|\uff08))', beam)
                        if beam_size:
                            size = beam_size.group(0).replace('X', 'x').strip()
                        if beam_name:
                            beam = beam_name.group(0).strip()

                    comma_char = ','
                    for char in weird_comma_list:
                        if char in beam:
                            comma_char = char
                    comma = beam.count(comma_char)
                    for i in range(comma + 1):
                        try:
                            coor_to_beam_set.add(
                                ((coor1, coor2), (beam.split(comma_char)[i].strip(), size, round(object.Rotation, 2))))
                        except Exception as ex:  # 只要不是0or1.57，後面核對的時候就會橫的值得都找。
                            coor_to_beam_set.add(
                                ((coor1, coor2), (beam.split(comma_char)[i].strip(), size, 1)))
                            error(
                                f'read_plan error in step 7: {(beam, size)} at {(coor1, coor2)} cannot find Rotation.')
                    continue
                # 為了排版好看的怪產物，目前看到的格式為'{\W0.7;B4-2\P(80x100)}'，所以使用分號及反斜線來切
                # 切爛了也不會報錯，直接反映在結果
                if object_layer in beam_text_layer and object.ObjectName == "AcDbMText":
                    beam = object.TextString.strip()
                    semicolon = beam.count(';')
                    size = ''
                    for i in range(semicolon + 1):
                        s = beam.split(';')[i]
                        if s[0] in beam_head1 or s[0:2] in beam_head2:
                            if '(' in s:
                                size = (((s.split('(')[1]).split(')')[0]).replace(
                                    ' ', '')).replace('X', 'x')
                                if 'x' not in size:
                                    size = ''
                                else:
                                    try:
                                        first = size.split('x')[0]
                                        second = size.split('x')[1]
                                        if not (float(first) and float(second)):
                                            size = ''
                                    except Exception as ex:
                                        size = ''
                                s = s.split('(')[0]
                            if '\\' in s:
                                s = s.split('\\')[0]
                            beam = s
                            continue

                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))

                    if '(' in beam:
                        beam = beam.split('(')[0]  # 取括號前內容即可
                    if ' ' in beam:
                        beam = beam.replace(' ', '')  # 有空格要把空格拔掉
                    if beam[0] in beam_head1 or beam[0:2] in beam_head2:
                        try:
                            coor_to_beam_set.add(
                                ((coor1, coor2), (beam, size, round(object.Rotation, 2))))
                        except Exception as ex:
                            error(
                                f'read_plan error in step 7: {(beam, size)} at {(coor1, coor2)} cannot find Rotation.')
                    continue
                # 找框框，完成block_coor_list，格式為((0.0, 0.0), (14275.54, 10824.61))
                # 此處不會報錯

                if object_layer in block_layer and \
                        (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    if _cal_ratio(coor1, coor2) >= 1/5 and _cal_ratio(coor1, coor2) <= 5:  # 避免雜訊影響框框
                        block_coor_list.append((coor1, coor2))
                    continue

                continue
            except Exception as ex:
                error_count += 1
                object_list.append(object)
                time.sleep(5)
                error(
                    f'read_plan error in step 7: {ex}, error_count = {error_count}.')
        # except Exception as e:
        #     error_class = e.__class__.__name__ #取得錯誤類型
        #     detail = e.args[0] #取得詳細內容
        #     cl, exc, tb = sys.exc_info() #取得Call Stack
        #     lastCallStack = traceback.extract_tb(tb)[-1] #取得Call Stack的最後一筆資料
        #     fileName = lastCallStack[0] #取得發生的檔案名稱
        #     lineNum = lastCallStack[1] #取得發生的行號
        #     funcName = lastCallStack[2] #取得發生的函數名稱
        #     errMsg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, error_class, detail)
        #     error_count += 1
        #     time.sleep(5)
        #     error(f'read_plan error in step 7: {e}, error_count = {error_count}.')
    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    # if error_count > 10:
    try:
        doc_plan.Close(SaveChanges=False)
    except Exception as ex:
        error('Fail To Close Plan File')
    # return False
    progress('平面圖讀取進度 7/13')
    # doc_plan.Close(SaveChanges=False)
    return {'coor_to_floor_set': coor_to_floor_set,
            'coor_to_beam_set': coor_to_beam_set,
            'block_coor_list': block_coor_list,
            'none_concat_size_text_list': none_concat_size_text_list,
            'coor_to_size_beam': coor_to_size_beam,
            'coor_to_size_string': coor_to_size_string,
            'beam_direction_mid_scale_set': beam_direction_mid_scale_set}


def sort_plan(layer_config: dict,
              plan_data: dict,
              sizing: bool,
              mline_scaling: bool,
              drawing_unit: Literal['cm', 'mm'] = "cm"):
    '''
    dic = {
        'mline':{
            '大梁':[],
            '小梁':[],
            '地梁':[],
        },
        'size':{
            '大梁':[],
            '小梁':[],
            '地梁':[],
        },
        'no_size':{
            '大梁':[],
            '小梁':[],
            '地梁':[],
        },
        'not_found':{
            '大梁':[],
            '小梁':[],
            '地梁':[],
        },
        'duplicate':{
            '大梁':[],
            '小梁':[],
            '地梁':[],
        }
        'plan':[],
        'warning':[]
    }
    '''
    error_count = 0
    warning_list = []
    error_list = []

    def get_distance(coor1, coor2):
        if isinstance(coor1[0], tuple) and isinstance(coor2[0], tuple):
            return abs(coor1[0][0]-coor2[0][0]) + abs(coor1[0][1]-coor2[0][1])
        if isinstance(coor1, tuple) and isinstance(coor2, tuple):
            return abs(coor1[0]-coor2[0]) + abs(coor1[1]-coor2[1])
        return 10000
    # 2023-0308
    # set (字串的coor, floor)，Ex. 求'1F'這個字串的座標在哪
    coor_to_floor_set = plan_data['coor_to_floor_set']
    # set (coor, [beam, size])，Ex. 求'B1-6'這個字串的座標在哪，如果後面有括號的話，順便紀錄尺寸，否則size = ''
    coor_to_beam_set: set[tuple[tuple, list[str]]
                          ] = plan_data['coor_to_beam_set']
    block_coor_list = plan_data['block_coor_list']  # 存取方框最左下角的點座標
    none_concat_size_text_list = plan_data['none_concat_size_text_list']
    # for sizing
    # set (coor, size_beam)，Ex. 紀錄表格中'Bn'這個字串的座標
    coor_to_size_beam: set = plan_data['coor_to_size_beam']
    # set (coor, size_string)，Ex. 紀錄表格中'25x50'這個字串的座標
    coor_to_size_string: set = plan_data['coor_to_size_string']

    big_beam_layer = layer_config['big_beam_layer']
    sml_beam_layer = layer_config['sml_beam_layer']

    # for mline_scaling
    # set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale)
    beam_direction_mid_scale_set = plan_data['beam_direction_mid_scale_set']
    # 2023-0119
    for none_concat_size in none_concat_size_text_list:
        coor, size = none_concat_size
        if drawing_unit == 'mm':
            size = convert_mm_to_cm(size)

        temp_list = [s for s in coor_to_beam_set if s[1][1] == '']
        if not temp_list:
            continue

        closet_beam = min(temp_list, key=lambda x: get_distance(x[0], coor))
        coor_to_beam_set.remove(closet_beam)
        coor_to_beam_set.add(
            (closet_beam[0], (closet_beam[1][0].strip(), size, closet_beam[1][2])))

        # closet_beam[1][1] = size

    # Step 8. 完成size_coor_set (size_beam, size_string, size_coor), Ex. 把表格中的 'Bn' 跟 '50x70' 連起來

    if sizing or mline_scaling:
        # 2024-0520 method update
        size_coor_set = set()
        match_results, _ = for_loop_min_match(points1=[x[0] for x in coor_to_size_beam],
                                              points2=[y[0] for y in coor_to_size_string])
        coor_to_size_beam_list = list(coor_to_size_beam)
        coor_to_size_string_list = list(coor_to_size_string)
        for i, j in match_results:
            coor, beam_size = coor_to_size_beam_list[i]
            size_coor, size_string = coor_to_size_string_list[j]
            if drawing_unit == 'mm':
                size_string = convert_mm_to_cm(size_string)
            size_coor_set.add((beam_size, size_string, coor, size_coor))

        # size_coor_set = set()
        # for x in coor_to_size_beam:
        #     coor = x[0]
        #     size_beam = x[1]
        #     min_size = ''
        #     min_dist = inf
        #     for y in coor_to_size_string:
        #         coor2 = y[0]
        #         size_string = y[1]
        #         dist = abs(coor[0]-coor2[0]) + abs(coor[1] - coor2[1])
        #         if dist < min_dist:
        #             min_size = size_string
        #             min_dist = dist
        #     if min_size != '':
        #         size_coor_set.add((size_beam, min_size, coor))
    progress('平面圖讀取進度 8/13')

    # Step 9. 透過 coor_to_floor_set 以及 block_coor_list 完成 floor_to_coor_set，格式為(floor, block左下角和右上角的coor), Ex. '1F' 左下角和右上角的座標分別為(0.0, 0.0) (14275.54, 10824.61)
    # 此處不會報錯，沒在框框裡就直接扔了
    floor_to_coor_set = set()
    for x in coor_to_floor_set:  # set (字串的coor, floor)
        string_coor: tuple = x[0]
        floor: str = x[1]
        for block_coor in block_coor_list:
            x_diff_left = string_coor[0] - block_coor[0][0]  # 和左下角的diff
            y_diff_left = string_coor[1] - block_coor[0][1]
            x_diff_right = string_coor[0] - block_coor[1][0]  # 和右上角的diff
            y_diff_right = string_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:  # 要在框框裡面才算
                floor_to_coor_set.add((floor, block_coor, string_coor))
    progress('平面圖讀取進度 9/13')

    # Step 10. 算出Bmax, Fmax, Rmax, 用途: 跑for迴圈的時候，知道哪些是空號
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
                        if turn_floor_to_float(temp):
                            tmp_floor_list.append(turn_floor_to_float(temp))
                        else:
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

    # 先把不合理的樓層踢掉
    new_floor_to_coor_list = []
    for x in floor_to_coor_set:
        floor_name = x[0]
        block_coor = x[1]
        string_coor = x[2]
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_floor_to_coor_list.append(
                (floor_list, block_coor, string_coor))

    floor_to_coor_set = new_floor_to_coor_list

    new_coor_to_floor_list = []
    for x in coor_to_floor_set:
        string_coor = x[0]
        floor_name = x[1]
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_coor_to_floor_list.append((string_coor, floor_list))

    coor_to_floor_set = new_coor_to_floor_list

    progress('平面圖讀取進度 10/13')

    # Step 11. 完成floor_beam_size_coor_set (floor, beam, size, coor), 找表格內的物件在哪一個框框裡面，進而找到所屬樓層

    if sizing or mline_scaling:
        floor_beam_size_coor_set = set()
        for x in size_coor_set:  # set(size_beam, min_size, coor)
            size_coor = x[2]
            size_string = x[1]
            size_beam = x[0]
            min_floor = []
            # list (floor_list, block左下角和右上角的coor ,string coor)
            for z in floor_to_coor_set:
                floor_list = z[0]
                block_coor = z[1]
                x_diff_left = size_coor[0] - block_coor[0][0]  # 和左下角的diff
                y_diff_left = size_coor[1] - block_coor[0][1]
                x_diff_right = size_coor[0] - block_coor[1][0]  # 和右上角的diff
                y_diff_right = size_coor[1] - block_coor[1][1]
                if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:
                    if len(min_floor) == 0 or min_floor[0] != floor_list:
                        min_floor.append(floor_list)

            if len(min_floor) != 0:
                for i, _ in enumerate(min_floor):
                    floor_list = min_floor[i]
                    for floor in floor_list:
                        floor_beam_size_coor_set.add(
                            (floor, size_beam, size_string, size_coor))
            else:
                error(
                    f'read_plan error in step 11: {(size_beam, size_string, size_coor)} cannot find min_floor.')
    progress('平面圖讀取進度 11/13')

    # Step 12. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補

    set_plan = set()  # set元素為 (樓層, 梁柱名稱, size)
    dic_plan = {}  # 透過(floor, beam, size)去找字串座標

    # 如果沒有要對size -> set元素為 (floor, beam)
    # 如果有要對size但沒有要對mline -> set元素為 (floor, beam, size)
    # 如果有要對size和mline -> set元素為 (floor, beam, size, rotate)
    check_list = []
    # 遍歷所有beam，找這是幾樓的
    for x in coor_to_beam_set:  # set(coor, (beam, size))
        beam_coor = x[0][0]  # 取左下角即可
        full_coor = x[0]  # 左下跟右上都有
        beam_name = x[1][0]
        beam_size = x[1][1]
        beam_rotate = x[1][2]
        min_floor = None

        if drawing_unit == 'mm':
            beam_size = convert_mm_to_cm(beam_size)

        # 我其實是list歐哈哈 (floor_list, block左下角和右上角的coor)
        block_list = [(floor_list, block_coor, string_coor) for floor_list, block_coor, string_coor in floor_to_coor_set if
                      (beam_coor[0] - block_coor[0][0]) * (beam_coor[0] - block_coor[1][0]) < 0 and
                      (beam_coor[1] - block_coor[0][1]) * (beam_coor[1] - block_coor[1][1]) < 0]
        if len(block_list) > 1:
            block_list = [
                block for block in block_list if block[2][1] < beam_coor[1]]

        if len(block_list) > 1:
            min_floor = min(
                block_list, key=lambda x: get_distance(x[2], beam_coor))[0]

        if len(block_list) == 1:
            min_floor = block_list[0][0]

        floor_list = min_floor

        # 樓層找到之後要去表格對自己的size多大(如果size = ''的話)
        if floor_list is None:
            continue
        for beam_floor in floor_list:
            if not (sizing and mline_scaling):
                set_plan.add((beam_floor, beam_name))
                dic_plan[(beam_floor, beam_name)] = full_coor
                continue
            if beam_size == '':
                # (floor, size_beam, size_string, size_coor)
                beam_size_coor_table = [(floor, size_beam, size_string, size_coor)
                                        for floor, size_beam, size_string, size_coor in
                                        floor_beam_size_coor_set if floor == beam_floor]
                # full match:g1,g2 ;
                # header before "-" include number:GA,GB,DB1,B1
                # header only include char:GA,GB
                # header with first char:Gn,Bn
                for string_pattern in [r".*", r"^\w*", r"^[a-zA-z]*", r"^[a-zA-Z]{2}", r"^[a-zA-z]"]:
                    header_string = re.search(string_pattern, beam_name)
                    if header_string is None:
                        continue

                    header_match = [
                        table for table in beam_size_coor_table if table[1] == header_string.group(0)]
                    if header_match:
                        min_header_match = min(
                            header_match, key=lambda table: get_distance(table[3], beam_coor))
                        beam_size = min_header_match[2]
                        break

            # check is duplicate in plan
            if beam_size != '':
                if (beam_floor, beam_name, '', beam_rotate) in set_plan:
                    set_plan.remove(
                        (beam_floor, beam_name, '', beam_rotate))
                    dic_plan.pop((beam_floor, beam_name, '', beam_rotate))
                    error(
                        f'read_plan error in step 12: {beam_floor} {beam_name} duplicate. ')
                    warning_list.append(
                        f'{beam_floor} {beam_name} duplicate. ')
                set_plan.add(
                    (beam_floor, beam_name, beam_size, beam_rotate))
                dic_plan[(beam_floor, beam_name, beam_size,
                          beam_rotate)] = full_coor
                check_list.append((beam_floor, beam_name))
            else:
                if (beam_floor, beam_name) not in check_list:
                    set_plan.add((beam_floor, beam_name, '', beam_rotate))
                    dic_plan[(beam_floor, beam_name, '',
                              beam_rotate)] = full_coor
                    error(
                        f'read_plan error in step 12: {beam_floor} {beam_name} cannot find size. ')
                    warning_list.append(
                        f'{beam_floor} {beam_name} cannot find size. ')

    progress('平面圖讀取進度 12/13')

    # Step 13. 用 dic_plan((floor, beam_name, beam_size, beam_rotate) -> full_coor) 和 beam_direction_mid_scale_set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale) 找圖是否畫錯
    output_drawing_error_mline_list = []
    if mline_scaling:
        # Step 13-5. 找最近的複線，有錯要畫圖 -> 中點找中點
        for x, item in dic_plan.items():
            if 'x' not in x[2]:
                continue
            beam_scale = float(x[2].split('x')[0])
            beam_coor = item
            beam_layer = sml_beam_layer
            if x[1][0].isupper():
                beam_layer = big_beam_layer
            if 'Fb' in x[1]:
                beam_layer = sml_beam_layer
            beam_rotate = x[3]
            midpoint = ((beam_coor[0][0] + beam_coor[1][0]) / 2,
                        (beam_coor[0][1] + beam_coor[1][1]) / 2)
            min_scale = ''
            min_coor = ''

            if abs(beam_rotate - 1.57) < 0.1:  # 橫的 or 歪的，90度 = pi / 2 = 1.57 (前面有取round到後二位)
                temp_list = [
                    mline for mline in beam_direction_mid_scale_set if mline[1] == 1 and mline[0] in beam_layer]
            elif abs(beam_rotate - 0) < 0.1:  # 直的 or 歪的
                temp_list = [
                    mline for mline in beam_direction_mid_scale_set if mline[1] == 0 and mline[0] in beam_layer]
            else:
                temp_list = [
                    mline for mline in beam_direction_mid_scale_set if mline[0] in beam_layer]

            if len(temp_list) == 0:
                continue

            closet_mline = min(temp_list, key=lambda m: abs(
                midpoint[0] - m[2][0]) + abs(midpoint[1]-m[2][1]))
            min_scale = closet_mline[3]
            min_coor = closet_mline[2]

            # 全部連線
            # coor_list = [min_coor[0], min_coor[1], 0, midpoint[0], midpoint[1], 0]
            # points = vtFloat(coor_list)
            # line = msp_plan.AddPolyline(points)
            # line.SetWidth(0, 3, 3)
            # line.color = 200
            if min_scale == '' or min_scale != beam_scale:
                error_list.append(
                    ((x[0], x[1]), 'mline', f'寬度有誤：文字為{beam_scale}，圖上為{min_scale}。\n'))
                coor = dic_plan[x]
                # 畫框框
                coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                             20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
                output_drawing_error_mline_list.append((coor_list, 130, 10, 4))

                # 只畫有錯的線
                if min_coor != '':
                    coor_list = [min_coor[0], min_coor[1],
                                 0, midpoint[0], midpoint[1], 0]
                    output_drawing_error_mline_list.append(
                        (coor_list, 200, 3, 1))

        # Step 13.6 寫入txt_filename

        error_list.sort(key=lambda x: turn_floor_to_float(x[0]))

    # Step 13.7 把set_plan跟dic_plan調回去
    new_set_plan = set()
    new_dic_plan = {}
    if not sizing:
        for x in dic_plan:
            new_set_plan.add((x[0], x[1]))
            new_dic_plan[(x[0], x[1])] = dic_plan[x]
    else:
        for x in dic_plan:
            new_set_plan.add((x[0], x[1], x[2]))
            new_dic_plan[(x[0], x[1], x[2])] = dic_plan[x]
    set_plan = new_set_plan
    dic_plan = new_dic_plan

    progress('平面圖讀取進度 13/13')
    progress('平面圖讀取完畢。')

    # result_dict['plan'] = sorted(list(set_plan))
    # result_dict['warning'] = warning_list

    return set_plan, dic_plan, warning_list, error_list, output_drawing_error_mline_list


def read_beam(beam_filename, layer_config):
    error_count = 0
    progress('開始讀取梁配筋圖')
    text_layer = layer_config['text_layer']

    doc_beam, msp_beam = activate_cad(beam_filename)
    # Step 4 解鎖所有圖層 -> 不然不能刪東西

    while not doc_beam and error_count <= 10:
        try:
            layer_count = doc_beam.Layers.count

            for x in range(layer_count):
                layer = doc_beam.Layers.Item(x)
                layer.Lock = False
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 4: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 4/9')

    # Step 7. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    progress('正在遍歷梁配筋圖上的物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候')
    floor_to_beam_set = set()
    flag = 0
    count = 0
    used_layer_list = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name
    total = msp_beam.Count
    progress(
        f'梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')

    for msp_object in msp_beam:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'梁配筋圖已讀取{count}/{total}個物件')
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
                    f'read_beam error in step 7-1: {ex}, error_count = {error_count}.')
        while error_count <= 3 and object_list:

            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer

                if object_layer in text_layer and \
                        object.ObjectName == "AcDbText" and\
                        ' ' in object.TextString:
                    pre_beam = (object.TextString.split(' ')[
                                1]).split('(')[0]  # 把括號以後的東西拔掉
                    if pre_beam == '':
                        if re.findall(r'\s{2,}', object.TextString):
                            pre_beam = re.sub(
                                r'\s{2,}', ' ', object.TextString)
                            pre_beam = (pre_beam.split(' ')[
                                1]).split('(')[0]  # 把括號以後的東西拔掉
                        if pre_beam == '':
                            print(object.TextString)
                            continue
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    comma_char = ','
                    for char in weird_comma_list:
                        if char in pre_beam:
                            comma_char = char
                            break
                    comma = pre_beam.count(comma_char)
                    for i in range(comma + 1):
                        beam = pre_beam.split(comma_char)[i]
                        if beam[0] in beam_head1 or beam[0:2] in beam_head2:
                            floor = object.TextString.split(' ')[0]
                            size = (((object.TextString.split('(')[1]).split(')')[0]).replace(
                                ' ', '')).replace('X', 'x')  # size 的格式就是 90x50, 沒空格且使用小寫x作為乘號
                            floor_to_beam_set.add(
                                (floor, beam, (coor1, coor2), size))
            except Exception as ex:
                object_list.append(object)
                error_count += 1
                time.sleep(5)
                error(
                    f'read_beam error in step 7: {ex}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 7/9')
    progress('梁配筋圖讀取進度 8/9')
    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    error_count = 0
    while error_count < 3:
        try:
            doc_beam.Close(SaveChanges=False)
            break
        except Exception as ex:
            error_count += 1
            time.sleep(3)
    # doc_beam.Close(SaveChanges=False)
    progress('梁配筋圖讀取進度 9/9')
    progress('梁配筋圖讀取完成。')
    return floor_to_beam_set


def sort_beam(floor_to_beam_set: set,
              sizing: bool,
              drawing_unit: Literal['cm', 'mm']):
    '''
        result_dict = {
            'beam':[]
        }
    '''
    # Step 8. 算出Bmax, Fmax, Rmax
    Bmax = 0
    Fmax = 0
    Rmax = 0
    for x in floor_to_beam_set:
        floor = x[0]
        tmp_floor_list = []
        to_bool = False
        for char in weird_to_list:
            if char in floor:
                to_char = char
                start = floor.split(to_char)[0]
                end = floor.split(to_char)[1]
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

    # Step 9. 完成set_beam和dic_beam
    dic_beam = {}
    set_beam = set()
    for x in floor_to_beam_set:
        floor: str = x[0]
        beam = x[1]
        coor = x[2]
        size = x[3]
        floor_list = []
        to_bool = False

        if drawing_unit == 'mm':
            beam_size = convert_mm_to_cm(beam_size)

        for char in weird_to_list:
            if char in floor:
                to_char = char
                start = floor.split(to_char)[0]
                end = floor.split(to_char)[1]
                try:
                    start = int(turn_floor_to_float(start))
                    end = int(turn_floor_to_float(end))
                    if start > end:
                        tmp = start
                        start = end
                        end = tmp
                    for i in range(start, end + 1):
                        if floor_exist(i, Bmax, Fmax, Rmax):
                            floor_list.append(turn_floor_to_string(i))
                except Exception:
                    error('read_beam error in step 9: The error above is from here.')
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
                new_floor = floor.split(comma_char)[i]
                new_floor = turn_floor_to_float(new_floor)
                new_floor = turn_floor_to_string(new_floor)
                if new_floor:
                    floor_list.append(new_floor)
                else:
                    error(f'read_beam error in step 9: new_floor is false.')

        for floor in floor_list:
            if sizing:
                if (floor, beam, size) in set_beam:
                    set_beam.add((floor, beam, 'replicate'))
                else:
                    set_beam.add((floor, beam, size))
                    dic_beam[(floor, beam, size)] = coor
            else:
                set_beam.add((floor, beam))
                dic_beam[(floor, beam)] = coor

    # result_dict['beam'] = sorted(list(set_beam))
    # beam.txt單純debug用，不想多新增檔案可以註解掉
    # with open(result_filename, "w") as f:
    #     f.write("in beam: \n")
    #     l = list(set_beam)
    #     l.sort()
    #     for x in l:
    #         f.write(f'{x}\n')

    return (set_beam, dic_beam)


# 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
def write_plan(plan_filename,
               plan_new_filename,
               set_plan,
               set_beam,
               dic_plan,
               date,
               drawing,
               output_drawing_error_mline_list: list[tuple],
               client_id) -> list:

    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    error_count = 0
    progress("開始標註平面圖(核對項目: 梁配筋)及輸出核對結果。")
    pythoncom.CoInitialize()
    set_in_plan = set_plan - set_beam
    list_in_plan = list(set_in_plan)
    list_in_plan.sort()
    set_in_beam = set_beam - set_plan
    list_in_beam = list(set_in_beam)
    list_in_beam = [beam for beam in list_in_beam if beam[2] != 'replicate']
    list_in_beam.sort()

    # Step 5. 完成in plan but not in beam，畫圖，以及計算錯誤率
    error_list = []
    for plan_beam in list_in_plan:
        error_beam = [beam_beam for beam_beam in list_in_beam if plan_beam[0]
                      == beam_beam[0] and plan_beam[1] == beam_beam[1]]
        beam_floor, beam_name, beam_size = plan_beam
        beam_drawing = 0
        if error_beam:
            if beam_size != '':
                error_list.append((plan_beam, 'error_size', error_beam[0][2]))
                beam_drawing = 1
            else:
                error_list.append((plan_beam, 'no_size', error_beam[0][2]))
        else:
            beam_drawing = 1
            error_list.append((plan_beam, 'no_beam', ''))

        if drawing and beam_drawing:
            coor = dic_plan[plan_beam]
            coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                         20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
            points = vtFloat(coor_list)
            output_drawing_error_mline_list.append((coor_list, 0, 10, 4))
            # pointobj = msp_plan.AddPolyline(points)
            # for i in range(4):
            #     pointobj.SetWidth(i, 10, 10)

    if drawing:
        document, model_space = activate_cad(plan_filename)
        # Step 13-4. 設定mark的圖層
        error_count = 0
        while document and error_count <= 10:
            try:
                layer_plan = document.Layers.Add(f"S-CLOUD_{date}")
                document.ActiveLayer = layer_plan
                layer_plan.color = 10
                layer_plan.Linetype = "Continuous"
                layer_plan.Lineweight = 0.5
                break
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 13-4, {ex}, error_count = {error_count}')
        if model_space:
            for error_mline in output_drawing_error_mline_list:
                coor_list, color, width, border = error_mline
                try:
                    points = vtFloat(coor_list)
                    ply = model_space.AddPolyline(points)
                    for i in range(border):
                        ply.SetWidth(i, width, width)
                    if color:
                        ply.color = color
                except:
                    continue
            try:
                document.SaveAs(plan_new_filename)
                document.Close(SaveChanges=True)
            except:
                error(f'Cant Save Output {plan_new_filename} File')

        # 在這之後就沒有while迴圈了，所以錯超過10次就出去
        if error_count > 10:
            try:
                document.Close(SaveChanges=False)
            except:
                error('Close File Error')

    # if drawing:
    #     doc_plan.SaveAs(plan_new_filename)
    #     doc_plan.Close(SaveChanges=True)
    return error_list


def output_error_list(error_list: list, title_text: str, set_item=set, cad_data=[]):
    result_dict = {
        'mline': {
            '大梁': [],
            '小梁': [],
            '地梁': [],
        },
        'size': {
            '大梁': [],
            '小梁': [],
            '地梁': [],
        },
        'no_size': {
            '大梁': [],
            '小梁': [],
            '地梁': [],
        },
        'not_found': {
            '大梁': [],
            '小梁': [],
            '地梁': [],
        },
        'duplicate': {
            '大梁': [],
            '小梁': [],
            '地梁': [],
        },
        'summary': {
            '大梁': {},
            '小梁': {},
            '地梁': {}
        },
        'item': [],
        'warning': [],
        'cad_data': []
    }
    error_counter = {
        '大梁': {
            'mline': [],
            'size': [],
            'not_found': [],
            'no_size': [],
            'duplicate': []
        },
        '小梁': {
            'mline': [],
            'size': [],
            'not_found': [],
            'no_size': [],
            'duplicate': []
        },
        '地梁': {
            'mline': [],
            'size': [],
            'not_found': [],
            'no_size': [],
            'duplicate': []
        }
    }

    # error_list = sorted(error_list,key = cmp_to_key(mycmp))
    mline_error = [e for e in error_list if e[1] == 'mline']
    error_size = [e for e in error_list if e[1] == 'error_size']
    no_size = [e for e in error_list if e[1] == 'no_size']
    no_beam = [e for e in error_list if e[1] == 'no_beam']
    replicate_beam = [e for e in error_list if e[1] == 'replicate']

    mline_error = sorted(mline_error, key=cmp_to_key(mycmp))
    error_size = sorted(error_size, key=cmp_to_key(mycmp))
    no_size = sorted(no_size, key=cmp_to_key(mycmp))
    no_beam = sorted(no_beam, key=cmp_to_key(mycmp))
    replicate_beam = sorted(replicate_beam, key=cmp_to_key(mycmp))

    beam_list = [b for b in set_item if b[1][0] ==
                 'B' or b[1][0] == 'C' or b[1][0] == 'G']
    fbeam_list = [b for b in set_item if b[1][0] == 'F']
    sbeam_list = [
        b for b in set_item if b not in beam_list and b not in fbeam_list]

    for error_type, error_result in [('mline', mline_error), ('size', error_size), ('not_found', no_beam), ('no_size', no_size), ('duplicate', replicate_beam)]:

        for e in error_result:
            beam = e[0]
            beam_name = beam[1]

            beam_type = ""

            if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G':
                beam_type = "大梁"
            elif beam_name[0] == 'F':
                beam_type = "地梁"
            else:
                beam_type = "小梁"
            if error_type == 'mline':
                error_message = e[2]
            if error_type == 'size':
                beam_size = beam[2]
                error_message = f"{beam_size} , 在{title_text}是{e[2]}"
            if error_type == 'not_found':
                error_message = f"在{title_text}找不到這根梁"
            if error_type == 'no_size':
                error_message = "找不到尺寸"
            if error_type == 'duplicate':
                error_message = "重複配筋"
            error_counter[beam_type][error_type].append(e)
            result_dict[error_type][beam_type].append(
                (beam[0], beam_name, error_message))

    if cad_data:
        result_dict['cad_data'] = cad_data

    if set_item:
        result_dict['item'] = sorted(
            list(set_item), key=lambda item: turn_floor_to_float(item[0]))

    for (item_list, type_name) in zip([beam_list, sbeam_list, fbeam_list], ['大梁', '小梁', '地梁']):
        if item_list:
            error_count = len(
                error_counter[type_name]['size']) + len(error_counter[type_name]['not_found'])
            result_dict['summary'].update({type_name: {
                '尺寸錯誤': len(error_counter[type_name]['size']),
                '缺少配筋': len(error_counter[type_name]['not_found']),
                '總共': len(item_list),
                '錯誤率': f'{round(error_count / len(item_list),2) * 100}%'
            }})
        else:
            result_dict['summary'].update({type_name: {
                '備註': f'平面圖中無{type_name}'
            }})

    progress('平面圖標註進度 5/5')
    progress("標註平面圖(核對項目: 梁配筋)及輸出核對結果完成。")
    return error_counter, result_dict


# 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
def write_beam(beam_filename,
               beam_new_filename,
               set_plan,
               set_beam,
               dic_beam,
               date,
               drawing,
               client_id):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    try:
        error_count = 0
        progress("開始標註梁配筋圖及輸出核對結果")
        pythoncom.CoInitialize()
        set1 = set_plan - set_beam
        list_in_plan = list(set1)
        list_in_plan.sort()
        set2 = set_beam - set_plan
        list_in_beam = list(set2)
        list_in_beam.sort()
        error_list = []

        if drawing:
            # Step 1. 開啟應用程式
            flag = 0
            while not flag and error_count <= 10:
                try:
                    wincad_beam = win32com.client.Dispatch(
                        "AutoCAD.Application")
                    flag = 1
                except Exception as e:
                    error_count += 1
                    time.sleep(5)
                    error(
                        f'write_beam error in step 1, {e}, error_count = {error_count}.')
            progress('梁配筋圖標註進度 1/5')
            # Step 2. 匯入檔案
            flag = 0
            while not flag and error_count <= 10:
                try:
                    doc_beam = wincad_beam.Documents.Open(beam_filename)
                    flag = 1
                except Exception as e:
                    error_count += 1
                    time.sleep(5)
                    error(
                        f'write_beam error in step 2, {e}, error_count = {error_count}.')
            progress('梁配筋圖標註進度 2/5')
            # Step 3. 載入modelspace(還要畫圖)
            flag = 0
            while not flag and error_count <= 10:
                try:
                    msp_beam = doc_beam.Modelspace
                    flag = 1
                except Exception as e:
                    error_count += 1
                    time.sleep(5)
                    error(
                        f'write_beam error in step 3, {e}, error_count = {error_count}.')
            time.sleep(5)
            progress('梁配筋圖標註進度 3/5')
            # Step 4. 設定mark的圖層
            flag = 0
            while not flag and error_count <= 10:
                try:
                    layer_beam = doc_beam.Layers.Add(f"S-CLOUD_{date}")
                    doc_beam.ActiveLayer = layer_beam
                    layer_beam.color = 10
                    layer_beam.Linetype = "Continuous"
                    layer_beam.Lineweight = 0.5
                    flag = 1
                except Exception as e:
                    error_count += 1
                    time.sleep(5)
                    error(
                        f'write_beam error in step 4, {e}, error_count = {error_count}.')
        progress('梁配筋圖標註進度 4/5')

        # 在這之後就沒有while迴圈了，所以錯超過10次就出去
        if error_count > 10:
            try:
                doc_beam.Close(SaveChanges=False)
            except:
                pass
            return False

        # Step 5. 完成in plan but not in beam，畫圖，以及計算錯誤率
        error_list = []
        for beam_beam in list_in_beam:
            error_beam = [plan_beam for plan_beam in list_in_plan if plan_beam[0]
                          == beam_beam[0] and plan_beam[1] == beam_beam[1]]
            beam_floor, beam_name, beam_size = beam_beam
            beam_drawing = 0
            if beam_size == 'replicate':
                error_list.append((beam_beam, 'replicate', 'replicate'))
                continue
            if error_beam:
                if beam_size != '':
                    error_list.append(
                        (beam_beam, 'error_size', error_beam[0][2]))
                    beam_drawing = 1
                else:
                    error_list.append((beam_beam, 'no_size', error_beam[0][2]))
            else:
                beam_drawing = 1
                error_list.append((beam_beam, 'no_beam', ''))

            if drawing and beam_drawing:
                coor = dic_beam[beam_beam]
                coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                             20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
                points = vtFloat(coor_list)
                pointobj = msp_beam.AddPolyline(points)
                for i in range(4):
                    pointobj.SetWidth(i, 10, 10)

        if drawing:
            doc_beam.SaveAs(beam_new_filename)
            doc_beam.Close(SaveChanges=True)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

    return error_list


def write_result_log(task_name, plan_result: dict[str, dict], beam_result: dict[str, dict]):
    # Step 1: Normalize the mixed data into a uniform dictionary structure
    def normalize_dict(cad_data):
        normalized_data = {}
        if isinstance(cad_data, dict):
            for key, value in cad_data.items():
                if isinstance(value, dict):
                    normalized_data[key] = value
                elif isinstance(value, list):
                    normalized_data[key] = {
                        f"item_{i}": v for i, v in enumerate(value)}
                else:
                    normalized_data[key] = {"value": value}
        return normalized_data
    plan_beam_df_list = []
    plan_sbeam_df_list = []
    plan_fbeam_df_list = []
    for error_type in ['mline', 'size', 'not_found', 'no_size', 'duplicate']:
        for beam_type, content in plan_result[error_type].items():
            if not content:
                continue
            df = pd.DataFrame(content, columns=['樓層', '編號', '錯誤'])
            if beam_type == '大梁':
                plan_beam_df_list.append(df)
            if beam_type == '小梁':
                plan_sbeam_df_list.append(df)
            if beam_type == '地梁':
                plan_fbeam_df_list.append(df)

    beam_beam_df_list = []
    beam_sbeam_df_list = []
    beam_fbeam_df_list = []
    for error_type in ['not_found', 'size', 'no_size', 'duplicate']:
        for beam_type, content in beam_result[error_type].items():
            if not content:
                continue
            # if beam_type not in beam_result:
            #     continue
            df = pd.DataFrame(content, columns=['樓層', '編號', '錯誤'])
            if beam_type == '大梁':
                beam_beam_df_list.append(df)
            if beam_type == '小梁':
                beam_sbeam_df_list.append(df)
            if beam_type == '地梁':
                beam_fbeam_df_list.append(df)

    return {
        'XS-PLAN 統整': [pd.DataFrame.from_dict(plan_result['summary'], orient='index')],
        'XS-PLAN 大梁結果': plan_beam_df_list,
        'XS-PLAN 小梁結果': plan_sbeam_df_list,
        'XS-PLAN 地梁結果': plan_fbeam_df_list,
        'XS-BEAM 統整': [pd.DataFrame.from_dict(plan_result['summary'], orient='index')],
        'XS-BEAM 大梁結果': beam_beam_df_list,
        'XS-BEAM 小梁結果': beam_sbeam_df_list,
        'XS-BEAM 地梁結果': beam_fbeam_df_list,
        'XS-PLAN 詳細內容': [pd.DataFrame(plan_result['item'] if 'item' in plan_result else {}, columns=['floor', 'serial', 'size'])],
        'XS-BEAM 詳細內容': [pd.DataFrame(beam_result['item'] if 'item' in beam_result else {}, columns=['floor', 'serial', 'size'])],
        'CAD data': [pd.DataFrame.from_dict(normalize_dict(plan_result['cad_data'])if 'cad_data' in plan_result else {}, orient='columns')]
    }


def run_plan(plan_filename,
             layer_config: dict,
             sizing,
             mline_scaling,
             client_id,
             drawing_unit,
             pkl=""):
    start_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    if pkl == "":
        plan_data = read_plan(plan_filename=plan_filename,
                              layer_config=layer_config,
                              sizing=sizing,
                              mline_scaling=mline_scaling)
        save_temp_file.save_pkl(
            data=plan_data, tmp_file=f'{os.path.splitext(plan_filename)[0]}_plan_set.pkl')
    else:
        plan_data = save_temp_file.read_temp(
            tmp_file=pkl)
    set_plan, dic_plan, warning_list, \
        mline_error_list, drawing_error_list = sort_plan(plan_data=plan_data,
                                                         layer_config=layer_config,
                                                         sizing=sizing,
                                                         mline_scaling=mline_scaling,
                                                         drawing_unit=drawing_unit)
    end_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    cad_data = output_progress_report(layer_config=layer_config,
                                      start_date=start_date,
                                      end_date=end_date,
                                      project_name=plan_filename,
                                      warning_list=warning_list,
                                      plan_filename=plan_filename,
                                      plan_data=plan_data)

    return (set_plan, dic_plan, mline_error_list, cad_data, drawing_error_list)


def run_beam(beam_filename,
             layer_config,
             sizing,
             client_id,
             drawing_unit,
             pkl: str = ""):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    if pkl == "":
        floor_to_beam_set = read_beam(
            beam_filename=beam_filename, layer_config=layer_config)
        save_temp_file.save_pkl(data=floor_to_beam_set,
                                tmp_file=f'{os.path.splitext(beam_filename)[0]}_beam_set.pkl')
    else:
        floor_to_beam_set = save_temp_file.read_temp(pkl)
    set_beam, dic_beam = sort_beam(floor_to_beam_set=floor_to_beam_set,
                                   sizing=sizing,
                                   drawing_unit=drawing_unit)
    return (set_beam, dic_beam)


def output_progress_report(layer_config: dict, start_date, end_date, plan_data: dict, project_name: str, plan_filename: str, warning_list: list):
    delimiter = '\n'
    cad_data = {
        '專案名稱': project_name,
        '平面圖名稱': plan_filename,
        '開始時間': start_date,
        '結束時間': end_date,
        '圖層參數': layer_config,
        '平面圖樓層': plan_data["coor_to_floor_set"],
        '錯誤訊息': delimiter.join(warning_list),
        '圖框': len(plan_data['block_coor_list']),
        '平面圖梁編號': len(plan_data['coor_to_beam_set']),
        '表格梁編號': len(plan_data['coor_to_size_beam']),
        '表格梁尺寸': len(plan_data['coor_to_size_string'])
    }
    return cad_data
    # with open(output_filename, 'w') as f:
    #     f.write(
    #         f'專案名稱:{project_name}\n平面圖名稱:{plan_filename}\n開始時間:{start_date} \n結束時間:{end_date} \n')
    #     f.write(f'圖層參數:{layer_config} \n')
    #     f.write(f'CAD資料:{cad_data}]\n')
    #     f.write(f'平面圖樓層:{plan_data["coor_to_floor_set"]}]\n')
    #     f.write(f'==========================\n')
    #     f.write('錯誤訊息:\n')
    #     f.write(f'{delimiter.join(warning_list)}')


if __name__ == '__main__':
    # # 在beam裡面自訂圖層
    text_layer = ['S-RC']  # sys.argv[7]

    # 在plan裡面自訂圖層
    block_layer = ['DwFm', '0', 'DEFPOINTS']  # sys.argv[8] # 框框的圖層
    floor_layer = ['S-TITLE']  # sys.argv[9] # 樓層字串的圖層
    size_layer = ['S-TEXT']  # sys.argv[12] # 梁尺寸字串圖層
    big_beam_layer = ['S-RCBMG']  # 大樑複線圖層
    big_beam_text_layer = ['S-TEXTG']  # 大樑文字圖層
    sml_beam_layer = ['S-RCBMB']  # 小梁複線圖層
    sml_beam_text_layer = ['S-TEXTB']  # 小梁文字圖層

    block_layer = ['DwFm', '0', 'DEFPOINTS', 'FRAME']  # sys.argv[8] # 框框的圖層
    floor_layer = ['S-TITLE']  # sys.argv[9] # 樓層字串的圖層
    size_layer = ['S-TEXT']  # sys.argv[12] # 梁尺寸字串圖層
    big_beam_layer = ['S-RCBMG']  # 大樑複線圖層
    big_beam_text_layer = ['S-TEXTG']  # 大樑文字圖層
    sml_beam_layer = ['S-RCBMB']  # 小梁複線圖層
    sml_beam_text_layer = ['S-TEXTB']  # 小梁文字圖層

    task_name = '1128-逢大段'  # sys.argv[13]

    progress_file = './result/tmp'  # sys.argv[14]
    output_folder = r'D:\Desktop\BeamQC\TEST\2024-1128'

    sizing = 1  # 要不要對尺寸
    mline_scaling = 1  # 要不要對複線寬度

    date = time.strftime("%Y-%m-%d", time.localtime())
    layer_config = {
        # 'line_layer':line_layer,
        'text_layer': text_layer,
        'block_layer': block_layer,
        'floor_layer': floor_layer,
        'big_beam_layer': big_beam_layer,
        'big_beam_text_layer': big_beam_text_layer,
        'sml_beam_layer': sml_beam_layer,
        'size_layer': size_layer,
        'sml_beam_text_layer': sml_beam_text_layer
    }
    pkls = [r'TEST\2024-1024\2024-11-08-12-37_2024-1108-2024-11-08-10-10_1-XS-BEAM_beam_set.pkl']
    plan_filename = r'D:\Desktop\BeamQC\TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-PLAN.dwg'
    plan_new_filename = f'{output_folder}\\勝利CDE.dwg'
    set_beam_all = set()

    set_plan, dic_plan, \
        plan_mline_error_list, plan_cad_data_list, \
        drawing_error_list = run_plan(plan_filename=plan_filename,
                                      layer_config=layer_config,
                                      drawing_unit='cm',
                                      sizing=True,
                                      mline_scaling=True,
                                      client_id='2024-1018',
                                      pkl=r'TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-PLAN_plan_set.pkl')
    set_beam, dic_beam = run_beam(
        beam_filename=r'D:\Desktop\BeamQC\TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-BEAM.dwg',
        layer_config=layer_config,
        sizing=True,
        client_id='2024-1018',
        drawing_unit='cm',
        pkl=r'TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-BEAM_beam_set.pkl'
    )
    # for pkl in pkls:
    #     floor_to_beam_set = save_temp_file.read_temp(pkl)
    #     set_beam, dic_beam = sort_beam(floor_to_beam_set=floor_to_beam_set,
    #                                    drawing_unit='cm',
    #                                    sizing=True)
    #     set_beam_all = set_beam | set_beam_all

    plan_error_list = write_plan(plan_filename=plan_filename,
                                 plan_new_filename=plan_new_filename,
                                 set_plan=set_plan,
                                 set_beam=set_beam,
                                 dic_plan=dic_plan,
                                 date=date,
                                 drawing=False,
                                 output_drawing_error_mline_list=drawing_error_list,
                                 client_id='temp')
    plan_error_list.extend(plan_mline_error_list)

    beam_error_list = write_beam(
        beam_filename=r'D:\Desktop\BeamQC\TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-BEAM.dwg',
        beam_new_filename=r'D:\Desktop\BeamQC\TEST\2024-1128\2024-11-28-16-38_2024-1128 逢大段-XS-BEAM_new.dwg',
        set_plan=set_plan,
        set_beam=set_beam,
        dic_beam=dic_beam,
        date=date,
        drawing=False,
        client_id='temp'
    )
    plan_error_counter, plan_result_dict = output_error_list(error_list=plan_error_list,
                                                             title_text='XS-BEAM',
                                                             set_item=set_plan,
                                                             cad_data=plan_cad_data_list)

    beam_error_counter, beam_result_dict = output_error_list(error_list=beam_error_list,
                                                             title_text='XS-PLAN',
                                                             set_item=set_beam)

    data_excel_file = os.path.join(
        output_folder, f'{task_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_梁Check_結果.xlsx')

    from src.main import OutputExcel

    output_data = write_result_log(task_name=task_name,
                                   plan_result=plan_result_dict,
                                   beam_result=beam_result_dict
                                   )
    for sheet_name, df_list in output_data.items():
        OutputExcel(df_list=df_list,
                    df_spacing=1,
                    file_path=data_excel_file,
                    sheet_name=sheet_name)
