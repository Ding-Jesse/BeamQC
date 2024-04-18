import win32com.client
import pythoncom
import re
import time
import multiprocessing
import os
import pandas as pd
from functools import cmp_to_key
import save_temp_file
import json
from gzip import READ
from io import TextIOWrapper
from math import inf, sqrt
from multiprocessing.spawn import prepare
from tabnanny import check
from tkinter import HIDDEN
from numpy import object_
from openpyxl import load_workbook
from collections import Counter


weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡B']
beam_head1 = ['B', 'b', 'G', 'g']
beam_head2 = ['FB', 'FG', 'Fb', 'CB', 'CG', 'cb']


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
        error(f'turn_floor_to_float error: {floor} cannot be turned to float.')
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


def turn_floor_to_list(floor, Bmax, Fmax, Rmax):
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
    f = open(error_file, 'a', encoding='utf-8')
    localtime = time.asctime(time.localtime(time.time()))
    f.write(f'{localtime} | {error_message}\n')
    f.close()
    return


def progress(message, progress_file):
    '''把進度印到progress裡面，在app.py會對這個檔案做事'''
    f = open(progress_file, 'a', encoding='utf-8')
    f.write(f'{message}\n')
    f.close()
    return

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


def read_plan(plan_filename, layer_config: dict, progress_file, sizing, mline_scaling):
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
    progress('開始讀取平面圖(核對項目: 梁配筋對應)', progress_file)
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 1: {ex} ,error in open Autocad, error_count = {error_count}.')
    progress('平面圖讀取進度 1/13', progress_file)

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 2: {ex} ,error in open dwg file , error_count = {error_count}.')
    progress('平面圖讀取進度 2/13', progress_file)

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as ex:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 3: {ex} ,error in reading ModelSpace, error_count = {error_count}.')
    progress('平面圖讀取進度 3/13', progress_file)

    # Step 4 解鎖所有圖層 -> 不然不能刪東西
    flag = 0
    while not flag and error_count <= 10:
        try:
            layer_count = doc_plan.Layers.count
            for x in range(layer_count):
                layer = doc_plan.Layers.Item(x)
                layer.Lock = False
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_plan error in step 4: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 4/13', progress_file)

    # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間
    # flag = 0
    # time.sleep(5)
    # layer_list = [floor_layer, size_layer, big_beam_layer,
    #               big_beam_text_layer, sml_beam_layer, sml_beam_text_layer]
    # non_trash_list = layer_list + [block_layer]
    # while not flag and error_count <= 0:
    #     try:
    #         count = 0
    #         total = msp_plan.Count
    #         progress(
    #             f'正在炸平面圖的圖塊及篩選判斷用的物件，平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
    #         for object in msp_plan:
    #             explode_fail = 0
    #             while explode_fail <= 3:
    #                 try:
    #                     count += 1
    #                     if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:  # block 不會自動炸，需要手動修正
    #                         object.Explode()
    #                     if object.Layer not in non_trash_list:
    #                         object.Delete()
    #                     if count % 1000 == 0:
    #                         # 每1000個跳一次，確認有在動
    #                         progress(
    #                             f'平面圖已讀取{count}/{total}個物件', progress_file)
    #                     break
    #                 except:
    #                     explode_fail += 1
    #                     time.sleep(5)
    #                     try:
    #                         msp_plan = doc_plan.Modelspace
    #                     except:
    #                         pass
    #         flag = 1

    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(
    #             f'read_plan error in step 5: {e}, error_count = {error_count}.')
    #         try:
    #             msp_plan = doc_plan.Modelspace
    #         except:
    #             pass

    # progress('平面圖讀取進度 5/13', progress_file)

    # Step 6. 重新匯入modelspace，剛剛炸出來的東西要重新讀一次
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         msp_plan = doc_plan.Modelspace
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(
    #             f'read_plan error in step 6: {e}, error_count = {error_count}.')
    # progress('平面圖讀取進度 6/13', progress_file)

    # Step 7. 遍歷所有物件 -> 完成各種我們要的set跟list

    progress('正在遍歷平面圖上的物件並篩選出有效信息，運行時間取決於平面圖大小，請耐心等候', progress_file)
    flag = 0
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
        f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
    for msp_object in msp_plan:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'平面圖已讀取{count}/{total}個物件', progress_file)
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
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer
                # 找size
                if sizing or mline_scaling:
                    if object_layer in size_layer and object.EntityName in text_object_type and object.TextString != '' and object.GetBoundingBox()[0][1] >= 0:
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
                        if 'Bn' in object.TextString and 'W' not in object.TextString and 'D' not in object.TextString and 'FBn' not in object.TextString:
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
                        if 'x' in object.TextString or 'X' in object.TextString or 'x' in object.TextString:
                            string = (object.TextString.replace(
                                ' ', '')).replace('X', 'x')
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
                if object_layer in floor_layer and object.EntityName in text_object_type and '(' in object.TextString and object.InsertionPoint[1] >= 0:
                    floor = object.TextString
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
                                ((coor1, coor2), size.group(0)))
                            continue
                        else:
                            print(object.TextString)
                if object.Layer in beam_text_layer and object.EntityName in text_object_type\
                        and object.TextString != '' and (object.TextString[0] in beam_head1 or object.TextString[0:2] in beam_head2):

                    beam = object.TextString

                    if ' ' in beam:
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
                            size = beam_size.group(0).replace('X', 'x')
                        if beam_name:
                            beam = beam_name.group(0)
                            # beam = beam.replace('X','x')
                    # if '(' in beam:
                    #     size = (((beam.split('(')[1]).split(')')[0]).replace(' ', '')).replace('X', 'x')
                    #     if 'x' not in size:
                    #         size = ''
                    #     else:
                    #         try:
                    #             first = size.split('x')[0]
                    #             second = size.split('x')[1]
                    #             if not (float(first) and float(second)):
                    #                 size = ''
                    #         except:
                    #             size = ''
                    #     beam = beam.split('(')[0] # 取括號前內容即可
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
                    beam = object.TextString
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

                if object_layer in block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
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
        pass
    # return False
    progress('平面圖讀取進度 7/13', progress_file)
    # doc_plan.Close(SaveChanges=False)
    return {'coor_to_floor_set': coor_to_floor_set,
            'coor_to_beam_set': coor_to_beam_set,
            'block_coor_list': block_coor_list,
            'none_concat_size_text_list': none_concat_size_text_list,
            'coor_to_size_beam': coor_to_size_beam,
            'coor_to_size_string': coor_to_size_string,
            'beam_direction_mid_scale_set': beam_direction_mid_scale_set}


def sort_plan(plan_filename: str, plan_new_filename: str, layer_config: dict, plan_data: dict, sizing: bool, mline_scaling: bool, big_file: str, sml_file: str, result_filename: str, progress_file: str, date, fbeam_file: str):
    error_count = 0
    warning_list = []

    def get_distance(coor1, coor2):
        if isinstance(coor1, tuple) and isinstance(coor2, tuple):
            return abs(coor1[0][0]-coor2[0][0]) + abs(coor1[0][1]-coor2[0][1])
        return 10000
    # 2023-0308
    # set (字串的coor, floor)，Ex. 求'1F'這個字串的座標在哪
    coor_to_floor_set = plan_data['coor_to_floor_set']
    # set (coor, [beam, size])，Ex. 求'B1-6'這個字串的座標在哪，如果後面有括號的話，順便紀錄尺寸，否則size = ''
    coor_to_beam_set = plan_data['coor_to_beam_set']
    block_coor_list = plan_data['block_coor_list']  # 存取方框最左下角的點座標
    none_concat_size_text_list = plan_data['none_concat_size_text_list']
    # for sizing
    # set (coor, size_beam)，Ex. 紀錄表格中'Bn'這個字串的座標
    coor_to_size_beam = plan_data['coor_to_size_beam']
    # set (coor, size_string)，Ex. 紀錄表格中'25x50'這個字串的座標
    coor_to_size_string = plan_data['coor_to_size_string']

    big_beam_layer = layer_config['big_beam_layer']
    sml_beam_layer = layer_config['sml_beam_layer']

    # for mline_scaling
    # set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale)
    beam_direction_mid_scale_set = plan_data['beam_direction_mid_scale_set']
    # 2023-0119
    for none_concat_size in none_concat_size_text_list:
        coor, size = none_concat_size
        temp_list = [s for s in coor_to_beam_set if s[1][1] == '']
        closet_beam = min(temp_list, key=lambda x: get_distance(x[0], coor))
        coor_to_beam_set.remove(closet_beam)
        coor_to_beam_set.add(
            (closet_beam[0], (closet_beam[1][0].strip(), size, closet_beam[1][2])))
        # closet_beam[1][1] = size

    # Step 8. 完成size_coor_set (size_beam, size_string, size_coor), Ex. 把表格中的 'Bn' 跟 '50x70' 連起來

    if sizing or mline_scaling:
        size_coor_set = set()
        for x in coor_to_size_beam:
            coor = x[0]
            size_beam = x[1]
            min_size = ''
            min_dist = inf
            for y in coor_to_size_string:
                coor2 = y[0]
                size_string = y[1]
                dist = abs(coor[0]-coor2[0]) + abs(coor[1] - coor2[1])
                if dist < min_dist:
                    min_size = size_string
                    min_dist = dist
            if min_size != '':
                size_coor_set.add((size_beam, min_size, coor))
    progress('平面圖讀取進度 8/13', progress_file)

    # Step 9. 透過 coor_to_floor_set 以及 block_coor_list 完成 floor_to_coor_set，格式為(floor, block左下角和右上角的coor), Ex. '1F' 左下角和右上角的座標分別為(0.0, 0.0) (14275.54, 10824.61)
    # 此處不會報錯，沒在框框裡就直接扔了
    floor_to_coor_set = set()
    for x in coor_to_floor_set:  # set (字串的coor, floor)
        string_coor = x[0]
        floor = x[1]
        for block_coor in block_coor_list:
            x_diff_left = string_coor[0] - block_coor[0][0]  # 和左下角的diff
            y_diff_left = string_coor[1] - block_coor[0][1]
            x_diff_right = string_coor[0] - block_coor[1][0]  # 和右上角的diff
            y_diff_right = string_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:  # 要在框框裡面才算
                floor_to_coor_set.add((floor, block_coor))
    progress('平面圖讀取進度 9/13', progress_file)

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
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_floor_to_coor_list.append((floor_list, block_coor))

    floor_to_coor_set = new_floor_to_coor_list

    new_coor_to_floor_list = []
    for x in coor_to_floor_set:
        string_coor = x[0]
        floor_name = x[1]
        floor_list = turn_floor_to_list(floor_name, Bmax, Fmax, Rmax)
        if len(floor_list) != 0:
            new_coor_to_floor_list.append((string_coor, floor_list))

    coor_to_floor_set = new_coor_to_floor_list

    progress('平面圖讀取進度 10/13', progress_file)

    # Step 11. 完成floor_beam_size_coor_set (floor, beam, size, coor), 找表格內的物件在哪一個框框裡面，進而找到所屬樓層

    if sizing or mline_scaling:
        floor_beam_size_coor_set = set()
        for x in size_coor_set:  # set(size_beam, min_size, coor)
            size_coor = x[2]
            size_string = x[1]
            size_beam = x[0]
            min_floor = []
            for z in floor_to_coor_set:  # list (floor_list, block左下角和右上角的coor)
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
    progress('平面圖讀取進度 11/13', progress_file)

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
        min_floor = ''

        # 我其實是list歐哈哈 (floor_list, block左下角和右上角的coor)
        for z in floor_to_coor_set:
            floor_list = z[0]
            block_coor = z[1]
            x_diff_left = beam_coor[0] - block_coor[0][0]  # 和左下角的diff
            y_diff_left = beam_coor[1] - block_coor[0][1]
            x_diff_right = beam_coor[0] - block_coor[1][0]  # 和右上角的diff
            y_diff_right = beam_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:
                if min_floor == '' or min_floor == floor_list:
                    min_floor = floor_list

                else:  # 有很多層在同一個block, 仍然透過字串的coor找樓層 -> 應從已知選項找最適合的，而不是全部重找，這樣會找到框框外面的東西
                    for y in coor_to_floor_set:  # set(字串的coor, floor)
                        if y[1] == min_floor:
                            string_coor = y[0]
                            x_diff = abs(beam_coor[0] - string_coor[0])
                            y_diff = beam_coor[1] - string_coor[1]
                            total = x_diff + y_diff
                        if y[1] == floor_list:
                            string_coor = y[0]
                            new_x_diff = abs(beam_coor[0] - string_coor[0])
                            new_y_diff = beam_coor[1] - string_coor[1]
                            new_total = new_x_diff + new_y_diff
                    if (new_y_diff > 0 and y_diff > 0 and new_total < total) or y_diff < 0:
                        min_floor = floor_list

        floor_list = min_floor

        # 樓層找到之後要去表格對自己的size多大(如果size = ''的話)

        if floor_list != '':
            for floor in floor_list:
                if sizing or mline_scaling:
                    if beam_size == '':
                        min_diff = inf
                        min_size = ''
                        for y in floor_beam_size_coor_set:
                            if y[0] == floor and y[1] == beam_name:  # 先看有沒有像g1, g2完全吻合的
                                tmp_coor = y[3]
                                diff = abs(
                                    beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
                                if diff < min_diff:
                                    min_diff = diff
                                    min_size = y[2]
                        if min_size != '':
                            beam_size = min_size
                        else:
                            header = beam_name
                            for char in header:
                                if char.isdigit():
                                    header = header.split(char)[0]
                                    break
                            for y in floor_beam_size_coor_set:
                                # 再看頭一不一樣(去掉數字的頭)
                                if y[0] == floor and y[1] == header:
                                    tmp_coor = y[3]
                                    diff = abs(
                                        beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
                                    if diff < min_diff:
                                        min_diff = diff
                                        min_size = y[2]
                        if min_size != '':
                            beam_size = min_size
                        else:
                            for y in floor_beam_size_coor_set:
                                # 再看頭有沒有包含到(ex. GA-1 算在 G)
                                if y[0] == floor and y[1] in header:
                                    tmp_coor = y[3]
                                    diff = abs(
                                        beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
                                    if diff < min_diff:
                                        min_diff = diff
                                        min_size = y[2]
                            beam_size = min_size
                    # if floor=='10F' and beam_name =='B1-4':
                    #     print(check_list)
                    if beam_size != '':
                        if (floor, beam_name, '', beam_rotate) in dic_plan:
                            set_plan.remove(
                                (floor, beam_name, '', beam_rotate))
                            dic_plan.pop((floor, beam_name, '', beam_rotate))
                            error(
                                f'read_plan error in step 12: {floor} {beam_name} duplicate. ')
                            warning_list.append(
                                f'{floor} {beam_name} duplicate. ')
                        set_plan.add(
                            (floor, beam_name, beam_size, beam_rotate))
                        dic_plan[(floor, beam_name, beam_size,
                                  beam_rotate)] = full_coor
                        check_list.append((floor, beam_name))
                    else:
                        if (floor, beam_name) not in check_list:
                            set_plan.add((floor, beam_name, '', beam_rotate))
                            dic_plan[(floor, beam_name, '',
                                      beam_rotate)] = full_coor
                            error(
                                f'read_plan error in step 12: {floor} {beam_name} cannot find size. ')
                            warning_list.append(
                                f'{floor} {beam_name} cannot find size. ')

                else:  # 不用對尺寸
                    set_plan.add((floor, beam_name))
                    dic_plan[(floor, beam_name)] = full_coor

    # doc_plan.Close(SaveChanges=False)
    progress('平面圖讀取進度 12/13', progress_file)

    # Step 13. 用 dic_plan((floor, beam_name, beam_size, beam_rotate) -> full_coor) 和 beam_direction_mid_scale_set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale) 找圖是否畫錯
    # 還要順便把結果寫入plan_new_file, big_file, sml_file，我懶得再把參數傳出來了哈哈
    if mline_scaling:
        # Step 13-1. 開啟應用程式
        flag = 0
        while not flag and error_count <= 10:
            try:
                wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 13-1, {ex}, error_count = {error_count}')

        # Step 13-2. 匯入檔案
        flag = 0
        while not flag and error_count <= 10:
            try:
                doc_plan = wincad_plan.Documents.Open(plan_filename)
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 13-2, {ex}, error_count = {error_count}')

        # Step 13-3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag and error_count <= 10:
            try:
                msp_plan = doc_plan.Modelspace
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 13-3, {ex}, error_count = {error_count}')
        time.sleep(5)

        # Step 13-4. 設定mark的圖層
        flag = 0
        while not flag and error_count <= 10:
            try:
                layer_plan = doc_plan.Layers.Add(f"S-CLOUD_{date}")
                doc_plan.ActiveLayer = layer_plan
                layer_plan.color = 10
                layer_plan.Linetype = "Continuous"
                layer_plan.Lineweight = 0.5
                flag = 1
            except Exception as ex:
                error_count += 1
                time.sleep(5)
                error(
                    f'read_plan error in step 13-4, {ex}, error_count = {error_count}')

        # Step 13-5. 找最近的複線，有錯要畫圖 -> 中點找中點
        error_list = []
        for x, item in dic_plan.items():
            if 'x' in x[2]:
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
                min_diff = inf
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
                if len(temp_list) != 0:
                    closet_mline = min(temp_list, key=lambda m: abs(
                        midpoint[0] - m[2][0]) + abs(midpoint[1]-m[2][1]))
                    min_scale = closet_mline[3]
                    min_coor = closet_mline[2]
                    # for y in beam_direction_mid_scale_set:
                    #     if y[0] == beam_layer and y[1] == 0:
                    #         coor = y[2]
                    #         diff = abs(midpoint[0] - coor[0]) + abs(midpoint[1] - coor[1])
                    #         if diff < min_diff:
                    #             min_diff = diff
                    #             min_scale = y[3]
                    #             min_coor = coor

                # if beam_rotate != 0: # 直的 or 歪的
                #     for y in beam_direction_mid_scale_set:
                #         if y[0] == beam_layer and y[1] == 1:
                #             coor = y[2]
                #             diff = abs(midpoint[0] - coor[0]) + abs(midpoint[1] - coor[1])
                #             if diff < min_diff:
                #                 min_diff = diff
                #                 min_scale = y[3]
                #                 min_coor = coor

                # 全部連線
                # coor_list = [min_coor[0], min_coor[1], 0, midpoint[0], midpoint[1], 0]
                # points = vtFloat(coor_list)
                # line = msp_plan.AddPolyline(points)
                # line.SetWidth(0, 3, 3)
                # line.color = 200

                if min_scale == '' or min_scale != beam_scale:
                    error_list.append(
                        (x[0], x[1], f'寬度有誤：文字為{beam_scale}，圖上為{min_scale}。\n'))
                    coor = dic_plan[x]
                    # 畫框框
                    coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                                 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
                    points = vtFloat(coor_list)
                    pointobj = msp_plan.AddPolyline(points)
                    pointobj.color = 130
                    for i in range(4):
                        pointobj.SetWidth(i, 10, 10)
                    # 只畫有錯的線
                    if min_coor != '':
                        coor_list = [min_coor[0], min_coor[1],
                                     0, midpoint[0], midpoint[1], 0]
                        points = vtFloat(coor_list)
                        line = msp_plan.AddPolyline(points)
                        line.SetWidth(0, 3, 3)
                        line.color = 200

        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)

        # Step 13.6 寫入txt_filename
        f_fbeam = open(fbeam_file, "w", encoding='utf-8')
        f_big = open(big_file, "w", encoding='utf-8')
        f_sml = open(sml_file, "w", encoding='utf-8')
        error_list.sort(key=lambda x: turn_floor_to_float(x[0]))
        f_big.write('核對mline寬度結果\n')
        f_sml.write('核對mline寬度結果\n')
        f_fbeam.write('核對mline寬度結果\n')
        for x in error_list:
            if x[1][0] == "B" or x[1][0] == "G" or x[1][0] == "C":
                f_big.write(f"('{x[0]}', '{x[1]}'): {x[2]}")
            elif x[1][0] == "F":
                f_fbeam.write(f"('{x[0]}', '{x[1]}'): {x[2]}")
            else:
                f_sml.write(f"('{x[0]}', '{x[1]}'): {x[2]}")

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

    progress('平面圖讀取進度 13/13', progress_file)
    progress('平面圖讀取完畢。', progress_file)
    # plan.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in plan: \n")
    l = list(set_plan)
    l.sort()
    for x in l:
        f.write(f'{x}\n')
    f.close()

    return (set_plan, dic_plan, warning_list)


def read_beam(beam_filename, text_layer, progress_file):
    error_count = 0
    progress('開始讀取梁配筋圖', progress_file)
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 1: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 1/9', progress_file)

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
                f'read_beam error in step 2: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 2/9', progress_file)

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 3: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 3/9', progress_file)

    # Step 4 解鎖所有圖層 -> 不然不能刪東西
    flag = 0
    while not flag and error_count <= 10:
        try:
            layer_count = doc_beam.Layers.count

            for x in range(layer_count):
                layer = doc_beam.Layers.Item(x)
                layer.Lock = False
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 4: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 4/9', progress_file)

    # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         count = 0
    #         total = msp_beam.Count
    #         progress(
    #             f'正在炸梁配筋圖的圖塊及篩選判斷用的物件，梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
    #         for object in msp_beam:
    #             count += 1
    #             explode_fail = 0
    #             while explode_fail <= 3:
    #                 try:
    #                     if object.EntityName == "AcDbBlockReference" and object.Layer == text_layer:
    #                         object.Explode()
    #                     if object.Layer != text_layer:
    #                         object.Delete()
    #                     if count % 1000 == 0:
    #                         progress(
    #                             f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
    #                     break
    #                 except:
    #                     explode_fail += 1
    #                     try:
    #                         msp_beam = doc_beam.Modelspace
    #                     except:
    #                         pass
    #         flag = 1

    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(
    #             f'read_beam error in step 5: {e}, error_count = {error_count}.')
    #         try:
    #             msp_beam = doc_beam.Modelspace
    #         except:
    #             pass
    # progress('梁配筋圖讀取進度 5/9', progress_file)

    # Step 6. 重新匯入modelspace
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         msp_beam = doc_beam.Modelspace
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(
    #             f'read_beam error in step 6: {e}, error_count = {error_count}.')
    # progress('梁配筋圖讀取進度 6/9', progress_file)

    # Step 7. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    progress('正在遍歷梁配筋圖上的物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候', progress_file)
    floor_to_beam_set = set()
    flag = 0
    count = 0
    used_layer_list = text_layer
    # for key, layer_name in layer_config.items():
    #     used_layer_list += layer_name
    total = msp_beam.Count
    progress(
        f'梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)

    for msp_object in msp_beam:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
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

                if object_layer in text_layer and object.ObjectName == "AcDbText" and ' ' in object.TextString:
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
    progress('梁配筋圖讀取進度 7/9', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False
    doc_beam.Close(SaveChanges=False)
    progress('梁配筋圖讀取進度 9/9', progress_file)
    progress('梁配筋圖讀取完成。', progress_file)
    return floor_to_beam_set


def sort_beam(floor_to_beam_set: set, result_filename: str, progress_file: str, sizing: bool):
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
    progress('梁配筋圖讀取進度 8/9', progress_file)

    # Step 9. 完成set_beam和dic_beam
    dic_beam = {}
    set_beam = set()
    for x in floor_to_beam_set:
        floor = x[0]
        beam = x[1]
        coor = x[2]
        size = x[3]
        floor_list = []
        to_bool = False
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
                    error(f'read_beam error in step 9: The error above is from here.')
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

    # beam.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in beam: \n")
    l = list(set_beam)
    l.sort()
    for x in l:
        f.write(f'{x}\n')
    f.close()

    return (set_beam, dic_beam)


# 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
def write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, drawing, progress_file, sizing, mline_scaling, fbeam_file):
    error_count = 0
    progress("開始標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'。", progress_file)
    pythoncom.CoInitialize()
    set_in_plan = set_plan - set_beam
    list_in_plan = list(set_in_plan)
    list_in_plan.sort()
    set_in_beam = set_beam - set_plan
    list_in_beam = list(set_in_beam)
    list_in_beam = [beam for beam in list_in_beam if beam[2] != 'replicate']
    list_in_beam.sort()

    f_fbeam = open(fbeam_file, "a", encoding='utf-8')
    f_big = open(big_file, "a", encoding='utf-8')
    f_sml = open(sml_file, "a", encoding='utf-8')

    f_fbeam.write("in plan but not in beam: \n")
    f_big.write("in plan but not in beam: \n")
    f_sml.write("in plan but not in beam: \n")

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
        progress('平面圖標註進度 1/5', progress_file)

        # Step 2. 匯入檔案
        flag = 0
        while not flag and error_count <= 10:
            try:
                if mline_scaling:
                    doc_plan = wincad_plan.Documents.Open(plan_new_filename)
                else:
                    doc_plan = wincad_plan.Documents.Open(plan_filename)
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_plan error in step 2, {e}, error_count = {error_count}.')
        progress('平面圖標註進度 2/5', progress_file)

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
        progress('平面圖標註進度 3/5', progress_file)

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
        progress('平面圖標註進度 4/5', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_plan.Close(SaveChanges=False)
        except:
            pass
        return False

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
        # if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G' :
        #     wrong_data = 0
        #     if sizing:
        #         error_beam = [beam_beam for beam_beam in list_in_beam if plan_beam[0] == beam_beam[0] and plan_beam[1] == beam_beam[1]]
        #         if error_beam:
        #             pass
        #         else:
        #             pass
        #         for beam_beam in list_in_beam:
        #             if plan_beam[0] == beam_beam[0] and plan_beam[1] == beam_beam[1] and plan_beam[2] != beam_beam[2]:
        #                 if plan_beam[2] != '':
        #                     err_list_big.append((plan_beam, 0, beam_beam[2])) # type(tuple of floor and wrong beam, err_message, correct) 0是尺寸錯誤
        #                     drawing = 1
        #                 else:
        #                     err_list_big_size.append(f'{(plan_beam[0], plan_beam[1])}\n')
        #                     drawing = 0
        #                 wrong_data = 1
        #                 break
        #     if not wrong_data:
        #         err_list_big.append((plan_beam, 1)) # type(tuple of floor and wrong beam, err_message) 1是找不到梁
        #     big_error += 1
        # elif x[1][0] == 'F':
        #     wrong_data = 0
        #     if sizing:
        #         pass
        # else:
        #     wrong_data = 0
        #     if sizing:
        #         for y in list2: # 去另一邊找有沒有floor跟beam相同但尺寸不同的東西
        #             if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
        #                 if x[2] != '':
        #                     err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
        #                     drawing = 1doc_plan
        #                 else:
        #                     err_list_sml_size.append(f'{(x[0], x[1])}\n')
        #                     drawing = 0
        #                 wrong_data = 1
        #                 break
        #     if not wrong_data:
        #         err_list_sml.append((x, 1)) # type(tuple of floor and wrong beam, err_message)
        #         # f_sml.write(f'{x}: 找不到這根梁\n')
        #     sml_error += 1

        if drawing and beam_drawing:
            coor = dic_plan[plan_beam]
            coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                         20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
            points = vtFloat(coor_list)
            pointobj = msp_plan.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)
    if drawing:
        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)
    return error_list, f_fbeam, f_big, f_sml


def output_error_list(error_list: list, f_big: TextIOWrapper, f_sml: TextIOWrapper, f_fbeam: TextIOWrapper, title_text='XS-BEAM', set_item=set, progress_file='./result/tmp'):
    beam_error_size_list = []
    beam_no_beam_list = []
    sbeam_error_size_list = []
    sbeam_no_beam_list = []
    fbeam_error_size_list = []
    fbeam_no_beam_list = []

    # error_list = sorted(error_list,key = cmp_to_key(mycmp))
    error_size = [e for e in error_list if e[1] == 'error_size']
    no_size = [e for e in error_list if e[1] == 'no_size']
    no_beam = [e for e in error_list if e[1] == 'no_beam']
    replicate_beam = [e for e in error_list if e[1] == 'replicate']

    error_size = sorted(error_size, key=cmp_to_key(mycmp))
    no_size = sorted(no_size, key=cmp_to_key(mycmp))
    no_beam = sorted(no_beam, key=cmp_to_key(mycmp))
    replicate_beam = sorted(replicate_beam, key=cmp_to_key(mycmp))

    beam_list = [b for b in set_item if b[1][0] ==
                 'B' or b[1][0] == 'C' or b[1][0] == 'G']
    fbeam_list = [b for b in set_item if b[1][0] == 'F']
    sbeam_list = [
        b for b in set_item if b not in beam_list and b not in fbeam_list]
    error_rate = 0
    sb_error_rate = 0
    fb_error_rate = 0
    for f in [f_big, f_sml, f_fbeam]:
        f.write(f'========================\n')
        f.write(f'尺寸錯誤:\n')

    for e in error_size:
        beam = e[0]
        beam_name = beam[1]
        if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G':
            beam_error_size_list.append(e)
            f_big.write(f'{beam}: 尺寸有誤，在{title_text}那邊是{e[2]}\n')
        elif beam_name[0] == 'F':
            fbeam_error_size_list.append(e)
            f_fbeam.write(f'{beam}: 尺寸有誤，在{title_text}那邊是{e[2]}\n')
        else:
            sbeam_error_size_list.append(e)
            f_sml.write(f'{beam}: 尺寸有誤，在{title_text}那邊是{e[2]}\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write(f'========================\n')
        f.write(f'{title_text}缺少:\n')
    for e in no_beam:
        beam = e[0]
        beam_name = e[0][1]
        if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G':
            beam_no_beam_list.append(e)
            f_big.write(f'{beam}: 找不到這根梁\n')
        elif beam_name[0] == 'F':
            fbeam_no_beam_list.append(e)
            f_fbeam.write(f'{beam}: 找不到這根梁\n')
        else:
            sbeam_no_beam_list.append(e)
            f_sml.write(f'{beam}: 找不到這根梁\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('========================\n')

    if beam_list:
        error_rate = round((len(beam_error_size_list) +
                           len(beam_no_beam_list)) / len(beam_list) * 100, 2)
        f_big.write(f'error rate = {error_rate} %\n')
        f_big.write(
            f'error rate = ({len(beam_error_size_list)}+{len(beam_no_beam_list)})/{len(beam_list)}={error_rate} %\n')
    else:
        f_big.write('平面圖中無大梁(B、G、C開頭)\n')

    if sbeam_list:
        sb_error_rate = round(
            (len(sbeam_error_size_list) + len(sbeam_no_beam_list)) / len(sbeam_list) * 100, 2)
        f_sml.write(
            f'error rate = ({len(sbeam_error_size_list)}+{len(sbeam_no_beam_list)})/{len(sbeam_list)}={sb_error_rate} %\n')
    else:
        f_sml.write('平面圖中無小梁\n')

    if fbeam_list:
        fb_error_rate = round(
            (len(fbeam_error_size_list) + len(fbeam_no_beam_list)) / len(fbeam_list) * 100, 2)
        f_fbeam.write(
            f'error rate = ({len(fbeam_error_size_list)}+{len(fbeam_no_beam_list)})/{len(fbeam_list)}={fb_error_rate} %\n')
    else:
        f_fbeam.write('平面圖中無地梁(F開頭)\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('========================\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('備註: (平面圖找不到尺寸)\n')

    for e in no_size:
        beam = e[0]
        beam_name = e[0][1]
        if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G':
            f_big.write(f'{beam}: 找不到尺寸\n')
        elif beam_name[0] == 'F':
            f_fbeam.write(f'{beam}: 找不到尺寸\n')
        else:
            f_sml.write(f'{beam}: 找不到尺寸\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('========================\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('備註: (重複配筋)\n')

    for e in replicate_beam:
        beam = e[0]
        beam_name = e[0][1]
        if beam_name[0] == 'B' or beam_name[0] == 'C' or beam_name[0] == 'G':
            f_big.write(f'{beam}: 重複配筋\n')
        elif beam_name[0] == 'F':
            f_fbeam.write(f'{beam}: 重複配筋\n')
        else:
            f_sml.write(f'{beam}: 重複配筋\n')

    for f in [f_big, f_sml, f_fbeam]:
        f.write('========================\n')

    # error_size_beam = [e for e in error_list if e[0][1][0] == 'error_size']
    # error_size_sbeam = [e for e in error_list if e[1] == 'error_size']
    # error_size_fbeam = [e for e in error_list if e[1] == 'error_size']
    # err_list_big = sorted(err_list_big, key = cmp_to_key(mycmp))
    # err_list_sml = sorted(err_list_sml, key = cmp_to_key(mycmp))

    # for y in err_list_big:
    #     if y[1] == 0:
    #         f_big.write(f'{y[0]}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
    #     else:
    #         f_big.write(f'{y[0]}: 找不到這根梁\n')

    # for y in err_list_sml:
    #     if y[1] == 0:
    #         f_sml.write(f'{y[0]}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
    #     else:
    #         f_sml.write(f'{y[0]}: 找不到這根梁\n')

    # 算分母
    # fb_count = 0
    # big_count = 0
    # sml_count = 0
    # for x in set_plan:
    #     if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
    #         big_count += 1
    #     elif x[1][0] == 'F':
    #         fb_count += 1
    #     else:
    #         sml_count += 1

    # # 計算錯誤率可能會噴錯，因為分母為0
    # try:
    #     big_rate = round(big_error / big_count * 100, 2)
    #     f_big.write(f'error rate = {big_rate} %\n')

    # except:
    #     big_rate = 'unfinish'
    #     error(f'write_plan error in step 5, there are no big beam in plan.txt?')

    # if len(err_list_big_size):
    #     f_big.write(f'備註: (平面圖找不到尺寸)\n')
    #     for y in err_list_big_size:
    #         f_big.write(y)
    # f_big.write(f'\n')

    # try:
    #     sml_rate = round(sml_error / sml_count * 100, 2)
    #     f_sml.write(f'error rate = {sml_rate} %\n')

    # except:
    #     sml_rate = 'unfinish'
    #     error(f'write_plan error in step 5, there are no small beam in plan.txt?')

    # if len(err_list_sml_size):
    #     f_sml.write(f'備註: (平面圖找不到尺寸)\n')
    #     for y in err_list_sml_size:
    #         f_sml.write(y)
    # f_sml.write(f'\n')

    f_big.close()
    f_sml.close()
    f_fbeam.close()
    progress('平面圖標註進度 5/5', progress_file)
    progress("標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'完成。", progress_file)
    return (error_rate, sb_error_rate, fb_error_rate)


# 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
def write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, drawing, progress_file, sizing, fbeam_file):
    error_count = 0
    progress("開始標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'。", progress_file)
    pythoncom.CoInitialize()
    set1 = set_plan - set_beam
    list_in_plan = list(set1)
    list_in_plan.sort()
    set2 = set_beam - set_plan
    list_in_beam = list(set2)
    list_in_beam.sort()
    error_list = []
    f_fbeam = open(fbeam_file, "a", encoding='utf-8')
    f_big = open(big_file, "a", encoding='utf-8')
    f_sml = open(sml_file, "a", encoding='utf-8')

    f_fbeam.write("in beam but not in plan: \n")
    f_big.write("in beam but not in plan: \n")
    f_sml.write("in beam but not in plan: \n")

    if drawing:
        # Step 1. 開啟應用程式
        flag = 0
        while not flag and error_count <= 10:
            try:
                wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(
                    f'write_beam error in step 1, {e}, error_count = {error_count}.')
        progress('梁配筋圖標註進度 1/5', progress_file)
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
        progress('梁配筋圖標註進度 2/5', progress_file)
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
        progress('梁配筋圖標註進度 3/5', progress_file)
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
    progress('梁配筋圖標註進度 4/5', progress_file)

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
                error_list.append((beam_beam, 'error_size', error_beam[0][2]))
                beam_drawing = 1
            else:
                error_list.append((beam_beam, 'no_size', error_beam[0][2]))
        else:
            beam_drawing = 1
            error_list.append((beam_beam, 'no_beam', ''))
    # big_error = 0
    # sml_error = 0
    # err_list_big = []
    # err_list_sml = []
    # print('hi')
    # for x in list2:
    #     if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
    #         wrong_data = 0
    #         if sizing:
    #             for y in list1:
    #                 if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
    #                     if y[2] != '':
    #                         err_list_big.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
    #                     wrong_data = 1
    #                     break
    #         if not wrong_data:
    #             err_list_big.append((x, 1)) # type(tuple of floor and wrong beam, err_message)
    #         big_error += 1
    #     else:
    #         wrong_data = 0
    #         if sizing:
    #             for y in list1:
    #                 if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
    #                     if y[2] != '':
    #                         err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)

    #                     wrong_data = 1
    #                     break
    #         if not wrong_data:
    #             err_list_sml.append((x, 1)) # type(tuple of floor and wrong beam, err_message)
    #             # f_sml.write(f'{x}: 找不到這根梁\n')
    #         sml_error += 1

        if drawing and beam_drawing:
            coor = dic_beam[beam_beam]
            coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] +
                         20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
            points = vtFloat(coor_list)
            pointobj = msp_beam.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)
    # print('hihi')
    # err_list_big = sorted(err_list_big, key = cmp_to_key(mycmp))
    # err_list_sml = sorted(err_list_sml, key = cmp_to_key(mycmp))
    # print('hihihi')
    # for y in err_list_big:
    #     if y[1] == 0:
    #         f_big.write(f'{y[0]}: 尺寸有誤，在XS-PLAN那邊是{y[2]}\n')
    #     else:
    #         f_big.write(f'{y[0]}: 找不到這根梁\n')
    # for y in err_list_sml:
    #     if y[1] == 0:
    #         f_sml.write(f'{y[0]}: 尺寸有誤，在XS-PLAN那邊是{y[2]}\n')
    #     else:
    #         f_sml.write(f'{y[0]}: 找不到這根梁\n')

    if drawing:
        doc_beam.SaveAs(beam_new_filename)
        doc_beam.Close(SaveChanges=True)

    # big_count = 0
    # sml_count = 0
    # for x in set_beam:
    #     if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
    #         big_count += 1
    #     else:
    #         sml_count += 1

    # # 計算錯誤率可能會噴錯，因為分母為0
    # try:
    #     big_rate = round(big_error / big_count * 100, 2)
    #     f_big.write(f'error rate = {big_rate} %\n')
    # except:
    #     big_rate = 'unfinish'
    #     error(f'write_beam error in step 5, there are no big beam in beam.txt?')

    # try:
    #     sml_rate = round(sml_error / sml_count * 100, 2)
    #     f_sml.write(f'error rate = {sml_rate} %\n')
    # except:
    #     sml_rate = 'unfinish'
    #     error(f'write_beam error in step 5, there are no small beam in beam.txt?')
    # progress('梁配筋圖標註進度 5/5', progress_file)
    # f_big.close()
    # f_sml.close()
    # progress("標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'完成。", progress_file)
    return (error_list, f_fbeam, f_big, f_sml)


def write_result_log(excel_file, task_name, plan_result, beam_result, date, runtime, other):
    sheet_name = 'result_log_new'
    if not plan_result:
        plan_result = ['', '', '']
    if not beam_result:
        beam_result = ['', '', '']
    plan_not_beam_big = plan_result[0]
    plan_not_beam_sml = plan_result[1]
    plan_not_beam_fb = plan_result[2]
    beam_not_plan_big = beam_result[0]
    beam_not_plan_sml = beam_result[1]
    beam_not_plan_fb = beam_result[2]
    new_list = [(task_name,
                 plan_not_beam_big,
                 plan_not_beam_sml,
                 plan_not_beam_fb,
                 beam_not_plan_big,
                 beam_not_plan_sml,
                 beam_not_plan_fb,
                 date, runtime, other)]
    dfNew = pd.DataFrame(new_list, columns=[
                         '名稱', '平面圖大梁錯誤率', '平面圖小梁錯誤率', '平面圖地梁錯誤率', '配筋圖大梁錯誤率', '配筋圖小梁錯誤率', '配筋圖地梁錯誤率', '執行時間', '執行日期', '備註'])
    if os.path.exists(excel_file):
        writer = pd.ExcelWriter(
            excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace')
        df = pd.read_excel(excel_file)
        df = pd.concat([df, dfNew], axis=0, ignore_index=True, join='inner')
    else:
        writer = pd.ExcelWriter(excel_file, engine='openpyxl')
        df = dfNew
    df.to_excel(writer, sheet_name)
    writer.save()
    return


def run_plan(plan_filename, plan_new_filename, big_file, sml_file, layer_config: dict, result_filename, progress_file, sizing, mline_scaling, date, fbeam_file):
    start_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    if False:
        plan_data = read_plan(plan_filename=plan_filename,
                              layer_config=layer_config,
                              progress_file=progress_file,
                              sizing=sizing,
                              mline_scaling=mline_scaling)
        save_temp_file.save_pkl(
            data=plan_data, tmp_file=f'{os.path.splitext(plan_filename)[0]}_plan_set.pkl')
    else:
        plan_data = save_temp_file.read_temp(
            tmp_file=r'TEST\2024-0412\2024-04-12-16-332024-0412 茂德新莊-XS-PLAN_plan_set.pkl')
    set_plan, dic_plan, warning_list = sort_plan(plan_filename=plan_filename,
                                                 plan_new_filename=plan_new_filename,
                                                 plan_data=plan_data,
                                                 layer_config=layer_config,
                                                 sizing=sizing,
                                                 mline_scaling=mline_scaling,
                                                 big_file=big_file,
                                                 sml_file=sml_file,
                                                 fbeam_file=fbeam_file,
                                                 result_filename=result_filename,
                                                 progress_file=progress_file,
                                                 date=date)
    output_txt = f'{os.path.splitext(plan_new_filename)[0]}_result.txt'
    end_date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    output_progress_report(output_filename=output_txt,
                           layer_config=layer_config,
                           start_date=start_date,
                           end_date=end_date,
                           project_name=plan_filename,
                           warning_list=warning_list,
                           plan_filename=plan_filename,
                           plan_data=plan_data)
    return (set_plan, dic_plan)


def run_beam(beam_filename, text_layer, result_filename, progress_file, sizing):
    if True:
        floor_to_beam_set = read_beam(
            beam_filename=beam_filename, text_layer=text_layer, progress_file=progress_file)
        save_temp_file.save_pkl(data=floor_to_beam_set,
                                tmp_file=f'{os.path.splitext(beam_filename)[0]}_beam_set.pkl')
    else:
        floor_to_beam_set = save_temp_file.read_temp(
            r'TEST\2023-1114\2023-11-14-14-47台電竹園-XS-BEAM_beam_set.pkl')
    set_beam, dic_beam = sort_beam(floor_to_beam_set=floor_to_beam_set,
                                   result_filename=result_filename,
                                   progress_file=progress_file,
                                   sizing=sizing)
    return (set_beam, dic_beam)


def output_progress_report(output_filename: str, layer_config: dict, start_date, end_date, plan_data: dict, project_name: str, plan_filename: str, warning_list: list):
    delimiter = '\n'
    cad_data = {
        '圖框': len(plan_data['block_coor_list']),
        '平面圖梁編號': len(plan_data['coor_to_beam_set']),
        '表格梁編號': len(plan_data['coor_to_size_beam']),
        '表格梁尺寸': len(plan_data['coor_to_size_string'])
    }
    with open(output_filename, 'w') as f:
        f.write(
            f'專案名稱:{project_name}\n平面圖名稱:{plan_filename}\n開始時間:{start_date} \n結束時間:{end_date} \n')
        f.write(f'圖層參數:{layer_config} \n')
        f.write(f'CAD資料:{cad_data}]\n')
        f.write(f'平面圖樓層:{plan_data["coor_to_floor_set"]}]\n')
        f.write(f'==========================\n')
        f.write('錯誤訊息:\n')
        f.write(f'{delimiter.join(warning_list)}')


error_file = './result/error_log.txt'  # error_log.txt的路徑

if __name__ == '__main__':
    start = time.time()
    # plan_data = save_temp_file.read_temp('plan_to_beam_0307-2.pkl')
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # beam_filename = r"D:/Desktop/BeamQC/TEST/INPUT\2023-03-03-15-45temp-temp.dwg"#sys.argv[1] # XS-BEAM的路徑
    # beam_filenames = [r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0224 UPBBR.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0224_UPBBL.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0224_UPsbR_v2.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0302_UPsbL_v2.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\1F 小梁.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\1F大樑.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0303 FB.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\B1大樑.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\B1小梁.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0303\2023-0303 小地梁.dwg"]
    beam_filenames = [
        r'D:\Desktop\BeamQC\TEST\2023-1016\2023-10-16-10-11中德三重-XS-BEAM.dwg']
    # sys.argv[2] # XS-PLAN的路徑
    plan_filenames = [
        r'D:\Desktop\BeamQC\TEST\2023-1016\1017-B1F.dwg']
    # sys.argv[3] # XS-BEAM_new的路徑
    beam_new_filename = r"D:\Desktop\BeamQC\TEST\2023-1016\1016-XS-BEAM_new.dwg"
    # sys.argv[4] # XS-PLAN_new的路徑
    plan_new_filename = r"D:\Desktop\BeamQC\TEST\2023-1016\1016-XS_PLAN_new.dwg"
    # sys.argv[5] # 大梁結果
    big_file = r"TEST\2023-1114\1114-big.txt"
    # sys.argv[6] # 小梁結果
    sml_file = r"TEST\2023-1114\1114-sml.txt"
    # sys.argv[6] # 地梁結果
    fbeam_file = r"TEST\2023-1114\1114-fb.txt"
    # 在beam裡面自訂圖層
    text_layer = ['S-RC']  # sys.argv[7]

    # 在plan裡面自訂圖層
    block_layer = ['0', 'DwFm', 'DEFPOINTS']  # sys.argv[8] # 框框的圖層
    floor_layer = ['S-TITLE']  # sys.argv[9] # 樓層字串的圖層
    size_layer = ['S-TEXT']  # sys.argv[12] # 梁尺寸字串圖層
    line_layer = ['S-TABLE']  # 梁表格圖層
    # beam_layer = ['S-RCBMG', 'S-RCBMG(FB)', 'S-RCBMB', 'S-RCBMB(FB)']  # 大樑複線圖層
    # beam_text_layer = ['S-TEXTG', 'S-TEXTB']  # 大樑文字圖層
    sml_beam_layer = ['S-RCBMB', 'S-RCBMB(FB)']  # 小梁複線圖層
    sml_beam_text_layer = ['S-TEXTB']  # 小梁文字圖層
    big_beam_layer = ['S-RCBMG', 'S-RCBMG(FB)']
    big_beam_text_layer = ['S-TEXTG']  # 小梁文字圖層
    task_name = '1017-B1F-temp'  # sys.argv[13]

    progress_file = './result/tmp'  # sys.argv[14]

    sizing = 1  # 要不要對尺寸
    mline_scaling = 1  # 要不要對複線寬度

    plan_file = './result/plan.txt'  # plan.txt的路徑
    beam_file = './result/beam.txt'  # beam.txt的路徑
    excel_file = './result/result_log.xlsx'  # result_log.xlsx的路徑

    date = time.strftime("%Y-%m-%d", time.localtime())
    layer_config = {
        # 'line_layer':line_layer,
        'text_layer': text_layer,
        'block_layer': block_layer,
        'floor_layer': floor_layer,
        # 'beam_layer': beam_layer,
        # 'beam_text_layer': beam_text_layer,
        'big_beam_text_layer': big_beam_text_layer,
        'sml_beam_text_layer': sml_beam_text_layer,
        'big_beam_layer': big_beam_layer,
        'sml_beam_layer': sml_beam_layer,
        'size_layer': size_layer,
        'line_layer': line_layer,
        # 'col_layer':col_layer
    }
    # 多檔案用','來連接，不用空格。Ex. 'file1,file2,file3'
    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()

    res_plan = []
    res_beam = []
    set_plan = set()
    dic_plan = {}
    set_beam = set()
    dic_beam = {}

    for plan_filename in plan_filenames:
        # run_plan(plan_filename=plan_filename,
        #          plan_new_filename=plan_new_filename,
        #          big_file=big_file,
        #          sml_file=sml_file,
        #          layer_config=layer_config,
        #          result_filename=plan_file,
        #          progress_file= progress_file,
        #          sizing= sizing,
        #          mline_scaling= mline_scaling,
        #          date = date,
        #          fbeam_file= fbeam_file)
        #     # res_plan.append(pool.apply_async(read_plan, (plan_filename, plan_new_filename, big_file, sml_file, floor_layer, big_beam_layer, big_beam_text_layer, sml_beam_layer, sml_beam_text_layer, block_layer, size_layer, plan_file, progress_file, sizing, mline_scaling, date,fbeam_file)))
        res_plan.append(pool.apply_async(run_plan, (plan_filename, plan_new_filename, big_file,
                        sml_file, layer_config, plan_file, progress_file, sizing, mline_scaling, date, fbeam_file)))
    for beam_filename in beam_filenames:
        # res_beam.append(pool.apply_async(read_beam, (beam_filename, text_layer, beam_file, progress_file, sizing)))
        res_beam.append(pool.apply_async(run_beam, (beam_filename,
                        layer_config['text_layer'], beam_file, progress_file, sizing)))
    # plan_filename = plan_filenames[0]
    # beam_filename = beam_filenames[0]
    # run_plan(plan_filenames[0], plan_new_filename, big_file, sml_file,layer_config , plan_file, progress_file, sizing, mline_scaling, date,fbeam_file)
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
            write_result_log(excel_file, task_name, '', '', '', '',
                             f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')

    for beam in res_beam:
        beam = beam.get()
        if beam:
            set_beam = set_beam | beam[0]
            if beam_drawing:
                dic_beam = beam[1]
        else:
            end = time.time()
            write_result_log(excel_file, task_name, '', '', '', '',
                             f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'failed')
    plan_error_list, f_fbeam, f_big, f_sml = write_plan(plan_filename, plan_new_filename, set_plan, set_beam,
                                                        dic_plan, big_file, sml_file, date, plan_drawing, progress_file, sizing, mline_scaling, fbeam_file=fbeam_file)
    plan_result = output_error_list(error_list=plan_error_list, f_fbeam=f_fbeam,
                                    f_big=f_big, f_sml=f_sml, title_text='XS-BEAM', set_item=set_plan)
    beam_error_list, f_fbeam, f_big, f_sml = write_beam(
        beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, beam_drawing, progress_file, sizing, fbeam_file=fbeam_file)
    beam_result = output_error_list(error_list=beam_error_list, f_sml=f_sml,
                                    f_big=f_big, f_fbeam=f_fbeam, title_text='XS-PLAN', set_item=set_beam)
    end = time.time()
    print(end - start)
    write_result_log(excel_file=excel_file, task_name=task_name, plan_result=plan_result, beam_result=beam_result,
                     runtime=f'{round(end - start, 2)}(s)', date=time.strftime("%Y-%m-%d %H:%M", time.localtime()), other='none')
