from gzip import READ
from math import inf, sqrt
from multiprocessing.spawn import prepare
from tabnanny import check
from tkinter import HIDDEN
from numpy import object_
from openpyxl import load_workbook
import win32com.client
import pythoncom
import re
import time
import multiprocessing
import os
import pandas as pd
import sys
from functools import cmp_to_key
from math import inf
import traceback

def turn_floor_to_float(floor): # 把字串變成小數 (因為1MF = 1.5, 所以不能用整數)

    if ' ' in floor: # 不小心有空格要把空格拔掉
        floor = floor.replace(' ', '')

    if floor == 'FB': # FB 直接變-1000層
        floor = str(-1000)

    if floor == 'PRF' or floor == 'PR': # PRF 直接變2000層
        floor = str(2000)
    
    if floor == 'RF': # RF = R1F
        floor = str(1001)

    if 'F' in floor: # 有F要把F拔掉
        floor = floor.replace("F", "")

    try: # 轉int可能會失敗
        if 'B' in floor: # 有B直接變負整數
            floor = str(-int(floor.replace("B", "")))
        
        if 'R' in floor: # 有R直接+1000
            floor = str(int(floor.replace("R", "")) + 1000)
        
        if 'M' in floor: # 半層以0.5表示
            floor = str(int(floor.replace("M", "")) + 0.5)
    except:
        pass

    try:
        floor = float(floor)
        return floor
    except:
        error(f'turn_floor_to_float error: {floor} cannot be turned to float.')
        return False

def turn_floor_to_string(floor): # 把數字變回字串
    if floor == -1000:
        floor = 'FBF' # 因為beam的部分字尾非F會自動補F，所以在diff的時候要一致

    elif floor > -1000 and floor < 0:
        floor = f'B{int(-floor)}F'

    elif floor > 0 and floor < 1000:
        if floor * 2 % 2 == 0: # 整數*2之後會是2的倍數
            floor = f'{int(floor)}F'
        else: # 如果有.5的話，*2之後會是奇數
            floor = f'{int(floor - 0.5)}MF'

    elif floor > 1000 and floor < 2000: 
        floor = f'R{int(floor - 1000)}F'

    elif floor == 2000:
        floor = 'PRF'

    else:
        error(f'turn_floor_to_string error: {floor} cannot be turned to string.')
        return False

    return floor

def turn_floor_to_list(floor, Bmax, Fmax, Rmax):
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
            except:
                error(f'turn_floor_to_list error: {floor} cannot be turned to list.')
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
                error(f'turn_floor_to_list error: {floor} cannot be turned to list.')
                
    return floor_list

def floor_exist(i, Bmax, Fmax, Rmax): # 判斷是否為空號，例如B2F-PRF會從-2跑到2000，但顯然區間裡面的值不可能都合法
    if i == -1000 or i == 2000: 
        return True
    
    elif i >= Bmax and i < 0: 
        return True
    
    elif i > 0 and i <= Fmax: 
        return True
    
    elif i > 1000 and i <= Rmax: 
        return True

    return False

def vtFloat(l): #要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)

def error(error_message): # 把錯誤訊息印到error.log裡面
    f = open(error_file, 'a', encoding = 'utf-8')
    localtime = time.asctime(time.localtime(time.time()))
    f.write(f'{localtime} | {error_message}\n')
    f.close
    return

def progress(message, progress_file): # 把進度印到progress裡面，在app.py會對這個檔案做事
    f = open(progress_file, 'a', encoding = 'utf-8')
    f.write(f'{message}\n')
    f.close
    return

# 可以先看完 write_plan 跟 write_beam 整理 txt 的規則再看這個函式在幹嘛
# 自定義排序規則

def mycmp(a, b): # a, b 皆為 tuple , 可能是 ((floor, beam), 0, correct) 或 ((floor, beam), 1) 
    if a[1] == b[1]: # err_message 一樣，比樓層
        if turn_floor_to_float(a[0][0]) >  turn_floor_to_float(b[0][0]):
            return 1
        elif turn_floor_to_float(a[0][0]) ==  turn_floor_to_float(b[0][0]):
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

weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡B']
beam_head1 = ['B', 'b', 'G', 'g']
beam_head2 = ['FB','Fb','CB', 'CG', 'cb']

def read_plan(plan_filename, plan_new_filename, big_file, sml_file, floor_layer, big_beam_layer, big_beam_text_layer, sml_beam_layer, sml_beam_text_layer, block_layer, size_layer, result_filename, progress_file, sizing, mline_scaling, date):
    def _cal_ratio(pt1,pt2):
        if abs(pt1[1]-pt2[1]) == 0:
            return 1000
        return abs(pt1[0]-pt2[0])/abs(pt1[1]-pt2[1])
    error_count = 0
    progress('開始讀取平面圖(核對項目: 梁配筋對應)', progress_file)
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 1: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 1/13', progress_file)

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 2: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 2/13', progress_file)

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 3: {e}, error_count = {error_count}.')
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
            error(f'read_plan error in step 4: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 4/13', progress_file)

    # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間
    flag = 0
    time.sleep(5)
    layer_list = [floor_layer, size_layer, big_beam_layer, big_beam_text_layer, sml_beam_layer, sml_beam_text_layer]
    non_trash_list = layer_list + [block_layer]
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_plan.Count
            progress(f'正在炸平面圖的圖塊及篩選判斷用的物件，平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
            for object in msp_plan:
                count += 1
                if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:
                    object.Explode()
                if object.Layer not in non_trash_list:
                    object.Delete()
                if count % 1000 == 0:
                    progress(f'平面圖已讀取{count}/{total}個物件', progress_file) # 每1000個跳一次，確認有在動
            flag = 1

        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 5: {e}, error_count = {error_count}.')
            try:
                msp_plan = doc_plan.Modelspace
            except:
                pass

    progress('平面圖讀取進度 5/13', progress_file)

    # Step 6. 重新匯入modelspace，剛剛炸出來的東西要重新讀一次
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 6: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 6/13', progress_file)
    
    # Step 7. 遍歷所有物件 -> 完成各種我們要的set跟list
    
    progress('正在遍歷平面圖上的物件並篩選出有效信息，運行時間取決於平面圖大小，請耐心等候', progress_file)
    flag = 0
    coor_to_floor_set = set() # set (字串的coor, floor)，Ex. 求'1F'這個字串的座標在哪
    coor_to_beam_set = set() # set (coor, (beam, size))，Ex. 求'B1-6'這個字串的座標在哪，如果後面有括號的話，順便紀錄尺寸，否則size = ''
    block_coor_list = [] # 存取方框最左下角的點座標

    # for sizing
    coor_to_size_beam = set() # set (coor, size_beam)，Ex. 紀錄表格中'Bn'這個字串的座標
    coor_to_size_string = set() # set (coor, size_string)，Ex. 紀錄表格中'25x50'這個字串的座標

    # for mline_scaling
    beam_direction_mid_scale_set = set() # set (beam_layer(big_beam_layer or sml_beam_layer), direction(0: 橫的, 1: 直的), midpoint, scale)

    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_plan.Count
            progress(f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
            for object in msp_plan:
                count += 1
                if count % 1000 == 0 or count == total:
                    progress(f'平面圖已讀取{count}/{total}個物件', progress_file)

                # 取floor的字串 -> 抓括號內的字串 (Ex. '十層至十四層結構平面圖(10F~14F)' -> '10F~14F')
                # 若此處報錯，可能原因: 1. 沒有括號, 2. 有其他括號在鬧(ex. )
                if object.Layer == floor_layer and object.ObjectName == "AcDbText" and '(' in object.TextString and object.InsertionPoint[1] >= 0:
                    floor = object.TextString
                    floor = re.search('\(([^)]+)', floor).group(1) #取括號內的樓層數
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2)) #不取概數的話後面抓座標會出問題，例如兩個樓層在同一格
                    no_chinese = False
                    for ch in floor: # 待修正
                        if ch == 'B' or ch == 'F' or ch == 'R' or ch.isdigit():
                            no_chinese = True
                            break
                    if floor != '' and no_chinese:
                        coor_to_floor_set.add((coor, floor))
                    else:
                        error(f'read_plan error in step 7: floor is an empty string or it is Chinese. ')

                # 取beam的字串
                # 此處會錯的地方在於可能會有沒遇過的怪怪comma，但報應不會在這裡產生，會直接反映到結果

                elif object.Layer in [big_beam_text_layer, sml_beam_text_layer] and (object.ObjectName == "AcDbText" or object.ObjectName == "AcDbMLeader")\
                        and object.TextString != '' and (object.TextString[0] in beam_head1 or object.TextString[0:2] in beam_head2):

                    beam = object.TextString
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    size = ''
                    if '(' in beam:
                        size = (((beam.split('(')[1]).split(')')[0]).replace(' ', '')).replace('X', 'x')
                        if 'x' not in size:
                            size = ''
                        else:
                            try:
                                first = size.split('x')[0]
                                second = size.split('x')[1]
                                if not (float(first) and float(second)):
                                    size = ''
                            except:
                                size = ''
                        beam = beam.split('(')[0] # 取括號前內容即可
                    comma_char = ','
                    for char in weird_comma_list:
                        if char in beam:
                            comma_char = char
                    comma = beam.count(comma_char)
                    for i in range(comma + 1):
                        try:
                            coor_to_beam_set.add(((coor1, coor2), (beam.split(comma_char)[i], size, round(object.Rotation, 2))))
                        except: # 只要不是0or1.57，後面核對的時候就會橫的值得都找。
                            coor_to_beam_set.add(((coor1, coor2), (beam.split(comma_char)[i], size, 1)))
                            error(f'read_plan error in step 7: {(beam, size)} at {(coor1, coor2)} cannot find Rotation.')

                # 為了排版好看的怪產物，目前看到的格式為'{\W0.7;B4-2\P(80x100)}'，所以使用分號及反斜線來切
                # 切爛了也不會報錯，直接反映在結果
                elif object.Layer in [big_beam_text_layer, sml_beam_text_layer] and object.ObjectName == "AcDbMText":
                    beam = object.TextString
                    semicolon = beam.count(';')
                    size = ''
                    for i in range(semicolon + 1):
                        s = beam.split(';')[i]
                        if s[0] in beam_head1 or s[0:2] in beam_head2:
                            if '(' in s:
                                size = (((s.split('(')[1]).split(')')[0]).replace(' ', '')).replace('X', 'x')
                                if 'x' not in size:
                                    size = ''
                                else:
                                    try:
                                        first = size.split('x')[0]
                                        second = size.split('x')[1]
                                        if not (float(first) and float(second)):
                                            size = ''
                                    except:
                                        size = ''
                                s = s.split('(')[0]
                            if '\\' in s:
                                s = s.split('\\')[0]
                            beam = s
                            break
                    
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))

                    if '(' in beam:
                        beam = beam.split('(')[0] # 取括號前內容即可
                    if ' ' in beam:
                        beam = beam.replace(' ', '') # 有空格要把空格拔掉
                    if beam[0] in beam_head1 or beam[0:2] in beam_head2:
                        try:
                            coor_to_beam_set.add(((coor1, coor2), (beam, size, round(object.Rotation, 2))))
                        except:
                            error(f'read_plan error in step 7: {(beam, size)} at {(coor1, coor2)} cannot find Rotation.')

                # 找框框，完成block_coor_list，格式為((0.0, 0.0), (14275.54, 10824.61))
                # 此處不會報錯

                elif object.Layer == block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    if _cal_ratio(coor1,coor2) >= 1/5 and _cal_ratio(coor1,coor2) <= 5: #避免雜訊影響框框
                        block_coor_list.append((coor1, coor2))
                
                # 找size

                if sizing or mline_scaling:
                    if object.Layer == size_layer and object.EntityName == "AcDbText" and object.GetBoundingBox()[0][1] >= 0:
                        coor = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        if 'FGn' in object.TextString:
                            coor_to_size_beam.add((coor, 'FG'))
                        if 'FBn' in object.TextString:
                            coor_to_size_beam.add((coor, 'FB'))
                        if 'FWB' in object.TextString:
                            coor_to_size_beam.add((coor, 'FWB'))
                        if 'Fbn' in object.TextString:
                            coor_to_size_beam.add((coor, 'Fbn'))
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
                                        coor_to_size_beam.add((coor, beam.split(')')[1]))
                                        coor_to_size_beam.add((coor, f"c{beam.split(')')[1]}"))
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
                            string = (object.TextString.replace(' ', '')).replace('X', 'x')
                            try:
                                first = string.split('x')[0]
                                second = string.split('x')[1]
                                if float(first) and float(second):
                                    coor_to_size_string.add((coor, string))
                            except:
                                pass
                    
                # 找複線
                if mline_scaling:
                    if object.Layer in [big_beam_layer, sml_beam_layer] and object.ObjectName == "AcDbMline":
                        start = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        end = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        x_diff = abs(start[0] - end[0])
                        y_diff = abs(start[1] - end[1])
                        mid = ((start[0] + end[0]) / 2, (start[1] + end[1]) / 2)
                        if x_diff + y_diff > 100: # 防超短的東東
                            if x_diff < y_diff: # 算直的, 1
                                beam_direction_mid_scale_set.add((object.Layer, 1, mid, abs(object.MLineScale)))
                            else: # 算橫的, 0
                                beam_direction_mid_scale_set.add((object.Layer, 0, mid, abs(object.MLineScale)))
            
            flag = 1

        except Exception as e:
            error_class = e.__class__.__name__ #取得錯誤類型
            detail = e.args[0] #取得詳細內容
            cl, exc, tb = sys.exc_info() #取得Call Stack
            lastCallStack = traceback.extract_tb(tb)[-1] #取得Call Stack的最後一筆資料
            fileName = lastCallStack[0] #取得發生的檔案名稱
            lineNum = lastCallStack[1] #取得發生的行號
            funcName = lastCallStack[2] #取得發生的函數名稱
            errMsg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, error_class, detail)
            error_count += 1
            time.sleep(5)
            error(f'read_plan error in step 7: {e}, error_count = {error_count}.')

    progress('平面圖讀取進度 7/13', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_plan.Close(SaveChanges=False)
        except:
            pass
        return False
    
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
    for x in coor_to_floor_set: # set (字串的coor, floor)
        string_coor = x[0]
        floor = x[1]
        for block_coor in block_coor_list:
            x_diff_left = string_coor[0] - block_coor[0][0] # 和左下角的diff
            y_diff_left = string_coor[1] - block_coor[0][1]
            x_diff_right = string_coor[0] - block_coor[1][0] # 和右上角的diff
            y_diff_right = string_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0: # 要在框框裡面才算
                floor_to_coor_set.add((floor, block_coor))
    progress('平面圖讀取進度 9/13', progress_file)

    # Step 10. 算出Bmax, Fmax, Rmax, 用途: 跑for迴圈的時候，知道哪些是空號
    # 此處可能報錯的地方在於turn_floor_to_float，但函式本身return false時就會報錯，所以此處不另外再報錯
    Bmax = 0 # 地下最深到幾層(不包括FB不包括FB)
    Fmax = 0 # 正常樓最高到幾層
    Rmax = 0 # R開頭最高到幾層(不包括PRF)
    for y in floor_to_coor_set:
        floor = y[0]
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
                tmp_floor_list.append(turn_floor_to_float(floor.split(comma_char)[i]))

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
        for x in size_coor_set: # set(size_beam, min_size, coor)
            size_coor = x[2]
            size_string = x[1]
            size_beam = x[0]
            min_floor = []
            for z in floor_to_coor_set: # list (floor_list, block左下角和右上角的coor)
                floor_list = z[0]
                block_coor = z[1] 
                x_diff_left = size_coor[0] - block_coor[0][0] # 和左下角的diff
                y_diff_left = size_coor[1] - block_coor[0][1]
                x_diff_right = size_coor[0] - block_coor[1][0] # 和右上角的diff
                y_diff_right = size_coor[1] - block_coor[1][1]
                if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:                    
                    if len(min_floor) == 0 or min_floor[0] != floor_list:
                        min_floor.append(floor_list)

            if len(min_floor) != 0:
                for i in range(len(min_floor)):
                    floor_list = min_floor[i]
                    for floor in floor_list:
                        floor_beam_size_coor_set.add((floor, size_beam, size_string, size_coor))                            
            else:
                error(f'read_plan error in step 11: {(size_beam, size_string, size_coor)} cannot find min_floor.')
    progress('平面圖讀取進度 11/13', progress_file)

    # Step 12. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補

    set_plan = set() # set元素為 (樓層, 梁柱名稱, size)
    dic_plan = {} # 透過(floor, beam, size)去找字串座標

    # 如果沒有要對size -> set元素為 (floor, beam)
    # 如果有要對size但沒有要對mline -> set元素為 (floor, beam, size)
    # 如果有要對size和mline -> set元素為 (floor, beam, size, rotate)
    check_list = []
    # 遍歷所有beam，找這是幾樓的
    for x in coor_to_beam_set: # set(coor, (beam, size))
        beam_coor = x[0][0] # 取左下角即可
        full_coor = x[0] # 左下跟右上都有
        beam_name = x[1][0]
        beam_size = x[1][1]
        beam_rotate = x[1][2]
        min_floor = ''
        
        for z in floor_to_coor_set: # 我其實是list歐哈哈 (floor_list, block左下角和右上角的coor)
            floor_list = z[0]
            block_coor = z[1]
            x_diff_left = beam_coor[0] - block_coor[0][0] # 和左下角的diff
            y_diff_left = beam_coor[1] - block_coor[0][1]
            x_diff_right = beam_coor[0] - block_coor[1][0] # 和右上角的diff
            y_diff_right = beam_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:             
                if min_floor == '' or min_floor == floor_list:
                    min_floor = floor_list

                else: # 有很多層在同一個block, 仍然透過字串的coor找樓層 -> 應從已知選項找最適合的，而不是全部重找，這樣會找到框框外面的東西
                    for y in coor_to_floor_set: # set(字串的coor, floor)
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
                            if y[0] == floor and y[1] == beam_name: # 先看有沒有像g1, g2完全吻合的
                                tmp_coor = y[3]
                                diff = abs(beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
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
                                if y[0] == floor and y[1] == header: # 再看頭一不一樣(去掉數字的頭)
                                    tmp_coor = y[3]
                                    diff = abs(beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
                                    if diff < min_diff:
                                        min_diff = diff
                                        min_size = y[2]
                        if min_size != '':
                            beam_size = min_size
                        else:
                            for y in floor_beam_size_coor_set:
                                if y[0] == floor and y[1] in header: # 再看頭有沒有包含到(ex. GA-1 算在 G)
                                    tmp_coor = y[3]
                                    diff = abs(beam_coor[0] - tmp_coor[0]) + abs(beam_coor[1] - tmp_coor[1])
                                    if diff < min_diff:
                                        min_diff = diff
                                        min_size = y[2]
                            beam_size = min_size
                    # if floor=='10F' and beam_name =='B1-4':
                    #     print(check_list)
                    if beam_size != '':
                        if (floor, beam_name, '', beam_rotate) in dic_plan:
                            set_plan.remove((floor, beam_name, '', beam_rotate))
                            dic_plan.pop((floor, beam_name, '', beam_rotate))
                            error(f'read_plan error in step 12: {floor} {beam_name} duplicate. ')
                        set_plan.add((floor, beam_name, beam_size, beam_rotate))
                        dic_plan[(floor, beam_name, beam_size, beam_rotate)] = full_coor
                        check_list.append((floor, beam_name))
                    else:
                        if not (floor, beam_name) in check_list:
                            set_plan.add((floor, beam_name, '', beam_rotate))
                            dic_plan[(floor, beam_name, '', beam_rotate)] = full_coor
                            error(f'read_plan error in step 12: {floor} {beam_name} cannot find size. ')

                else: # 不用對尺寸
                    set_plan.add((floor, beam_name))
                    dic_plan[(floor, beam_name)] = full_coor

    doc_plan.Close(SaveChanges=False)
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
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(f'read_plan error in step 13-1, {e}, error_count = {error_count}')

        # Step 13-2. 匯入檔案
        flag = 0
        while not flag and error_count <= 10:
            try:
                doc_plan = wincad_plan.Documents.Open(plan_filename)
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(f'read_plan error in step 13-2, {e}, error_count = {error_count}')

        # Step 13-3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag and error_count <= 10:
            try:
                msp_plan = doc_plan.Modelspace
                flag = 1
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(f'read_plan error in step 13-3, {e}, error_count = {error_count}')
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
            except Exception as e:
                error_count += 1
                time.sleep(5)
                error(f'read_plan error in step 13-4, {e}, error_count = {error_count}')

        # Step 13-5. 找最近的複線，有錯要畫圖 -> 中點找中點
        error_list = []
        for x in dic_plan:
            if 'x' in x[2]:
                beam_scale = float(x[2].split('x')[0])
                beam_coor = dic_plan[x]
                beam_layer = sml_beam_layer
                if x[1][0].isupper():
                    beam_layer = big_beam_layer
                if 'Fb' in x[1]:
                    beam_layer = sml_beam_layer
                beam_rotate = x[3]
                midpoint = ((beam_coor[0][0] + beam_coor[1][0]) / 2, (beam_coor[0][1] + beam_coor[1][1]) / 2)
                min_diff = inf
                min_scale = ''
                min_coor = ''
                if beam_rotate != 1.57: # 橫的 or 歪的，90度 = pi / 2 = 1.57 (前面有取round到後二位)
                    for y in beam_direction_mid_scale_set:
                        if y[0] == beam_layer and y[1] == 0:
                            coor = y[2]
                            diff = abs(midpoint[0] - coor[0]) + abs(midpoint[1] - coor[1])
                            if diff < min_diff:
                                min_diff = diff
                                min_scale = y[3]
                                min_coor = coor

                if beam_rotate != 0: # 直的 or 歪的
                    for y in beam_direction_mid_scale_set:
                        if y[0] == beam_layer and y[1] == 1:
                            coor = y[2]
                            diff = abs(midpoint[0] - coor[0]) + abs(midpoint[1] - coor[1])
                            if diff < min_diff:
                                min_diff = diff
                                min_scale = y[3]
                                min_coor = coor

                # 全部連線
                # coor_list = [min_coor[0], min_coor[1], 0, midpoint[0], midpoint[1], 0]
                # points = vtFloat(coor_list)
                # line = msp_plan.AddPolyline(points)
                # line.SetWidth(0, 3, 3)
                # line.color = 200
                
                if min_scale == '' or min_scale != beam_scale:
                    error_list.append((x[0], x[1], f'寬度有誤：文字為{beam_scale}，圖上為{min_scale}。\n'))
                    coor = dic_plan[x]
                    # 畫框框
                    coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
                    points = vtFloat(coor_list)
                    pointobj = msp_plan.AddPolyline(points)
                    pointobj.color = 130
                    for i in range(4):
                        pointobj.SetWidth(i, 10, 10)
                    # 只畫有錯的線
                    if min_coor != '':
                        coor_list = [min_coor[0], min_coor[1], 0, midpoint[0], midpoint[1], 0]
                        points = vtFloat(coor_list)
                        line = msp_plan.AddPolyline(points)
                        line.SetWidth(0, 3, 3)
                        line.color = 200

        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)

        # Step 13.6 寫入txt_filename

        f_big = open(big_file, "w", encoding = 'utf-8')
        f_sml = open(sml_file, "w", encoding = 'utf-8')
        error_list.sort(key = lambda x: turn_floor_to_float(x[0]))
        f_big.write('核對mline寬度結果\n')
        f_sml.write('核對mline寬度結果\n')
        for x in error_list: 
            if x[1][0].isupper():
                f_big.write(f"('{x[0]}', '{x[1]}'): {x[2]}")
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

    return (set_plan, dic_plan)

def read_beam(beam_filename, text_layer, result_filename, progress_file, sizing):
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
            error(f'read_beam error in step 1: {e}, error_count = {error_count}.')
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
            error(f'read_beam error in step 2: {e}, error_count = {error_count}.')
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
            error(f'read_beam error in step 3: {e}, error_count = {error_count}.')
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
            error(f'read_beam error in step 4: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 4/9', progress_file)

    # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間  
    flag = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_beam.Count
            progress(f'正在炸梁配筋圖的圖塊及篩選判斷用的物件，梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
            for object in msp_beam:
                count += 1
                if object.EntityName == "AcDbBlockReference" and object.Layer == text_layer:
                    object.Explode()
                if object.Layer != text_layer:
                    object.Delete()
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
            flag = 1
        
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 5: {e}, error_count = {error_count}.')
            try:
                msp_beam = doc_beam.Modelspace
            except:
                pass
    progress('梁配筋圖讀取進度 5/9', progress_file)

    # Step 6. 重新匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 6: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 6/9', progress_file)
    
    # Step 7. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    progress('正在遍歷梁配筋圖上的物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候', progress_file)
    floor_to_beam_set = set()
    flag = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_beam.Count
            progress(f'梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
            for object in msp_beam:
                count += 1
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
                if object.Layer == text_layer and object.ObjectName == "AcDbText" and ' ' in object.TextString:
                    pre_beam = (object.TextString.split(' ')[1]).split('(')[0] # 把括號以後的東西拔掉
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
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
                            size = (((object.TextString.split('(')[1]).split(')')[0]).replace(' ', '')).replace('X', 'x') # size 的格式就是 90x50, 沒空格且使用小寫x作為乘號
                            floor_to_beam_set.add((floor, beam, (coor1, coor2), size))
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 7: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 7/9', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False

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
                tmp_floor_list.append(turn_floor_to_float(floor.split(comma_char)[i]))

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
                except:
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
                set_beam.add((floor, beam, size))
                dic_beam[(floor, beam, size)] = coor
            else:
                set_beam.add((floor, beam))
                dic_beam[(floor, beam)] = coor

    doc_beam.Close(SaveChanges=False)
    progress('梁配筋圖讀取進度 9/9', progress_file)
    progress('梁配筋圖讀取完成。', progress_file)

    # beam.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in beam: \n")
    l = list(set_beam)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()
    
    return (set_beam, dic_beam)

def write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, drawing, progress_file, sizing, mline_scaling): # 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
    error_count = 0
    progress("開始標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'。", progress_file)
    pythoncom.CoInitialize()
    set1 = set_plan - set_beam
    list1 = list(set1)
    list1.sort()
    set2 = set_beam - set_plan
    list2 = list(set2)
    list2.sort()

    f_big = open(big_file, "a", encoding = 'utf-8')
    f_sml = open(sml_file, "a", encoding = 'utf-8')

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
                error(f'write_plan error in step 1, {e}, error_count = {error_count}.')
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
                error(f'write_plan error in step 2, {e}, error_count = {error_count}.')
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
                error(f'write_plan error in step 3, {e}, error_count = {error_count}.')
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
                error(f'write_plan error in step 4, {e}, error_count = {error_count}.')
        progress('平面圖標註進度 4/5', progress_file)
    
    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_plan.Close(SaveChanges=False)
        except:
            pass
        return False
    
    # Step 5. 完成in plan but not in beam，畫圖，以及計算錯誤率
    big_error = 0
    sml_error = 0
    err_list_big = []
    err_list_sml = []
    err_list_big_size = []
    err_list_sml_size = []
    for x in list1: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G' or x[1][0] == 'F':
            wrong_data = 0
            if sizing:
                for y in list2:
                    if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                        if x[2] != '':
                            err_list_big.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct) 0是尺寸錯誤
                            drawing = 1
                        else:
                            err_list_big_size.append(f'{(x[0], x[1])}\n')
                            drawing = 0
                        wrong_data = 1
                        break
            if not wrong_data:
                err_list_big.append((x, 1)) # type(tuple of floor and wrong beam, err_message) 1是找不到梁            
            big_error += 1

        else:
            wrong_data = 0
            if sizing:
                for y in list2: # 去另一邊找有沒有floor跟beam相同但尺寸不同的東西
                    if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                        if x[2] != '':
                            err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                            drawing = 1
                        else:
                            err_list_sml_size.append(f'{(x[0], x[1])}\n')
                            drawing = 0
                        wrong_data = 1
                        break
            if not wrong_data:   
                err_list_sml.append((x, 1)) # type(tuple of floor and wrong beam, err_message)   
                # f_sml.write(f'{x}: 找不到這根梁\n')
            sml_error += 1
        
        if drawing:
            coor = dic_plan[x]
            coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
            points = vtFloat(coor_list)
            pointobj = msp_plan.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)

    err_list_big = sorted(err_list_big, key = cmp_to_key(mycmp))
    err_list_sml = sorted(err_list_sml, key = cmp_to_key(mycmp))

    for y in err_list_big:
        if y[1] == 0:
            f_big.write(f'{y[0]}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
        else:
            f_big.write(f'{y[0]}: 找不到這根梁\n')
    
    for y in err_list_sml:
        if y[1] == 0:
            f_sml.write(f'{y[0]}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
        else:
            f_sml.write(f'{y[0]}: 找不到這根梁\n')

    if drawing:
        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)

    # 算分母
    big_count = 0
    sml_count = 0
    for x in set_plan:
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':        
            big_count += 1
        else:
            sml_count += 1
    
    # 計算錯誤率可能會噴錯，因為分母為0
    try:
        big_rate = round(big_error / big_count * 100, 2)
        f_big.write(f'error rate = {big_rate} %\n')
        
    except:
        big_rate = 'unfinish'
        error(f'write_plan error in step 5, there are no big beam in plan.txt?')
    
    if len(err_list_big_size):
        f_big.write(f'備註: (平面圖找不到尺寸)\n')
        for y in err_list_big_size:
            f_big.write(y)
    f_big.write(f'\n')

    try:
        sml_rate = round(sml_error / sml_count * 100, 2)
        f_sml.write(f'error rate = {sml_rate} %\n')
        
    except:
        sml_rate = 'unfinish'
        error(f'write_plan error in step 5, there are no small beam in plan.txt?')

    if len(err_list_sml_size):
        f_sml.write(f'備註: (平面圖找不到尺寸)\n')
        for y in err_list_sml_size:
            f_sml.write(y)
    f_sml.write(f'\n')

    f_big.close()
    f_sml.close()
    progress('平面圖標註進度 5/5', progress_file)
    progress("標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'完成。", progress_file)
    return (big_rate, sml_rate)
    

def write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, drawing, progress_file, sizing): # 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
    error_count = 0
    progress("開始標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'。", progress_file)
    pythoncom.CoInitialize()
    set1 = set_plan - set_beam
    list1 = list(set1)
    list1.sort()
    set2 = set_beam - set_plan
    list2 = list(set2)
    list2.sort()

    f_big = open(big_file, "a", encoding = 'utf-8')
    f_sml = open(sml_file, "a", encoding = 'utf-8')

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
                error(f'write_beam error in step 1, {e}, error_count = {error_count}.')
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
                error(f'write_beam error in step 2, {e}, error_count = {error_count}.')
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
                error(f'write_beam error in step 3, {e}, error_count = {error_count}.')
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
                error(f'write_beam error in step 4, {e}, error_count = {error_count}.')
    progress('梁配筋圖標註進度 4/5', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False

    # Step 5. 完成in beam but not in plan，畫圖，以及計算錯誤率
    big_error = 0
    sml_error = 0
    err_list_big = []
    err_list_sml = []
    print('hi')
    for x in list2: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
            wrong_data = 0
            if sizing:
                for y in list1:
                    if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                        if y[2] != '':
                            err_list_big.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                        wrong_data = 1
                        break
            if not wrong_data:
                err_list_big.append((x, 1)) # type(tuple of floor and wrong beam, err_message)   
            big_error += 1
        else:
            wrong_data = 0
            if sizing:
                for y in list1:
                    if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                        if y[2] != '':
                            err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)

                        wrong_data = 1
                        break
            if not wrong_data:   
                err_list_sml.append((x, 1)) # type(tuple of floor and wrong beam, err_message)   
                # f_sml.write(f'{x}: 找不到這根梁\n')
            sml_error += 1

        if drawing:
            coor = dic_beam[x]
            coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
            points = vtFloat(coor_list)
            pointobj = msp_beam.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)
    print('hihi')
    err_list_big = sorted(err_list_big, key = cmp_to_key(mycmp))
    err_list_sml = sorted(err_list_sml, key = cmp_to_key(mycmp))
    print('hihihi')
    for y in err_list_big:
        if y[1] == 0:
            f_big.write(f'{y[0]}: 尺寸有誤，在XS-PLAN那邊是{y[2]}\n')
        else:
            f_big.write(f'{y[0]}: 找不到這根梁\n')
    for y in err_list_sml:
        if y[1] == 0:
            f_sml.write(f'{y[0]}: 尺寸有誤，在XS-PLAN那邊是{y[2]}\n')
        else:
            f_sml.write(f'{y[0]}: 找不到這根梁\n')
        
    if drawing:
        doc_beam.SaveAs(beam_new_filename)
        doc_beam.Close(SaveChanges=True)

    big_count = 0
    sml_count = 0
    for x in set_beam:
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':        
            big_count += 1
        else:
            sml_count += 1
    
    # 計算錯誤率可能會噴錯，因為分母為0
    try:
        big_rate = round(big_error / big_count * 100, 2)
        f_big.write(f'error rate = {big_rate} %\n')
    except:
        big_rate = 'unfinish'
        error(f'write_beam error in step 5, there are no big beam in beam.txt?')
    
    try:
        sml_rate = round(sml_error / sml_count * 100, 2)
        f_sml.write(f'error rate = {sml_rate} %\n')
    except:
        sml_rate = 'unfinish'
        error(f'write_beam error in step 5, there are no small beam in beam.txt?')
    progress('梁配筋圖標註進度 5/5', progress_file)
    f_big.close()
    f_sml.close()
    progress("標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'完成。", progress_file)
    return (big_rate, sml_rate)

def write_result_log(excel_file, task_name, plan_not_beam_big, plan_not_beam_sml, beam_not_plan_big, beam_not_plan_sml, date, runtime, other):
    sheet_name = 'result_log'
    new_list = [(task_name, plan_not_beam_big, plan_not_beam_sml, beam_not_plan_big, beam_not_plan_sml, date, runtime, other)]
    dfNew=pd.DataFrame(new_list, columns = ['名稱' , 'in plan not in beam 大梁', 'in plan not in beam 小梁','in beam not in plan 大梁', 'in plan not In beam 小梁', '執行時間', '執行日期' , '備註'])
    if os.path.exists(excel_file):
        writer = pd.ExcelWriter(excel_file,engine='openpyxl',mode='a', if_sheet_exists='replace')
        df = pd.read_excel(excel_file)  
        df = pd.concat([df, dfNew], axis=0, ignore_index = True, join = 'inner')
    else:
        writer = pd.ExcelWriter(excel_file, engine='openpyxl') 
        df = dfNew
    df.to_excel(writer,sheet_name)
    writer.save()    
    return

error_file = './result/error_log.txt' # error_log.txt的路徑

if __name__=='__main__':
    start = time.time()
    
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    beam_filename = r"D:\Desktop\BeamQC\TEST\2023-0118\小檜溪D_S4-01~S6-22_梁配筋詳圖_1110825.dwg"#sys.argv[1] # XS-BEAM的路徑
    plan_filename = r"D:\Desktop\BeamQC\TEST\2023-0118\小檜溪D_S2-01~28_結構平面圖_1110825.dwg"#sys.argv[2] # XS-PLAN的路徑
    beam_new_filename = r"D:\Desktop\BeamQC\TEST\XS-BEAM_new.dwg"#sys.argv[3] # XS-BEAM_new的路徑
    plan_new_filename = r"D:\Desktop\BeamQC\TEST\XS-PLAN_new.dwg"#sys.argv[4] # XS-PLAN_new的路徑
    big_file = r"D:\Desktop\BeamQC\TEST\big.txt"#sys.argv[5] # 大梁結果
    sml_file = r"D:\Desktop\BeamQC\TEST\sml.txt"#sys.argv[6] # 小梁結果

    # 在beam裡面自訂圖層
    text_layer = 'S-RC'#sys.argv[7]

    # 在plan裡面自訂圖層
    block_layer = '0'#sys.argv[8] # 框框的圖層
    floor_layer = 'S-TITLE'#sys.argv[9] # 樓層字串的圖層
    size_layer = 'S-TEXT'#sys.argv[12] # 梁尺寸字串圖層
    big_beam_layer = 'S-RCBMG'#大樑複線圖層
    big_beam_text_layer = 'S-TEXTG'#大樑文字圖層
    sml_beam_layer = 'S-RCBMB'#小梁複線圖層
    sml_beam_text_layer = 'S-TEXTB'#小梁文字圖層
    task_name = 'temp'#sys.argv[13]

    progress_file = './result/tmp'#sys.argv[14]

    sizing = 1 # 要不要對尺寸
    mline_scaling = 1 # 要不要對複線寬度

    plan_file = './result/plan.txt' # plan.txt的路徑
    beam_file = './result/beam.txt' # beam.txt的路徑
    excel_file = './result/result_log.xlsx' # result_log.xlsx的路徑
    
    date = time.strftime("%Y-%m-%d", time.localtime())

    # 多檔案用','來連接，不用空格。Ex. 'file1,file2,file3'
    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()

    plan_file_count = plan_filename.count(',') + 1
    beam_file_count = beam_filename.count(',') + 1

    res_plan = [None] * plan_file_count
    res_beam = [None] * beam_file_count
    set_plan = set()
    dic_plan = {}
    set_beam = set()
    dic_beam = {}

    for i in range(plan_file_count):
        res_plan[i] = pool.apply_async(read_plan, (plan_filename.split(',')[i], plan_new_filename, big_file, sml_file, floor_layer, big_beam_layer, big_beam_text_layer, sml_beam_layer, sml_beam_text_layer, block_layer, size_layer, plan_file, progress_file, sizing, mline_scaling, date))

    for i in range(beam_file_count):
        res_beam[i] = pool.apply_async(read_beam, (beam_filename.split(',')[i], text_layer, beam_file, progress_file, sizing))
    
    for i in range(plan_file_count):
        final_plan = res_plan[i].get()
        set_plan = set_plan | final_plan[0]
        if plan_file_count == 1 and beam_file_count == 1:
            dic_plan = final_plan[1]

    for i in range(beam_file_count):
        final_beam = res_beam[i].get()
        set_beam = set_beam | final_beam[0]
        if plan_file_count == 1 and beam_file_count == 1:
            dic_beam = final_beam[1]

    drawing = 0
    if plan_file_count == 1 and beam_file_count == 1:
        drawing = 1
    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, drawing, progress_file, sizing, mline_scaling)
    beam_result = write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, drawing, progress_file, sizing)

    end = time.time()
    write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    # write_result_log(excel_file,'','','','','','',time.strftime("%Y-%m-%d %H:%M", time.localtime()),'')