from gzip import READ
from multiprocessing.spawn import prepare
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

def turn_floor_to_float(floor): # turn string to float
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

    if 'B' in floor: # 有B直接變負整數
        floor = str(-int(floor.replace("B", "")))
    
    if 'R' in floor: # 有R直接+1000
        floor = str(int(floor.replace("R", "")) + 1000)
    
    if 'M' in floor: # 半層以0.5表示
        floor = str(int(floor.replace("M", "")) + 0.5)
    try:
        floor = float(floor)
        return floor
    except:
        error(f'turn_floor_to_float error: {floor} cannot be turned to float.')
        return False

def turn_floor_to_string(floor): # turn float to string
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
    f = open(error_file, 'a')
    localtime = time.asctime(time.localtime(time.time()))
    f.write(f'{localtime} | {error_message}\n')
    f.close
    return

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
beam_head2 = ['CB', 'CG', 'cb']

def read_plan(plan_filename, floor_layer, beam_layer, block_layer, size_layer, result_filename, explode):
    print('開始讀取平面圖(核對項目: 梁配筋對應)')
    # Step 1. 打開應用程式
    flag = 0
    while not flag:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 1: {e}.')
    print('平面圖讀取進度 1/11')
    # Step 2. 匯入檔案
    flag = 0
    while not flag:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 2: {e}.')
    print('平面圖讀取進度 2/11')
    # Step 3. 匯入modelspace
    flag = 0
    while not flag:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 3: {e}.')
    print('平面圖讀取進度 3/11')
    if explode: # 需要提前炸圖塊再進來
        print('XS-PLAN 正在炸圖塊，運行時間取決於平面圖大小，請耐心等候')
        # Step 4. 遍歷所有物件 -> 炸圖塊
        # 炸圖塊看性質即可，不用看圖層      
        flag = 0
        layer_list = [floor_layer, size_layer] + beam_layer
        while not flag:
            try:
                count = 0
                total = msp_plan.Count
                print(f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
                for object in msp_plan:
                    count += 1
                    if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:
                        object.Explode()
                    if count % 1000 == 0:
                        print(f'平面圖已讀取{count}/{total}個物件')
                flag = 1
                print('平面圖讀取進度 4/11')
            except Exception as e:
                time.sleep(5)
                error(f'read_plan error in step 4: {e}.')

        # Step 5. 重新匯入modelspace
        flag = 0
        while not flag:
            try:
                msp_plan = doc_plan.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'read_plan error in step 5: {e}.')
        print('平面圖讀取進度 5/11')
    
    # Step 6. 遍歷所有物件 -> 完成 coor_to_floor_set, coor_to_beam_set, block_coor_list
    coor_to_floor_set = set() # set (字串的coor, floor)
    coor_to_beam_set = set() # set (coor, (beam, size))
    coor_to_size_beam = set() # set (coor, size_beam)
    coor_to_size_string = set() # set (coor, size_string)
    block_coor_list = [] # 存取方框最左下角的點座標
    print('正在遍歷平面圖上所有物件並篩選出有效信息，運行時間取決於平面圖大小，請耐心等候')
    flag = 0
    while not flag:
        try:
            count = 0
            total = msp_plan.Count
            print(f'平面圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
            for object in msp_plan:
                count += 1
                if count % 1000 == 0:
                    print(f'平面圖已讀取{count}/{total}個物件')
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
                        error(f'read_plan error in step 6: floor is an empty string or it is Chinese. ')
                # 取beam的字串 -> 只取括號前的東西 (Ex. 'GC-3(50x95)' -> 'GC-3')
                # 此處會錯的地方在於可能會有沒遇過的怪怪comma，但報應不會在這裡產生，會直接反映到結果
                if object.Layer in beam_layer and (object.ObjectName == "AcDbText" or object.ObjectName == "AcDbMLeader") and object.GetBoundingBox()[0][1] >= 0 \
                        and (object.TextString[0] in beam_head1 or object.TextString[0:2] in beam_head2):
                    beam = object.TextString
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    size = ''
                    if '(' in beam:
                        size = (((beam.split('(')[1]).split(')')[0]).replace(' ', '')).replace('X', 'x')
                        if 'x' not in size:
                            size = ''
                        beam = beam.split('(')[0] # 取括號前內容即可
                    comma_char = ','
                    for char in weird_comma_list:
                        if char in beam:
                            comma_char = char
                    comma = beam.count(comma_char)
                    for i in range(comma + 1):
                        coor_to_beam_set.add(((coor1, coor2), (beam.split(comma_char)[i], size)))

                # 為了排版好看的怪產物，目前看到的格式為'{\W0.7;B4-2\P(80x100)}'，所以使用分號及反斜線來切
                # 切爛了也不會報錯，直接反映在結果
                if object.Layer in beam_layer and object.ObjectName == "AcDbMText" and object.GetBoundingBox()[0][1] >= 0:
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
                        coor_to_beam_set.add(((coor1, coor2), (beam, size)))

                # 找框框，完成block_coor_list，格式為((0.0, 0.0), (14275.54, 10824.61))
                # 此處不會報錯
                if object.Layer == block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    block_coor_list.append((coor1, coor2))

                if object.Layer == size_layer and object.EntityName == "AcDbText" and object.GetBoundingBox()[0][1] >= 0:
                    coor = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    if 'Gn' in object.TextString:
                        if 'C' in object.TextString and '(' in object.TextString:
                            coor_to_size_beam.add((coor, 'CG'))
                            coor_to_size_beam.add((coor, 'G'))
                        elif 'C' in object.TextString and '(' not in object.TextString:
                            coor_to_size_beam.add((coor, 'CG'))
                        else:
                            coor_to_size_beam.add((coor, 'G'))
                    if 'Bn' in object.TextString and 'W' not in object.TextString and 'D' not in object.TextString:
                        if 'C' in object.TextString and '(' in object.TextString:
                            coor_to_size_beam.add((coor, 'CB'))
                            coor_to_size_beam.add((coor, 'B'))
                        elif 'C' in object.TextString and '(' not in object.TextString:
                            coor_to_size_beam.add((coor, 'CB'))
                        else:
                            coor_to_size_beam.add((coor, 'B'))
                    if 'bn' in object.TextString:
                        if 'c' in object.TextString and '(' in object.TextString:
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
                            else:
                                if 'c' in object.TextString and '(' in object.TextString:
                                    coor_to_size_beam.add((coor, 'cg'))
                                    coor_to_size_beam.add((coor, 'g'))
                                elif 'c' in object.TextString and '(' not in object.TextString:
                                    coor_to_size_beam.add((coor, 'cg'))
                                else:
                                    coor_to_size_beam.add((coor, 'g'))
                    if 'x' in object.TextString or 'X' in object.TextString:
                        string = (object.TextString.replace(' ', '')).replace('X', 'x')
                        coor_to_size_string.add((coor, string))
            flag = 1

        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 6: {e}.')

    print('平面圖讀取進度 6/11')
    print('註: 如果不需要炸圖塊的話，會自動跳過Step 4, 5。')   
    # Step 7. 完成size_coor_set (size_beam, size_string, size_coor)
    size_coor_set = set()
    for x in coor_to_size_beam:
        coor = x[0]
        size_beam = x[1]
        min_size = ''
        min_dist = 10000
        for y in coor_to_size_string:
            coor2 = y[0]
            size_string = y[1]
            dist = abs(coor[0]-coor2[0]) + abs(coor[1] - coor2[1])
            if dist < min_dist:
                min_size = size_string
                min_dist = dist
        if min_size != '':
            size_coor_set.add((size_beam, min_size, coor))
    print('平面圖讀取進度 7/11')
    # Step 8. 透過 coor_to_floor_set 以及 block_coor_list 完成 floor_to_coor_set，格式為(floor, block左下角和右上角的coor)
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
    print('平面圖讀取進度 8/11')
    # Step 9. 算出Bmax, Fmax, Rmax
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
    print('平面圖讀取進度 9/11')
    # Step 10. 完成floor_beam_size_coor_set (floor, beam, size, coor)
    floor_beam_size_coor_set = set()
    for x in size_coor_set: # set(size_beam, min_size, coor)
        size_coor = x[2]
        size_string = x[1]
        size_beam = x[0]
        min_floor = []
        for z in floor_to_coor_set: # set (floor, block左下角和右上角的coor)
            floor_name = z[0]
            block_coor = z[1] 
            x_diff_left = size_coor[0] - block_coor[0][0] # 和左下角的diff
            y_diff_left = size_coor[1] - block_coor[0][1]
            x_diff_right = size_coor[0] - block_coor[1][0] # 和右上角的diff
            y_diff_right = size_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:                    
                if len(min_floor) == 0 or min_floor[0] != floor_name:
                    min_floor.append(floor_name)

        if len(min_floor) != 0:
            for i in range(len(min_floor)):
                floor = min_floor[i]                            
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
                                    floor_beam_size_coor_set.add((turn_floor_to_string(i), size_beam, size_string, size_coor))
                        except:
                            error(f'read_plan error in step 10: The error above is from here.')
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
                            floor_beam_size_coor_set.add((new_floor, size_beam, size_string, size_coor))
                        else:
                            error(f'read_plan error in step 10: new_floor is false.')
        else:
            error('read_plan error in step 10: min_floor cannot be found.')
    print('平面圖讀取進度 10/11')
    # Step 11. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補
    set_plan = set() # set元素為 (樓層, 梁柱名稱, size)
    dic_plan = {} # 透過(floor, beam, size)去找字串座標
    for x in coor_to_beam_set: # set(coor, (beam, size))
        beam_coor = x[0][0] # 取左下角即可
        full_coor = x[0] # 左下跟右上都有
        beam_name = x[1][0]
        beam_size = x[1][1]
        min_floor = ''
        for z in floor_to_coor_set: # set (floor, block左下角和右上角的coor)
            floor_name = z[0]
            block_coor = z[1] 
            x_diff_left = beam_coor[0] - block_coor[0][0] # 和左下角的diff
            y_diff_left = beam_coor[1] - block_coor[0][1]
            x_diff_right = beam_coor[0] - block_coor[1][0] # 和右上角的diff
            y_diff_right = beam_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:                    
                if min_floor == '' or min_floor == floor_name:
                    min_floor = floor_name

                else: # 有很多層在同一個block, 仍然透過字串的coor找樓層
                    new_min = 1000000
                    new_min_floor = ''
                    for y in coor_to_floor_set: # set(字串的coor, floor)
                        string_coor = y[0]
                        new_floor = y[1]
                        new_x_diff = abs(beam_coor[0] - string_coor[0])
                        new_y_diff = beam_coor[1] - string_coor[1]
                        new_total = new_x_diff + new_y_diff
                        if new_y_diff > 0 and new_total < new_min:
                            new_min = new_total
                            new_min_floor = new_floor
                    min_floor = new_min_floor

        floor = min_floor

        if floor != '':
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
                        error(f'read_plan error in step 11: The error above is from here.')
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
                        error(f'read_plan error in step 11: new_floor is false.')
            for floor in floor_list:
                if beam_size == '':
                    min_diff = 100000
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
                
                if beam_size != '':
                    set_plan.add((floor, beam_name, beam_size))
                    dic_plan[(floor, beam_name, beam_size)] = full_coor
                else:
                    error(f'read_plan error in step 10: {floor} {beam_name} cannot find size. ')

        else:
            error('read_plan error in step 11: min_floor cannot be found.')

    doc_plan.Close(SaveChanges=False)
    print('平面圖讀取進度 11/11')
    print('平面圖讀取完畢。')
    # plan.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in plan: \n")
    l = list(set_plan)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()

    return (set_plan, dic_plan)

def read_beam(beam_filename, text_layer, result_filename, explode):
    print('開始讀取梁配筋圖')
    # Step 1. 打開應用程式
    flag = 0
    while not flag:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 1: {e}.')
    print('梁配筋圖讀取進度 1/8')
    # Step 2. 匯入檔案
    flag = 0
    while not flag:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 2: {e}.')
    print('梁配筋圖讀取進度 2/8')
    # Step 3. 匯入modelspace
    flag = 0
    while not flag:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 3: {e}.')
    print('梁配筋圖讀取進度 3/8')
    if explode: # 需要提前炸圖塊再進來
        print('XS-BEAM 正在炸圖塊，運行時間取決於平面圖大小，請耐心等候')
        # Step 4. 遍歷所有物件 -> 炸圖塊
        # 炸圖塊看性質即可，不用看圖層      
        flag = 0
        while not flag:
            try:
                count = 0
                total = msp_beam.Count
                print(f'梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
                for object in msp_beam:
                    count += 1
                    if object.EntityName == "AcDbBlockReference" and object.Layer == text_layer:
                        object.Explode()
                    if count % 1000 == 0:
                        print(f'梁配筋圖已讀取{count}/{total}個物件')
                flag = 1
                print('梁配筋圖讀取進度 4/8')
            except Exception as e:
                time.sleep(5)
                error(f'read_beam error in step 4: {e}.')

        # Step 5. 重新匯入modelspace
        flag = 0
        while not flag:
            try:
                msp_beam = doc_beam.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'read_beam error in step 5: {e}.')
        print('梁配筋圖讀取進度 5/8')
    
    # Step 6. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    print('正在遍歷梁配筋圖上所有物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候')
    floor_to_beam_set = set()
    flag = 0
    while not flag:
        try:
            count = 0
            total = msp_beam.Count
            print(f'梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候')
            for object in msp_beam:
                count += 1
                if count % 1000 == 0:
                    print(f'梁配筋圖已讀取{count}/{total}個物件')
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
            time.sleep(5)
            error(f'read_beam error in step 4: {e}.')
    print('梁配筋圖讀取進度 6/8')
    # Step 7. 算出Bmax, Fmax, Rmax
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
    print('梁配筋圖讀取進度 7/8')
    # Step 8. 完成set_beam和dic_beam
    dic_beam = {}
    set_beam = set()
    for x in floor_to_beam_set:
        floor = x[0]
        beam = x[1]
        coor = x[2]
        size = x[3]
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
                            set_beam.add((turn_floor_to_string(i), beam, size))
                            dic_beam[(turn_floor_to_string(i), beam, size)] = coor
                except:
                    error(f'read_beam error in step 8: The error above is from here.')
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
                    set_beam.add((new_floor, beam, size))
                    dic_beam[(new_floor, beam, size)] = coor
                else:
                    error(f'read_beam error in step 8: new_floor is false.')

    doc_beam.Close(SaveChanges=False)
    print('梁配筋圖讀取進度 8/8')
    print('梁配筋圖讀取完成。')
    # beam.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in beam: \n")
    l = list(set_beam)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()
    
    return (set_beam, dic_beam)

def write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, drawing): # 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
    print("開始標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'。")
    pythoncom.CoInitialize()
    set1 = set_plan - set_beam
    list1 = list(set1)
    list1.sort()
    set2 = set_beam - set_plan
    list2 = list(set2)
    list2.sort()

    f_big = open(big_file, "w", encoding = 'utf-8')
    f_sml = open(sml_file, "w", encoding = 'utf-8')

    f_big.write("in plan but not in beam: \n")
    f_sml.write("in plan but not in beam: \n")
    if drawing:
        # Step 1. 開啟應用程式
        flag = 0
        while not flag:
            try:
                wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_plan error in step 1, {e}')
        # Step 2. 匯入檔案
        flag = 0
        while not flag:
            try:
                doc_plan = wincad_plan.Documents.Open(plan_filename)
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_plan error in step 2, {e}')
        # Step 3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag:
            try:
                msp_plan = doc_plan.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_plan error in step 3, {e}')
        time.sleep(5)

        # Step 4. 設定mark的圖層
        flag = 0
        while not flag:
            try:
                layer_plan = doc_plan.Layers.Add(f"S-CLOUD_{date}")
                doc_plan.ActiveLayer = layer_plan
                layer_plan.color = 10
                layer_plan.Linetype = "Continuous"
                layer_plan.Lineweight = 0.5
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_plan error in step 4, {e}')
    
    # Step 5. 完成in plan but not in beam，畫圖，以及計算錯誤率
    big_error = 0
    sml_error = 0
    err_list_big = []
    err_list_sml = []
    for x in list1: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
            wrong_data = 0
            for y in list2:
                if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                    err_list_big.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                    wrong_data = 1
                    break
            if not wrong_data:
                err_list_big.append((x, 1)) # type(tuple of floor and wrong beam, err_message)   
            big_error += 1
        else:
            wrong_data = 0
            for y in list2:
                if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                    err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                    # f_sml.write(f'{x}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
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
    
    try:
        sml_rate = round(sml_error / sml_count * 100, 2)
        f_sml.write(f'error rate = {sml_rate} %\n')
    except:
        sml_rate = 'unfinish'
        error(f'write_plan error in step 5, there are no small beam in plan.txt?')

    f_big.close()
    f_sml.close()
    print("標註平面圖(核對項目: 梁配筋)及輸出核對結果至'大梁.txt'和'小梁.txt'完成。")
    return (big_rate, sml_rate)
    

def write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, drawing): # 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
    print("開始標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'。")
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
        while not flag:
            try:
                wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_beam error in step 1, {e}')
        # Step 2. 匯入檔案
        flag = 0
        while not flag:
            try:
                doc_beam = wincad_beam.Documents.Open(beam_filename)
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_beam error in step 2, {e}')
        # Step 3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag:
            try:
                msp_beam = doc_beam.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_beam error in step 3, {e}')
        time.sleep(5)

        # Step 4. 設定mark的圖層
        flag = 0
        while not flag:
            try:
                layer_beam = doc_beam.Layers.Add(f"S-CLOUD_{date}")
                doc_beam.ActiveLayer = layer_beam
                layer_beam.color = 10
                layer_beam.Linetype = "Continuous"
                layer_beam.Lineweight = 0.5
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_beam error in step 4, {e}')

    # Step 5. 完成in plan but not in beam，畫圖，以及計算錯誤率
    big_error = 0
    sml_error = 0
    err_list_big = []
    err_list_sml = []
    for x in list2: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
            wrong_data = 0
            for y in list1:
                if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                    err_list_big.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                    wrong_data = 1
                    break
            if not wrong_data:
                err_list_big.append((x, 1)) # type(tuple of floor and wrong beam, err_message)   
            big_error += 1
        else:
            wrong_data = 0
            for y in list1:
                if x[0] == y[0] and x[1] == y[1] and x[2] != y[2]:
                    err_list_sml.append((x, 0, y[2])) # type(tuple of floor and wrong beam, err_message, correct)
                    # f_sml.write(f'{x}: 尺寸有誤，在XS-BEAM那邊是{y[2]}\n')
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
        
    err_list_big = sorted(err_list_big, key = cmp_to_key(mycmp))
    err_list_sml = sorted(err_list_sml, key = cmp_to_key(mycmp))

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
    
    f_big.close()
    f_sml.close()
    print("標註梁配筋圖及輸出核對結果至'大梁.txt'和'小梁.txt'完成。")
    return (big_rate, sml_rate)

def write_result_log(excel_file, task_name, plan_not_beam_big, plan_not_beam_sml, beam_not_plan_big, beam_not_plan_sml, date, runtime, other):
    sheet_name = 'result_log'
    new_list = [(task_name, plan_not_beam_big, plan_not_beam_sml, beam_not_plan_big, beam_not_plan_sml, date, runtime, other)]
    dfNew=pd.DataFrame(new_list, columns = ['名稱' , 'in plan not in beam 大梁', 'in plan not in beam 小梁','in beam not in plan 大梁', 'in plan not In beam 小梁', '執行時間', '執行日期' , '備註'])
    if os.path.exists(excel_file):
        writer = pd.ExcelWriter(excel_file,engine='xlsxwriter')
        df = pd.read_excel(excel_file) 
        df = pd.concat([df, dfNew], axis=0, ignore_index = True, join = 'inner')
    else:
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter') 
        df = dfNew
    df.to_excel(writer,sheet_name=sheet_name)
    writer.save()    
    return

error_file = './result/error_log.txt' # error_log.txt的路徑

if __name__=='__main__':
    start = time.time()
    task_name = 'task20'#sys.argv[13]
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    plan_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\清豐1F大梁-XS-PLAN.dwg'#sys.argv[2] # XS-PLAN的路徑
    beam_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\清豐1F大梁-1FHB.dwg'#sys.argv[1] # XS-BEAM的路徑
    plan_new_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\XS-PLAN_new'#sys.argv[4] # XS-PLAN_new的路徑
    beam_new_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\XS-BEAM_new'#sys.argv[3] # XS-BEAM_new的路徑
    plan_file = './result/plan.txt' # plan.txt的路徑
    beam_file = './result/beam.txt' # beam.txt的路徑
    excel_file = './result/result_log.xlsx' # result_log.xlsx的路徑
    big_file = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\big.txt'#sys.argv[5] # 大梁結果
    sml_file = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task21-31000\sml.txt'#sys.argv[6] # 小梁結果

    date = time.strftime("%Y-%m-%d", time.localtime())
    
    # 在plan裡面自訂圖層
    floor_layer = 'S-TITLE'#sys.argv[9] # 樓層字串的圖層
    beam_layer = ['S-TEXTB', 'S-TEXTG']#[sys.argv[10], sys.argv[11]] # beam的圖層，因為有兩個以上，所以用list來存
    block_layer = 'DwFm'#sys.argv[8] # 框框的圖層
    explode_plan = 0#sys.argv[14] # XS-PLAN需不需要提前炸圖塊(0:不需要 1:需要)
    explode_beam = 0#sys.argv[15] # XS-BEAM需不需要提前炸圖塊(0:不需要 1:需要)
    size_layer = 'S-TEXT'#sys.argv[12] # 梁尺寸字串圖層

    # 在beam裡面自訂圖層
    text_layer = 'S-RC'#sys.argv[7]

    # 多檔案接用','來連接，不用空格。Ex. 'file1,file2,file3'
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
        res_plan[i] = pool.apply_async(read_plan, (plan_filename.split(',')[i], floor_layer, beam_layer, block_layer, size_layer, plan_file, explode_plan))

    for i in range(beam_file_count):
        res_beam[i] = pool.apply_async(read_beam, (beam_filename.split(',')[i], text_layer, beam_file, explode_beam))
    
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
    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date, drawing)
    beam_result = write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date, drawing)

    end = time.time()
    # write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    # write_result_log(excel_file,'','','','','','',time.strftime("%Y-%m-%d %H:%M", time.localtime()),'')