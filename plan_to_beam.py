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

weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡']
beam_head1 = ['B', 'b', 'G', 'g']
beam_head2 = ['CB', 'CG', 'cb']

def read_plan(plan_filename, floor_layer, beam_layer, block_layer, result_filename, explode):
    # Step 1. 打開應用程式
    flag = 0
    while not flag:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 1: {e}.')

    # Step 2. 匯入檔案
    flag = 0
    while not flag:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 2: {e}.')

    # Step 3. 匯入modelspace
    flag = 0
    while not flag:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 3: {e}.')

    if explode: # 需要提前炸圖塊再進來

        # Step 4. 遍歷所有物件 -> 炸圖塊
        # 炸圖塊看性質即可，不用看圖層      
        flag = 0
        while not flag:
            try:
                for object in msp_plan:
                    if object.EntityName == "AcDbBlockReference":
                        object.Explode()
                flag = 1
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
    
    # Step 6. 遍歷所有物件 -> 完成 coor_to_floor_set, coor_to_beam_set, block_coor_list
    coor_to_floor_set = set() # set (字串的coor, floor)
    coor_to_beam_set = set() # set (coor, beam)
    block_coor_list = [] # 存取方框最左下角的點座標

    flag = 0
    while not flag:
        try:
            for object in msp_plan:
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
                    if '(' in beam:
                        beam = beam.split('(')[0] # 取括號前內容即可
                    comma_char = ','
                    for char in weird_comma_list:
                        if char in beam:
                            comma_char = char
                    comma = beam.count(comma_char)
                    for i in range(comma + 1):
                        coor_to_beam_set.add(((coor1, coor2), beam.split(comma_char)[i]))

                # 為了排版好看的怪產物，目前看到的格式為'{\W0.7;B4-2\P(80x100)}'，所以使用分號及反斜線來切
                # 切爛了也不會報錯，直接反映在結果
                if object.Layer in beam_layer and object.ObjectName == "AcDbMText" and object.GetBoundingBox()[0][1] >= 0:
                    beam = object.TextString
                    semicolon = beam.count(';')
                    for i in range(semicolon + 1):
                        s = beam.split(';')[i]
                        if s[0] in beam_head1 or s[0:2] in beam_head2:
                            if '(' in s:
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
                        coor_to_beam_set.add(((coor1, coor2), beam))

                # 找框框，完成block_coor_list，格式為((0.0, 0.0), (14275.54, 10824.61))
                # 此處不會報錯
                if object.Layer == block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    block_coor_list.append((coor1, coor2))
            flag = 1

        except Exception as e:
            time.sleep(5)
            error(f'read_plan error in step 6: {e}.')

    # Step 7. 透過 coor_to_floor_set 以及 block_coor_list 完成 floor_to_coor_set，格式為(floor, block左下角和右上角的coor)
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

    # Step 8. 算出Bmax, Fmax, Rmax
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

    # Step 9. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補
    set_plan = set() # set元素為 (樓層, 梁柱名稱)
    dic_plan = {} # 透過(floor, beam)去找字串座標
    for x in coor_to_beam_set: # set(coor, beam)
        beam_coor = x[0][0] # 取左下角即可
        full_coor = x[0] # 左下跟右上都有
        beam_name = x[1]
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
                                set_plan.add((turn_floor_to_string(i), beam_name))
                                dic_plan[(turn_floor_to_string(i), beam_name)] = full_coor
                    except:
                        error(f'read_plan error in step 9: The error above is from here.')
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
                        set_plan.add((new_floor, beam_name))
                        dic_plan[(new_floor, beam_name)] = full_coor
                    else:
                        error(f'read_plan error in step 9: new_floor is false.')
        else:
            error('read_plan error in step 9: min_floor cannot be found.')

    doc_plan.Close(SaveChanges=False)

    # plan.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in plan: \n")
    l = list(set_plan)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()

    return (set_plan, dic_plan)

def read_beam(beam_filename, text_layer, result_filename):
    # Step 1. 打開應用程式
    flag = 0
    while not flag:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 1: {e}.')
    # Step 2. 匯入檔案
    flag = 0
    while not flag:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 2: {e}.')
    # Step 3. 匯入modelspace
    flag = 0
    while not flag:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 3: {e}.')
    
    # Step 4. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor)
    floor_to_beam_set = set()
    flag = 0
    while not flag:
        try:
            for object in msp_beam:
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
                            floor_to_beam_set.add((floor, beam, (coor1, coor2)))
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_beam error in step 4: {e}.')

    # Step 5. 算出Bmax, Fmax, Rmax
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

    # Step 6. 完成set_beam和dic_beam
    dic_beam = {}
    set_beam = set()
    for x in floor_to_beam_set:
        floor = x[0]
        beam = x[1]
        coor = x[2]
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
                            set_beam.add((turn_floor_to_string(i), beam))
                            dic_beam[(turn_floor_to_string(i), beam)] = coor
                except:
                    error(f'read_beam error in step 6: The error above is from here.')
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
                    set_beam.add((new_floor, beam))
                    dic_beam[(new_floor, beam)] = coor
                else:
                    error(f'read_beam error in step 6: new_floor is false.')

    doc_beam.Close(SaveChanges=False)

    # beam.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w")
    f.write("in beam: \n")
    l = list(set_beam)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()
    
    return (set_beam, dic_beam)

def write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date): # 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
    pythoncom.CoInitialize()
    set1 = set_plan - set_beam
    list1 = list(set1)
    list1.sort()

    f_big = open(big_file, "w")
    f_sml = open(sml_file, "w")

    f_big.write("in plan but not in beam: \n")
    f_sml.write("in plan but not in beam: \n")
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

    for x in list1: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':        
            f_big.write(f'{x}\n')
            big_error += 1
        else:
            f_sml.write(f'{x}\n')
            sml_error += 1
        
        coor = dic_plan[x]
        coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
        points = vtFloat(coor_list)
        pointobj = msp_plan.AddPolyline(points)
        for i in range(4):
            pointobj.SetWidth(i, 10, 10)
    
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
    return (big_rate, sml_rate)

def write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date): # 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
    pythoncom.CoInitialize()
    set2 = set_beam - set_plan
    list2 = list(set2)
    list2.sort()

    f_big = open(big_file, "a")
    f_sml = open(sml_file, "a")

    f_big.write("in beam but not in plan: \n")
    f_sml.write("in beam but not in plan: \n")
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
    for x in list2: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
            f_big.write(f'{x}\n')
            big_error += 1
        else:
            f_sml.write(f'{x}\n')
            sml_error += 1
        
        coor = dic_beam[x]
        coor_list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
        points = vtFloat(coor_list)
        pointobj = msp_beam.AddPolyline(points)
        for i in range(4):
            pointobj.SetWidth(i, 10, 10)

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
    task_name = sys.argv[12]
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    plan_filename = sys.argv[2] # XS-PLAN的路徑
    beam_filename = sys.argv[1] # XS-BEAM的路徑
    plan_new_filename = sys.argv[4] # XS-PLAN_new的路徑
    beam_new_filename = sys.argv[3] # XS-BEAM_new的路徑
    plan_file = './result/plan.txt' # plan.txt的路徑
    beam_file = './result/beam.txt' # beam.txt的路徑
    excel_file = './result/result_log.xlsx' # result_log.xlsx的路徑
    big_file = sys.argv[5] # 大梁結果
    sml_file = sys.argv[6] # 小梁結果

    date = time.strftime("%Y-%m-%d", time.localtime())
    
    # 在plan裡面自訂圖層
    floor_layer = sys.argv[9] # 樓層字串的圖層
    beam_layer = [sys.argv[10], sys.argv[11]] # beam的圖層，因為有兩個以上，所以用list來存
    block_layer = sys.argv[8] # 框框的圖層
    explode = sys.argv[13] # 需不需要提前炸圖塊(0:不需要 1:需要)

    # 在beam裡面自訂圖層
    text_layer = sys.argv[7]
    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()
    res_plan = pool.apply_async(read_plan, (plan_filename, floor_layer, beam_layer, block_layer, plan_file, explode))
    res_beam = pool.apply_async(read_beam, (beam_filename, text_layer, beam_file))
    final_plan = res_plan.get()
    final_beam = res_beam.get()

    set_plan = final_plan[0]
    dic_plan = final_plan[1]
    set_beam = final_beam[0]
    dic_beam = final_beam[1]

    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_beam, dic_plan, big_file, sml_file, date)
    beam_result = write_beam(beam_filename, beam_new_filename, set_plan, set_beam, dic_beam, big_file, sml_file, date)

    end = time.time()
    # write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    # write_result_log(excel_file,'','','','','','',time.strftime("%Y-%m-%d %H:%M", time.localtime()),'')