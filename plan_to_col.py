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

from plan_to_beam import turn_floor_to_float, turn_floor_to_string, floor_exist, vtFloat, error

weird_to_list = ['-', '~']
weird_comma_list = [',', '、', '¡']

def read_plan(plan_filename, floor_layer, col_layer, block_layer, result_filename, explode):
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

    if explode:
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
    
    # Step 6. 遍歷所有物件 -> 完成 coor_to_floor_set, coor_to_col_set, block_coor_list
    coor_to_floor_set = set() # set (字串的coor, floor)
    coor_to_col_set = set() # set (coor, col)
    coor_to_size_set = set() # set (coor, size)
    block_coor_list = [] # 存取方框最左下角的點座標

    flag = 0
    while not flag:
        try:
            for object in msp_plan:
                # 取floor的字串 -> 抓括號內的字串 (Ex. '十層至十四層結構平面圖(10F~14F)' -> '10F~14F')
                # 若此處報錯，可能原因: 1. 沒有括號, 2. 待補
                if object.Layer == floor_layer and object.ObjectName == "AcDbText" and '(' in object.TextString and object.InsertionPoint[1] >= 0:
                    floor = object.TextString
                    floor = re.search('\(([^)]+)', floor).group(1) #取括號內的樓層數
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2)) #不取概數的話後面抓座標會出問題，例如兩個樓層在同一格
                    no_chinese = False
                    for ch in floor: # 待修正
                        if ch == 'F' or ch.isdigit():
                            no_chinese = True
                            break
                    if floor != '' and no_chinese:
                        coor_to_floor_set.add((coor, floor))
                    else:
                        error(f'read_plan error in step 6: floor is an empty string. ')
                # 取col的字串
                if object.Layer == col_layer and object.ObjectName == "AcDbText" and (object.TextString[0] == 'C' or ('¡æ' in object.TextString and 'C' in object.TextString)) and 'S' not in object.TextString:
                    col = f"C{object.TextString.split('C')[1]}"
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    coor_to_col_set.add(((coor1, coor2), col))

                # 取size
                if object.Layer == col_layer and object.ObjectName == "AcDbText" and '(' in object.TextString:
                    size = (object.TextString.split('(')[1]).split(')')[0] # 取括號內東西即可
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    if 'x' in size:
                        coor_to_size_set.add(((coor1, coor2), size))
                    else:
                        coor_to_size_set.add(((coor1, coor2), '另詳'))

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

    # Step 9. 完成col_size_coor_set，格式: set(col, size, the coor of big_block(left, right, up, down))
    col_size_coor_set = set() 
    for x in coor_to_col_set:
        col_coor = x[0][0]
        col_full_coor = x[0]
        col_name = x[1]
        min_diff = 1000
        match_size = ''
        match_size_coor = ''
        for y in coor_to_size_set:
            size_coor = y[0][0]
            size_full_coor = y[0]
            size = y[1]
            x_diff = abs(col_coor[0] - size_coor[0])
            y_diff = abs(col_coor[1] - size_coor[1])
            if x_diff + y_diff < min_diff:
                min_diff = x_diff + y_diff
                match_size = size
                match_size_coor = size_full_coor
        
        if min_diff != 1000 and match_size != '' and match_size_coor != '':
            left = min(col_full_coor[0][0], match_size_coor[0][0])
            right = max(col_full_coor[1][0], match_size_coor[1][0])
            up = max(col_full_coor[1][1], match_size_coor[1][1])
            down = min(col_full_coor[0][1], match_size_coor[0][1])

            col_size_coor_set.add((col_name, match_size, (left, right, up, down)))

    # # DEBUG: 檢查col跟size有沒有被圈在一起，或者被亂圈到其他地方
    # for x in col_size_coor_set:
    #     coor = x[2]
    #     coor_list = [coor[0] - 20, coor[3] - 20, 0, coor[1] + 20, coor[3] - 20, 0, coor[1] + 20, coor[2] + 20, 0, coor[0] - 20, coor[2] + 20, 0, coor[0] - 20, coor[3] - 20, 0]
    #     points = vtFloat(coor_list)
    #     pointobj = msp_plan.AddPolyline(points)
    #     for i in range(4):
    #         pointobj.SetWidth(i, 10, 10)
    # return

    # Step 10. 完成 set_plan 以及 dic_plan
    # 此處可能錯的地方在於找不到min_floor，可能原因: 1. 框框沒有被掃到, 導致東西在框框外面找不到家，2. 待補
    set_plan = set() # set元素為 (樓層, col, size)
    dic_plan = {} # 透過(樓層, col, size)去找col跟size的整體座標
    for x in col_size_coor_set: # set(col, size, the coor of big_block(left, right, up, down))
        col_coor = (x[2][0], x[2][3]) # 取左下角即可
        full_coor = x[2] # 左下跟右上都有
        col_name = x[0]
        col_size = x[1]
        min_floor = ''
        for z in floor_to_coor_set: # set (floor, block左下角和右上角的coor)
            floor_name = z[0]
            block_coor = z[1] 
            x_diff_left = col_coor[0] - block_coor[0][0] # 和左下角的diff
            y_diff_left = col_coor[1] - block_coor[0][1]
            x_diff_right = col_coor[0] - block_coor[1][0] # 和右上角的diff
            y_diff_right = col_coor[1] - block_coor[1][1]
            if x_diff_left > 0 and y_diff_left > 0 and x_diff_right < 0 and y_diff_right < 0:                    
                if min_floor == '' or min_floor == floor_name:
                    min_floor = floor_name

                else: # 有很多層在同一個block, 仍然透過字串的coor找樓層
                    new_min = 1000000
                    new_min_floor = ''
                    for y in coor_to_floor_set: # set(字串的coor, floor)
                        string_coor = y[0]
                        new_floor = y[1]
                        new_x_diff = abs(col_coor[0] - string_coor[0])
                        new_y_diff = col_coor[1] - string_coor[1]
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
                                set_plan.add((turn_floor_to_string(i), col_name, col_size))
                                dic_plan[(turn_floor_to_string(i), col_name, col_size)] = full_coor
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
                        set_plan.add((new_floor, col_name, col_size))
                        dic_plan[(new_floor, col_name, col_size)] = full_coor
                    else:
                        error(f'read_plan error in step 10: new_floor is false.')
        else:
            error('read_plan error in step 10: min_floor cannot be found.')

    doc_plan.Close(SaveChanges=False)

    # plan.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w", encoding = 'utf-8')
    f.write("in plan: \n")
    l = list(set_plan)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()

    return (set_plan, dic_plan)

def read_col(col_filename, text_layer, line_layer, result_filename, explode):
    # Step 1. 打開應用程式
    flag = 0
    while not flag:
        try:
            wincad_col = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_col error in step 1: {e}.')
    # Step 2. 匯入檔案
    flag = 0
    while not flag:
        try:
            doc_col = wincad_col.Documents.Open(col_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_col error in step 2: {e}.')
    # Step 3. 匯入modelspace
    flag = 0
    while not flag:
        try:
            msp_col = doc_col.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            error(f'read_col error in step 3: {e}.')

    if explode:
        # Step 4. 遍歷所有物件 -> 炸圖塊
        # 炸圖塊看性質即可，不用看圖層   
        flag = 0
        while not flag:
            try:
                for object in msp_col:
                    if object.EntityName == "AcDbBlockReference":
                        object.Explode()
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'read_col error in step 4: {e}.')

        # Step 5. 重新匯入modelspace
        flag = 0
        while not flag:
            try:
                msp_col = doc_col.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'read_col error in step 5: {e}.')
    
    # Step 6. 遍歷所有物件 -> 完成一堆座標對應的set跟list
    coor_to_floor_set = set() # set(coor, floor)
    coor_to_col_set = set() # set(coor, col)
    coor_to_size_set = set() # set(coor, size)
    coor_to_floor_line_list = [] # (橫線y座標, start, end)
    coor_to_col_line_list = [] # (縱線x座標, start, end)
    flag = 0
    while not flag:
        try:
            for object in msp_col:
                if object.Layer in text_layer and object.ObjectName == "AcDbText": 
                    if object.TextString[0] == 'C' and (len(object.TextString) <= 7):
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        coor_to_col_set.add(((coor1, coor2), object.TextString))

                    elif 'x' in object.TextString:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        coor_to_size_set.add(((coor1, coor2), object.TextString))
                    elif 'F' in object.TextString or 'B' in object.TextString or 'R' in object.TextString: # 可能有樓層
                        floor = object.TextString
                        if '_' in floor: # 可能有B_6F表示B棟的6F
                            floor = floor.split('_')[1]
                        if turn_floor_to_float(floor):
                            coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                            coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                            floor = turn_floor_to_float(floor)
                            floor = turn_floor_to_string(floor)
                            coor_to_floor_set.add(((coor1, coor2), floor))
                
                if object.Layer in line_layer:
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    if coor1[0] == coor2[0]:
                        coor_to_col_line_list.append((coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1])))
                    elif coor1[1] == coor2[1]:
                        coor_to_floor_line_list.append((coor1[1], min(coor1[0], coor2[0]), max(coor1[0], coor2[0])))
            flag = 1
            coor_to_col_line_list.sort(key = lambda x: x[0])
            coor_to_floor_line_list.sort(key = lambda x: x[0])
        except Exception as e:
            time.sleep(5)
            error(f'read_col error in step 6: {e}.')
    
    # Step 7. 完成col_to_line_set 格式:(col, left, right, up)
    col_to_line_set = set()
    for x in coor_to_col_set:
        coor = x[0]
        col = x[1]
        new_coor_to_col_line_list = []
        for y in coor_to_col_line_list: # (縱線x座標, start, end)
            if y[1] <= coor[0][1] <= y[2]: # 先看y座標有沒有被夾住
                new_coor_to_col_line_list.append(y)
        tmp_set = set(new_coor_to_col_line_list)
        new_coor_to_col_line_list = list(tmp_set)
        new_coor_to_col_line_list.sort(key = lambda x: x[0])
        for y in range(len(new_coor_to_col_line_list)): # 再看x座標被哪兩條線夾住
            if new_coor_to_col_line_list[y][0] < coor[0][0] < new_coor_to_col_line_list[y+1][0]:
                col_to_line_set.add((col, new_coor_to_col_line_list[y][0], new_coor_to_col_line_list[y+1][0], coor[1][1]))

    # Step 8. 完成floor_to_line_set 格式:(floor, down, up, left)
    floor_to_line_set = set()
    for x in coor_to_floor_set:
        coor = x[0]
        floor = x[1]
        new_coor_to_floor_line_list = []
        for y in coor_to_floor_line_list: # (橫線y座標, start, end)
            if y[1] <= coor[0][0] <= y[2]: # 先看x座標有沒有被夾住
                new_coor_to_floor_line_list.append(y)
        tmp_set = set(new_coor_to_floor_line_list)
        new_coor_to_floor_line_list = list(tmp_set)
        new_coor_to_floor_line_list.sort(key = lambda x: x[0])
        for y in range(len(new_coor_to_floor_line_list)): # 再看y座標被哪兩條線夾住，下面那條要往下平移一格
            if new_coor_to_floor_line_list[y][0] < coor[0][1] < new_coor_to_floor_line_list[y+1][0]:
                floor_to_line_set.add((floor, new_coor_to_floor_line_list[y-1][0], new_coor_to_floor_line_list[y+1][0], coor[0][0]))

    # Step 9. 完成set_col和dic_col
    dic_col = {}
    set_col = set()
    for x in coor_to_size_set:
        coor = x[0]
        size = x[1]
        min_floor = '' 
        min_floor_coor = ''
        min_floor_diff = 10000
        min_col = ''
        min_col_coor = ''
        min_col_diff = 10000
        for y in floor_to_line_set: # (floor, down, up, left)
            if y[1] <= coor[1][1] <= y[2] and coor[1][0] - y[3] >= 0 and coor[1][0] - y[3] <= min_floor_diff:
                min_floor = y[0]
                min_floor_coor = (y[1], y[2])
                min_floor_diff = coor[1][0] - y[3]
        for y in col_to_line_set: # (col, left, right, up)
            if y[1] <= coor[1][0] <= y[2] and y[3] - coor[1][1] >= 0 and y[3] - coor[1][1] <= min_col_diff:
                min_col = y[0]
                min_col_coor = (y[1], y[2])
                min_col_diff = y[3] - coor[1][1]
        if min_floor != '' and min_col != '':
            if '-' in min_col:
                try:
                    start = int((min_col.split('-')[0]).split('C')[1])
                    end = int((min_col.split('-')[1]).split('C')[1])
                    for i in range(start, end + 1):
                        set_col.add((min_floor, f'C{i}', size))
                        dic_col[(min_floor, f'C{i}', size)] = (min_col_coor[0], min_col_coor[1], min_floor_coor[1], min_floor_coor[0]) # (left, right, up, down)
                except: # CW3-1 之類的不是區間
                    set_col.add((min_floor, min_col, size))
                    dic_col[(min_floor, min_col, size)] = (min_col_coor[0], min_col_coor[1], min_floor_coor[1], min_floor_coor[0]) # (left, right, up, down)
            else:
                set_col.add((min_floor, min_col, size))
                dic_col[(min_floor, min_col, size)] = (min_col_coor[0], min_col_coor[1], min_floor_coor[1], min_floor_coor[0]) # (left, right, up, down)
    
    doc_col.Close(SaveChanges=False)

    # col.txt單純debug用，不想多新增檔案可以註解掉
    f = open(result_filename, "w", encoding = 'utf-8')
    f.write("in col: \n")
    l = list(set_col)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()
    
    return (set_col, dic_col)

def write_plan(plan_filename, plan_new_filename, set_plan, set_col, dic_plan, result_filename, date, drawing): # 完成 in plan but not in col 的部分並在圖上mark有問題的部分
    set1 = set_plan - set_col
    list1 = list(set1)
    list1.sort()
    set2 = set_col - set_plan
    list2 = list(set2)
    list2.sort()

    f = open(result_filename, "w", encoding = 'utf-8')

    f.write("in plan but not in col: \n")

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
    
    # Step 5. 完成in plan but not in col，畫圖，以及計算錯誤率
    error_num = 0

    for x in list1: 
        if x[0] != 'FBF':
            wrong_data = 0
            for y in list2:
                if x[0] == y[0] and x[1] == y[1]:
                    f.write(f'{x}: 尺寸有誤，在XS-COL那邊是{y[2]}\n')
                    wrong_data = 1
                    break
            if not wrong_data:
                f.write(f'{x}: 找不到這根柱子\n')
            
            error_num += 1
            
            if drawing:
                coor = dic_plan[x]
                coor_list = [coor[0] - 20, coor[3] - 20, 0, coor[1] + 20, coor[3] - 20, 0, coor[1] + 20, coor[2] + 20, 0, coor[0] - 20, coor[2] + 20, 0, coor[0] - 20, coor[3] - 20, 0]
                points = vtFloat(coor_list)
                pointobj = msp_plan.AddPolyline(points)
                for i in range(4):
                    pointobj.SetWidth(i, 10, 10)
    
    if drawing:
        doc_plan.SaveAs(plan_new_filename)
        doc_plan.Close(SaveChanges=True)

    count = 0
    for x in set_plan:
        count += 1
    
    # 計算錯誤率可能會噴錯，因為分母為0
    try:
        rate = round(error_num / count * 100, 2)
        f.write(f'error rate = {rate} %\n')
    except:
        rate = 'unfinish'
        error(f'write_plan error in step 5, there are no col in plan.txt?')

    f.close()
    return rate

def write_col(col_filename, col_new_filename, set_plan, set_col, dic_col, result_filename, date, drawing): # 完成 in beam but not in plan 的部分並在圖上mark有問題的部分
    set1 = set_plan - set_col
    list1 = list(set1)
    list1.sort()
    set2 = set_col - set_plan
    list2 = list(set2)
    list2.sort()

    f = open(result_filename, "a", encoding = 'utf-8')
    f.write("in col but not in plan: \n")

    if drawing:
        # Step 1. 開啟應用程式
        flag = 0
        while not flag:
            try:
                wincad_col = win32com.client.Dispatch("AutoCAD.Application")
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_col error in step 1, {e}')
        # Step 2. 匯入檔案
        flag = 0
        while not flag:
            try:
                doc_col = wincad_col.Documents.Open(col_filename)
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_col error in step 2, {e}')
        # Step 3. 載入modelspace(還要畫圖)
        flag = 0
        while not flag:
            try:
                msp_col = doc_col.Modelspace
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_col error in step 3, {e}')
        time.sleep(5)

        # Step 4. 設定mark的圖層
        flag = 0
        while not flag:
            try:
                layer_col = doc_col.Layers.Add(f"S-CLOUD_{date}")
                doc_col.ActiveLayer = layer_col
                layer_col.color = 10
                layer_col.Linetype = "Continuous"
                layer_col.Lineweight = 0.5
                flag = 1
            except Exception as e:
                time.sleep(5)
                error(f'write_col error in step 4, {e}')

    # Step 5. 完成in plan but not in col，畫圖，以及計算錯誤率
    error_num = 0
    for x in list2: 
        wrong_data = 0
        for y in list1:
            if x[0] == y[0] and x[1] == y[1]:
                f.write(f'{x}: 尺寸有誤，在XS-PLAN那邊是{y[2]}\n')
                wrong_data = 1
                break
        if not wrong_data:
            f.write(f'{x}: 找不到這根柱子\n')

        error_num += 1
        
        if drawing:
            coor = dic_col[x]
            coor_list = [coor[0], coor[3], 0, coor[1], coor[3], 0, coor[1], coor[2], 0, coor[0], coor[2], 0, coor[0], coor[3], 0]
            points = vtFloat(coor_list)
            pointobj = msp_col.AddPolyline(points)
            for i in range(4):
                pointobj.SetWidth(i, 10, 10)

    if drawing:
        doc_col.SaveAs(col_new_filename)
        doc_col.Close(SaveChanges=True)

    count = 0

    for x in set_col:
        count += 1
    
    # 計算錯誤率可能會噴錯，因為分母為0
    try:
        rate = round(error_num / count * 100, 2)
        f.write(f'error rate = {rate} %\n')
    except:
        rate = 'unfinish'
        error(f'write_col error in step 5, there are no col in col.txt?')
    
    f.close()

    return rate

# 待改
def write_result_log(excel_file, task_name, plan_not_col, col_not_plan,date, runtime, other):
    sheet_name = 'result_log'
    new_list = [(task_name, plan_not_col, col_not_plan, date, runtime, other)]
    dfNew=pd.DataFrame(new_list, columns = ['名稱' , 'in plan not in col 柱', 'in col not in plan 柱', '執行時間', '執行日期' , '備註'])
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
    task_name = 'task18' #sys.argv[11]
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    plan_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task18\XS-PLAN.dwg' #sys.argv[2] # XS-PLAN的路徑
    col_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task18\XS-COL.dwg' #sys.argv[1] # XS-COL的路徑
    plan_new_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task18\XS-PLAN_new.dwg' #sys.argv[4] # XS-PLAN_new的路徑
    col_new_filename = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task18\XS-COL_new.dwg' #sys.argv[3] # XS-COL_new的路徑
    plan_file = './result/col_plan.txt' # plan.txt的路徑
    col_file = './result/col.txt' # col.txt的路徑
    excel_file = './result/result_log_col.xlsx' # result_log.xlsx的路徑
    result_file = r'K:\100_Users\EI 202208 Bamboo\BeamQC\task18\柱配筋.txt' #sys.argv[5] # 柱配筋結果

    date = time.strftime("%Y-%m-%d", time.localtime())
    
    # 在plan裡面自訂圖層
    floor_layer = 'S-TITLE' #sys.argv[9] # 樓層字串的圖層
    col_layer = 'S-TEXTC' #sys.argv[10] # col的圖層
    block_layer = '0' #sys.argv[8] # 圖框的圖層
    explode = 0 #sys.argv[12] # 需不需要提前炸圖塊

    # 在col裡面自訂圖層
    text_layer = 'S-TEXT' #sys.argv[6] # 文字的圖層
    line_layer = 'S-STUD' #sys.argv[7] # 線的圖層

    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()

    plan_file_count = plan_filename.count(',') + 1
    col_file_count = col_filename.count(',') + 1
    res_plan = [None] * plan_file_count
    res_col = [None] * col_file_count
    set_plan = set()
    dic_plan = {}
    set_col = set()
    dic_col = {}

    for i in range(plan_file_count):
        res_plan[i] = pool.apply_async(read_plan, (plan_filename.split(',')[i], floor_layer, col_layer, block_layer, plan_file, explode))

    for i in range(col_file_count):
        res_col[i] = pool.apply_async(read_col, (col_filename.split(',')[i], text_layer, line_layer, col_file, explode))

    for i in range(plan_file_count):
        final_plan = res_plan[i].get()
        set_plan = set_plan | final_plan[0]
        if plan_file_count == 1 and col_file_count == 1:
            dic_plan = final_plan[1]

    for i in range(col_file_count):
        final_col = res_col[i].get()
        set_col = set_col | final_col[0]
        if plan_file_count == 1 and col_file_count == 1:
            dic_col = final_col[1]

    drawing = 0
    if plan_file_count == 1 and col_file_count == 1:
        drawing = 1

    plan_result = write_plan(plan_filename, plan_new_filename, set_plan, set_col, dic_plan, result_file, date, drawing)
    col_result = write_col(col_filename, col_new_filename, set_plan, set_col, dic_col, result_file, date, drawing)

    end = time.time()
    # write_result_log(excel_file, task_name, plan_result[0], plan_result[1], beam_result[0], beam_result[1], f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')
    # write_result_log(excel_file,'','','','','','',time.strftime("%Y-%m-%d %H:%M", time.localtime()),'')
    # write_result_log(excel_file,task_name,'plan_result','col_result',f'{round(end - start, 2)}(s)', time.strftime("%Y-%m-%d %H:%M", time.localtime()), 'none')