from gzip import READ
from multiprocessing.spawn import prepare
from tkinter import HIDDEN
from numpy import object_
import win32com.client
import pythoncom
import re
import time
import multiprocessing
import os

start_time = time.time()

def turn_floor_to_int(floor): # turn string to int
    if floor == 'FB': # FB 直接變-1000層
        floor = str(-1000)

    if floor == 'PRF' or floor == 'PR': # PRF 直接變2000層
        floor = str(2000)

    if 'F' in floor: # 有F要把F拔掉
        floor = floor.replace("F", "")

    if 'B' in floor: # 有B直接變負整數
        floor = str(-int(floor.replace("B", "")))
    
    if 'R' in floor: # 有R直接+1000
        floor = str(int(floor.replace("R", "")) + 1000)
    
    floor = int(floor)
    return floor

def turn_floor_to_string(floor): # turn int to string
    if floor == -1000:
        floor = 'FBF' # 因為beam的部分字尾非F會自動補F，所以在diff的時候要一致

    elif floor > -1000 and floor < 0:
        floor = f'B{-floor}F'

    elif floor > 0 and floor < 1000:
        floor = f'{floor}F'

    elif floor > 1000 and floor < 2000:
        floor = f'R{floor - 1000}F'

    elif floor == 2000:
        floor = 'PRF'

    else:
        return False

    return floor

def floor_exist(i, Bmax, Fmax, Rmax): # 判斷是否為空號，例如B2F-PRF會從-2跑到2000，但顯然區間裡面的值不可能都合法
    if i == -1000: 
        return True
    
    elif i >= Bmax and i < 0: 
        return True
    
    elif i > 0 and i <= Fmax: 
        return True
    
    elif i > 1000 and i <= Rmax: 
        return True

    return False

def vtFloat(list): #要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)

def read_plan(plan_filename, floor_layer, col_layer, block_layer):
    flag = 0
    while not flag:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 1, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 2, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 3, {e}') # 報錯不要緊張，等一下就好，debug用
    
    time.sleep(5)
    
    
    coor_to_floor_set = set() # set (字串的coor, floor)
    coor_to_col_set = set() # set (coor, col)
    block_coor_list = [] # 存取方框最左下角的點座標
    Bmax = 0 # 地下最深到幾層(不包括FB不包括FB)
    Fmax = 0 # 正常樓最高到幾層
    Rmax = 0 # R開頭最高到幾層(不包括PRF)

    flag = 0
    while not flag:
        try:
            for object in msp_plan:
                if object.Layer == floor_layer and object.ObjectName == "AcDbText" and '(' in object.TextString and object.InsertionPoint[1] >= 0:
                    floor = object.TextString
                    floor = re.search('\(([^)]+)', floor).group(1) #取括號內的樓層數
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2)) #不取概數的話後面抓座標會出問題，例如兩個樓層在同一格
                    coor_to_floor_set.add((coor, floor))
                    
                if object.Layer in col_layer and (object.ObjectName == "AcDbText" or object.ObjectName == "AcDbMLeader") and object.GetBoundingBox()[0][1] >= 0 \
                        and (object.TextString[0] == 'G' or object.TextString[0] == 'B' or object.TextString[0] == 'g' or object.TextString[0] == 'b' or object.TextString[0:1] == 'cb' \
                        or object.TextString[0:1] == 'CB' or object.TextString[0:1] == 'CG'):
                    beam = object.TextString
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    if '(' in beam:
                        beam = beam.split('(')[0] # 取括號前內容即可
                    comma = beam.count(',')
                    for i in range(comma + 1):
                        coor_to_beam_set.add(((coor1, coor2), beam.split(',')[i]))
                
                if object.Layer in beam_layer and object.ObjectName == "AcDbMText" and object.GetBoundingBox()[0][1] >= 0:
                    # 這個case很特別，因為圖的排版要好看?所以格式頗特別，目前是用分號、左括號、反斜線來切割
                    beam = object.TextString
                    semicolon = beam.count(';')
                    for i in range(semicolon + 1):
                        s = beam.split(';')[i]
                        if s[0] == 'G' or s[0] == 'B' or s[0] == 'g' or s[0] == 'b' or s[0:1] == 'cb' or s[0:1] == 'CB' or s[0:1] == 'CG':
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
                    comma = beam.count(',')
                    for i in range(comma + 1):
                        coor_to_beam_set.add(((coor1, coor2), beam.split(',')[i]))
                
                if object.Layer == block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"): # 取大框框的左下角點座標
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    block_coor_list.append((coor1, coor2)) # 格式:((0.0, 0.0), (14275.54, 10824.61))

            flag = 1

        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 4, {e}') # 報錯不要緊張，等一下就好，debug用
    
    floor_to_coor_set = set() # set (floor, block左下角和右上角的coor)

    # 完成 floor_to_coor_set
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

    for y in floor_to_coor_set: # 算出Bmax, Fmax, Rmax
        floor = y[0]
        tmp_floor_list = []
        if '-' in floor or '~' in floor:
            if '-' in floor:
                start = floor.split('-')[0]
                end = floor.split('-')[1]
            else:
                start = floor.split('~')[0]
                end = floor.split('~')[1]
            tmp_floor_list.append(turn_floor_to_int(start))
            tmp_floor_list.append(turn_floor_to_int(end))
        else:
            comma = floor.count(',')
            for i in range(comma + 1):
                tmp_floor_list.append(turn_floor_to_int(floor.split(',')[i]))

        for x in tmp_floor_list:
            if x < 0 and x < Bmax and x != -1000:
                Bmax = x
            elif x > 0 and x < 1000 and x > Fmax:
                Fmax = x
            elif x > 1000 and x != 2000:
                Rmax = x

    set_plan = set() # set元素為 (樓層, 梁柱名稱)
    dic_plan = {} # 透過(floor, beam)去找字串座標

    # 完成 set_plan 以及 dic_plan
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
        
        if '-' in floor or '~' in floor: # 有區間
            if '-' in floor:
                start = floor.split('-')[0]
                end = floor.split('-')[1]
            else:
                start = floor.split('~')[0]
                end = floor.split('~')[1]
            
            start = turn_floor_to_int(start)
            end = turn_floor_to_int(end)

            if start > end:
                tmp = start
                start = end
                end = tmp

            for i in range(start, end + 1):
                if floor_exist(i, Bmax, Fmax, Rmax):
                    set_plan.add((turn_floor_to_string(i), beam_name))
                    dic_plan[(turn_floor_to_string(i), beam_name)] = full_coor
    
        elif ',' in floor: # 有逗號
            comma = floor.count(',')
            for i in range(comma + 1):
                new_floor = floor.split(',')[i]
                if new_floor[-1] != 'F': # 字尾不是F要加F
                    set_plan.add((f'{new_floor}F', beam_name))
                    dic_plan[(f'{new_floor}F', beam_name)] = full_coor
                else:
                    set_plan.add((new_floor, beam_name))
                    dic_plan[(new_floor, beam_name)] = full_coor
        
        else:
            if floor[-1] != 'F': # 字尾不是F要加F
                set_plan.add((f'{floor}F', beam_name))
                dic_plan[(f'{floor}F', beam_name)] = full_coor
            else:
                set_plan.add((floor, beam_name))
                dic_plan[(floor, beam_name)] = full_coor

    doc_plan.Close(SaveChanges=False)

    # plan.txt單純debug用，不想多新增檔案可以註解掉
    f = open("plan.txt", "w")
    f.write("in plan: \n")
    l = list(set_plan)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()

    return (set_plan, dic_plan)

def read_col(col_filename, text_layer):
    flag = 0
    while not flag:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 1, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 2, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 3, {e}') # 報錯不要緊張，等一下就好，debug用
    
    time.sleep(5)
    floor_to_beam_set = set() # set(floor, beam, coor)
    beam_dic = {}
    Bmax = 0
    Fmax = 0
    Rmax = 0
    set_beam = set()
    # 完成 set_beam
    flag = 0
    while not flag:
        try:
            for object in msp_beam:
                if object.Layer == text_layer and object.ObjectName == "AcDbText" and ' ' in object.TextString:
                    pre_beam = (object.TextString.split(' ')[1]).split('(')[0] # 把括號以後的東西拔掉
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    comma = pre_beam.count(',')
                    for i in range(comma + 1):
                        beam = pre_beam.split(',')[i]
                        if beam[0] == 'G' or beam[0] == 'B' or beam[0] == 'g' or beam[0] == 'b' or beam[0:1] == 'CB' or beam[0:1] == 'CG' or beam[0:1] == 'cb':
                            floor = object.TextString.split(' ')[0]
                            floor_to_beam_set.add((floor, beam, (coor1, coor2)))
                            tmp_floor_list = []
                            if '-' in floor or '~' in floor:
                                if '-' in floor:
                                    start = floor.split('-')[0]
                                    end = floor.split('-')[1]
                                else:
                                    start = floor.split('~')[0]
                                    end = floor.split('~')[1]
                                tmp_floor_list.append(turn_floor_to_int(start))
                                tmp_floor_list.append(turn_floor_to_int(end))
                            else:
                                comma = floor.count(',')
                                for i in range(comma + 1):
                                    tmp_floor_list.append(turn_floor_to_int(floor.split(',')[i]))

                            for x in tmp_floor_list:
                                if x < 0 and x < Bmax and x != -1000:
                                    Bmax = x
                                elif x > 0 and x < 1000 and x > Fmax:
                                    Fmax = x
                                elif x > 1000 and x != 2000:
                                    Rmax = x
                flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 4, {e}') # 報錯不要緊張，等一下就好，debug用
        
        for x in floor_to_beam_set:
            floor = x[0]
            beam = x[1]
            coor = x[2]
            if '-' in floor or '~' in floor: # 有區間
                if '-' in floor:
                    start = floor.split('-')[0]
                    end = floor.split('-')[1]
                else:
                    start = floor.split('~')[0]
                    end = floor.split('~')[1]

                start = turn_floor_to_int(start)
                end = turn_floor_to_int(end)

                if start > end:
                    tmp = start
                    start = end
                    end = tmp
                
                for i in range(start, end + 1): # 把數字變回樓層，然後F要加回來
                    if floor_exist(i, Bmax, Fmax, Rmax):
                        set_beam.add((turn_floor_to_string(i), beam))
                        beam_dic[(turn_floor_to_string(i), beam)] = coor
            
            elif ',' in floor: # 有逗號
                comma = floor.count(',')
                for i in range(comma + 1):
                    new_floor = floor.split(',')[i]
                    if new_floor[-1] != 'F': # 字尾不是F要加F
                        set_beam.add((f'{new_floor}F', beam))
                        beam_dic[(f'{new_floor}F', beam)] = coor
                    else:
                        set_beam.add((new_floor, beam))
                        beam_dic[(new_floor, beam)] = coor
            else:
                if floor[-1] != 'F': # 字尾不是F要加F
                    set_beam.add((f'{floor}F', beam))
                    beam_dic[(f'{floor}F', beam)] = coor
                else:
                    set_beam.add((floor, beam))
                    beam_dic[(floor, beam)] = coor

    doc_beam.Close(SaveChanges=False)

    # beam.txt單純debug用，不想多新增檔案可以註解掉
    f = open("beam.txt", "w")
    f.write("in beam: \n")
    l = list(set_beam)
    l.sort()
    for x in l: 
        f.write(f'{x}\n')
    f.close()
    
    return (set_beam, beam_dic)

if __name__=='__main__':

    plan_filename = "C:\\Users\\Vince\\Desktop\\test\\task2\\XS-PLAN2.dwg"
    beam_filename = "C:\\Users\\Vince\\Desktop\\test\\task2\\XS-BEAM2.dwg"
    plan_new_filename = "C:\\Users\\Vince\\Desktop\\test\\task2\\XS-PLAN2_new.dwg"
    beam_new_filename = "C:\\Users\\Vince\\Desktop\\test\\task2\\XS-BEAM2_new.dwg"

    final_filename = 'task2'
    date = '2022-0712'
    # 在plan裡面自訂圖層
    floor_layer = "S-TITLE" # 樓層字串的圖層
    beam_layer = ["S-TEXTG", "S-TEXTB"] # beam的圖層，因為有兩個以上，所以用list來存
    block_layer = "DEFPOINTS" # 框框的圖層

    # 在beam裡面自訂圖層
    text_layer = "S-RC"

    multiprocessing.freeze_support()
    pool = multiprocessing.Pool()
    res_plan = pool.apply_async(read_plan, (plan_filename, floor_layer, beam_layer, block_layer))
    res_beam = pool.apply_async(read_beam, (beam_filename, text_layer))
    final_plan = res_plan.get()
    final_beam = res_beam.get()
    set_plan = final_plan[0]
    dic_plan = final_plan[1]
    set_beam = final_beam[0]
    dic_beam = final_beam[1]

    set1 = set_plan - set_beam
    list1 = list(set1)
    list1.sort()

    set2 = set_beam - set_plan
    list2 = list(set2)
    list2.sort()

    # 完成 in plan but not in beam 的部分並在圖上mark有問題的部分
    f_big = open(f"big{final_filename}.txt", "w")
    f_sml = open(f"sml{final_filename}.txt", "w")

    f_big.write("in plan but not in beam: \n")
    f_sml.write("in plan but not in beam: \n")

    flag = 0
    while not flag:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 1, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 2, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Plan Error 3, {e}') # 報錯不要緊張，等一下就好，debug用
    
    time.sleep(5)

    # 設定mark的圖層
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
            print(f'Plan Error 4, {e}') # 報錯不要緊張，等一下就好，debug用

    for x in list1: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':        
            f_big.write(f'({x[0]}, {x[1]})\n')
        else:
            f_sml.write(f'({x[0]}, {x[1]})\n')
        
        coor = dic_plan[x]
        list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
        points = vtFloat(list)
        pointobj = msp_plan.AddPolyline(points)
        for i in range(4):
            pointobj.SetWidth(i, 10, 10)
    
    doc_plan.SaveAs(plan_new_filename)
    doc_plan.Close(SaveChanges=True)
    
    # 完成 in beam but not in plan 的部分並在圖上mark有問題的部分

    f_big.write("in beam but not in plan: \n")
    f_sml.write("in beam but not in plan: \n")

    flag = 0
    while not flag:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 1, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 2, {e}') # 報錯不要緊張，等一下就好，debug用
    flag = 0
    while not flag:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            time.sleep(5)
            print(f'Beam Error 3, {e}') # 報錯不要緊張，等一下就好，debug用
    
    time.sleep(5)

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
            print(f'Beam Error 4, {e}') # 報錯不要緊張，等一下就好，debug用

    for x in list2: 
        if x[1][0] == 'B' or x[1][0] == 'C' or x[1][0] == 'G':
            f_big.write(f'{x}\n')
        else:
            f_sml.write(f'{x}\n')
        
        coor = dic_beam[x]
        list = [coor[0][0] - 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[0][1] - 20, 0, coor[1][0] + 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[1][1] + 20, 0, coor[0][0] - 20, coor[0][1] - 20, 0]
        points = vtFloat(list)
        pointobj = msp_beam.AddPolyline(points)
        for i in range(4):
            pointobj.SetWidth(i, 10, 10)

    doc_beam.SaveAs(beam_new_filename)
    doc_beam.Close(SaveChanges=True)
    
    f_big.close()
    f_sml.close()
    
    end_time = time.time()
    print(f'I spend {end_time - start_time} seconds. ')