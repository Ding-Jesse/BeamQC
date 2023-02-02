from __future__ import annotations
import time
import pythoncom
import win32com.client
import re
import os
import save_temp_file
from math import sqrt,pow
from rebar import RebarInfo
from plan_to_beam import turn_floor_to_float, turn_floor_to_string, turn_floor_to_list, floor_exist, vtFloat, error
from column import Column
def read_column_cad(column_filename,layer_config):
    text_layer = layer_config['text_layer']
    line_layer = layer_config['line_layer']
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
            error(f'read_beam error in step 1: {e}, error_count = {error_count}.')

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_column = wincad_column.Documents.Open(column_filename)
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 2: {e}, error_count = {error_count}.')

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_column = doc_column.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 3: {e}, error_count = {error_count}.')

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_column.Close(SaveChanges=False)
        except:
            pass
        return False
    flag = 0
    # Step 4. 解鎖所有圖層 -> 不然不能刪東西
    while not flag and error_count <= 10:
        try:
            layer_count = doc_column.Layers.count

            for x in range(layer_count):
                layer = doc_column.Layers.Item(x)
                layer.Lock = False
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            msp_column  = doc_column.Modelspace
            error(f'read_col error in step 4: {e}, error_count = {error_count}.')
    # progress('柱配筋圖讀取進度 4/10', progress_file)

    # Step 5. 遍歷所有物件 -> 炸圖塊  
    flag = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_col.Count
            layer_list = [text_layer, line_layer]
            for object in msp_col:
                count += 1
                if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:
                    object.Explode()
                if object.Layer not in layer_list:
                    object.Delete()
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            try:
                msp_col = doc_column.Modelspace
            except:
                pass
            error(f'read_col error in step 5: {e}, error_count = {error_count}.')
    # progress('柱配筋圖讀取進度 5/10', progress_file)

    # Step 6. 重新匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_column = doc_column.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_col error in step 6: {e}, error_count = {error_count}.')
    return msp_column,doc_column

def sort_col_cad(msp_column,layer_config,temp_file):
    text_layer = layer_config['text_layer']
    line_layer = layer_config['line_layer']
    rebar_text_layer = layer_config['rebar_text_layer']
    rebar_layer = layer_config['rebar_text_layer']
    tie_layer = layer_config['tie_layer']
    tie_text_layer = layer_config['tie_text_layer']
    coor_to_floor_set = set() # set(coor, floor)
    coor_to_col_set = set() # set(coor, col)
    coor_to_size_set = set() # set(coor, size)
    coor_to_floor_line_list = [] # (橫線y座標, start, end)
    coor_to_col_line_list = [] # (縱線x座標, start, end)
    coor_to_rebar_text_list = []
    coor_to_rebar_list = []
    coor_to_tie_list = []
    flag = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_column.Count
            # progress(f'柱配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
            for object in msp_column:
                count += 1
                # if count % 1000 == 0:
                #     progress(f'柱配筋圖已讀取{count}/{total}個物件', progress_file)
                if object.Layer in rebar_text_layer and object.ObjectName == "AcDbText":
                    if re.match(r'\d+.#\d+',object.TextString):
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        coor_to_rebar_text_list.append(((coor1,coor2),object.TextString))
                if object.Layer in coor_to_rebar_list:
                    if object.ObjectName == "AcDbCircle":
                        pass
                    if object.ObjectName == ""
                    coor_to_rebar_list
                if object.Layer in text_layer and object.ObjectName == "AcDbText": 
                    if object.TextString[0] == 'C' and len(object.TextString) <= 7:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        coor_to_col_set.add(((coor1, coor2), object.TextString))

                    elif 'x' in object.TextString or 'X' in object.TextString:
                        size = object.TextString.replace('X', 'x')
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                        coor_to_size_set.add(((coor1, coor2), size))
                    elif ('F' in object.TextString or 'B' in object.TextString or 'R' in object.TextString) and 'O' not in object.TextString: # 可能有樓層
                        floor = object.TextString
                        if '_' in floor: # 可能有B_6F表示B棟的6F
                            floor = floor.split('_')[1]
                        if turn_floor_to_float(floor):
                            coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                            coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                            floor = turn_floor_to_float(floor)
                            floor = turn_floor_to_string(floor)
                            coor_to_floor_set.add(((coor1, coor2), floor))
                
                elif object.Layer in line_layer:
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
            error_count += 1
            time.sleep(5)
            error(f'read_col error in step 7: {e}, error_count = {error_count}.')
        save_temp_file.save_pkl({'coor_to_col_set':coor_to_col_set,
                        'coor_to_size_set':coor_to_size_set,
                        'coor_to_floor_set': coor_to_floor_set,
                        'coor_to_col_line_list':coor_to_col_line_list,
                        },temp_file)

def cal_column_rebar(data={},output_folder = '',project_name = ''):
    output_txt =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar.txt')
    output_txt_2 =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar_floor.txt')
    output_excel = ''
    if not data:
        return
    coor_to_col_set = data['coor_to_col_set']
    coor_to_size_set = data['coor_to_size_set']
    coor_to_floor_set = data['coor_to_floor_set']
    coor_to_col_line_list = data['coor_to_col_line_list']

def concat_grid_line(line_list:list,start_line:list,overlap:function):
    while True:
        temp_coor = (start_line[1],start_line[2])
        temp_line_list = [l for l in line_list if l[0] == start_line[0] and overlap(l,start_line)]
        if len(temp_line_list) == 0:
            break
        new_line_top = max(temp_line_list,key = lambda l:l[2])[2]
        new_line_bot = min(temp_line_list,key = lambda l:l[1])[1]
        if start_line[2] == new_line_top and start_line[1] == new_line_bot:
            break
        start_line[2] = new_line_top
        start_line[1] = new_line_bot
    return start_line
def concat_col_to_grid(coor_to_col_set:set,coor_to_col_line_list:list):
    def _overlap(l1,l2):
        return (l2[1] - l1[2])*(l2[2] - l1[1]) <= 0
    # Step 8. 完成col_to_line_set 格式:(col, left, right, up)
    # coor_to_col_set:((coor1,coor2),string)
    # coor_to_col_line_list:[(coor1[0], min(coor1[1], coor2[1]), max(coor1[1], coor2[1]))] y向格線
    
    for element in coor_to_col_set:
        coor = element[0]
        col = element[1]
        new_coor_to_col_line_list = []
        left_temp_list = [l for l in coor_to_col_line_list if l[1] <= coor[0][1] and coor[0][1] <= l[2] and l[0] <= coor[0][0]]
        right_temp_list = [l for l in coor_to_col_line_list if l[1] <= coor[0][1] and coor[0][1] <= l[2] and l[0] >= coor[0][0]]
        left_closet_line = [0,float("inf"),float("-inf")]
        right_closet_line = [0,float("inf"),float("-inf")]
        if len(left_temp_list) > 0:
            left_closet_line = list(min(left_temp_list,key = lambda l:abs(coor[0][0] - l[0])))
            left_closet_line = concat_grid_line(coor_to_col_line_list,left_closet_line,_overlap)
        if len(right_temp_list) > 0:
            right_closet_line = list(min(right_temp_list,key = lambda l:abs(coor[0][0] - l[0])))
            right_closet_line = concat_grid_line(coor_to_col_line_list,right_closet_line,_overlap)
        top = max(left_closet_line[2],right_closet_line[2])
        bot = min(left_closet_line[1],right_closet_line[1])
        new_coor_to_col_line_list.append((col,(left_closet_line[0],right_closet_line[0],bot,top)))
        return left_closet_line,right_closet_line 
        # for vert_grid_line in coor_to_col_line_list: # (縱線x座標, start, end)
        #     # line_x = vert_grid_line[0]
        #     # line_top_y = vert_grid_line[1]
        #     # line_bot_y = vert_grid_line[2]
            
        #     if y[1] <= coor[0][1] <= y[2]: # 先看y座標有沒有被夾住
        #         new_coor_to_col_line_list.append(y)
        # tmp_set = set(new_coor_to_col_line_list)
        # new_coor_to_col_line_list = list(tmp_set)
        # new_coor_to_col_line_list.sort(key = lambda x: x[0])
        # for y in range(len(new_coor_to_col_line_list)): # 再看x座標被哪兩條線夾住
        #     if new_coor_to_col_line_list[y][0] < coor[0][0] < new_coor_to_col_line_list[y+1][0]:
        #         col_to_line_set.add((col, new_coor_to_col_line_list[y][0], new_coor_to_col_line_list[y+1][0], coor[1][1]))
def concat_floor_to_grid(coor_to_floor_set:set,coor_to_floor_line_list:list):
    def _overlap(l1,l2):
        return (l2[1] - l1[2])*(l2[2] - l1[1]) <= 0
    for element in coor_to_floor_set:
        coor = element[0]
        floor = element[1]
        new_coor_to_floor_line_list = []
        top_temp_list = [l for l in coor_to_floor_line_list if l[1] <= coor[0][1] and coor[0][1] <= l[2] and l[0] >= coor[0][0]]
        bot_temp_list = [l for l in coor_to_floor_line_list if l[1] <= coor[0][1] and coor[0][1] <= l[2] and l[0] <= coor[0][0]]
        top_closet_line = [0,float("inf"),float("-inf")]
        bot_closet_line = [0,float("inf"),float("-inf")]
        if len(top_temp_list) > 0:
            top_closet_line = list(min(top_temp_list,key = lambda l:abs(coor[0][0] - l[0])))
            top_closet_line = concat_grid_line(coor_to_floor_line_list,top_closet_line,_overlap)
        if len(bot_temp_list) > 0:
            bot_closet_line = list(min(bot_temp_list,key = lambda l:abs(coor[0][0] - l[0])))
            bot_closet_line = concat_grid_line(coor_to_floor_line_list,bot_closet_line,_overlap)
        right = max(top_closet_line[2],bot_closet_line[2])
        left = min(top_closet_line[1],bot_closet_line[1])
        new_coor_to_floor_line_list.append((floor,(left,right,bot_closet_line[0],top_closet_line[0])))

def concat_name_to_col_floor(coor_to_size_set:set,new_coor_to_col_line_list:list,new_coor_to_floor_line_list:list):
    def _ingrid(size_coor,grid_coor):
        pt_x = size_coor[0]
        pt_y = size_coor[1]
        if len(grid_coor) == 0:return False
        if (pt_x - grid_coor[1][0])*(pt_x - grid_coor[1][1])<0 and (pt_y - grid_coor[1][2])*(pt_y - grid_coor[1][3])<0:
            return True
        return False
    output_column_list:list[Column]
    output_column_list = []
    for element in coor_to_size_set:
        new_column = Column()
        coor = element[0]
        size = element[1]
        new_column.size = size
        col = [c for c in new_coor_to_col_line_list if _ingrid(size_coor=coor,grid_coor=c[1])]
        floor = [f for f in new_coor_to_floor_line_list if _ingrid(size_coor=coor,grid_coor=f[1])]
        if len(col) > 1:
            print(f'{size}:{coor} => {list(map(lambda c:c[0],col))}')
        if len(floor) > 1:
            print(f'{size}:{coor} => {list(map(lambda c:c[0],floor))}')
        if len(col) > 0:
            new_column.serial = col[0][0]
        if len(floor) > 0:
            new_column.floor = floor[0][0]
        output_column_list.append(new_column)
    return output_column_list
    pass
if __name__ == '__main__':
    # coor_to_col_set = set()
    # coor_to_col_set.add((((0,0),(10,10)),"C1"))
    # coor_to_col_line_list = [(-5,0,10),(-5,0,10),(-5,5,15),(-5,15,20),(-5,18,25),(-5,-5,30),(-5,-5,30),(10,0,10)]
    # print(temp(coor_to_col_set,coor_to_col_line_list))
    temp = [[1,2],[3,4]]
    print(f'{list(map(lambda t:t[0],temp))}')