from __future__ import annotations
import time
import pythoncom
import win32com.client
import re
import os
import save_temp_file
import pandas as pd
from math import sqrt,pow
from item.beam import Beam
from item.rebar import RebarInfo
from item.floor import Floor,read_parameter_df,summary_floor_rebar
from beam_scan import create_beam_scan,beam_check
from main import OutputExcel
error_file = './result/error_log.txt' # error_log.txt的路徑
def vtFloat(l): #要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)
def vtPnt(x, y, z=0):
    """座標點轉化爲浮點數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

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

def read_beam_cad(beam_filename, progress_file):
    error_count = 0
    pythoncom.CoInitialize()
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
    progress('梁配筋圖讀取進度 1/15', progress_file)

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
    progress('梁配筋圖讀取進度 2/15', progress_file)

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
    progress('梁配筋圖讀取進度 3/15', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False

    return msp_beam,doc_beam
    # # Step 4 解鎖所有圖層 -> 不然不能刪東西
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         layer_count = doc_beam.Layers.count
    #         for x in range(layer_count):
    #             layer = doc_beam.Layers.Item(x)
    #             layer.Lock = False
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 4: {e}, error_count = {error_count}.')
    # progress('梁配筋圖讀取進度 4/15', progress_file)

    # # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間  
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         count = 0
    #         total = msp_beam.Count
    #         progress(f'正在炸梁配筋圖的圖塊及篩選判斷用的物件，梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
    #         layer_list = [rebar_layer, rebar_data_layer]
    #         for object in msp_beam:
    #             count += 1
    #             if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:
    #                 object.Explode()
    #             if object.Layer not in layer_list:
    #                 object.Delete()
    #             if count % 1000 == 0:
    #                 progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
    #         flag = 1
        
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 5: {e}, error_count = {error_count}.')
    #         msp_beam = doc_beam.Modelspace
    # progress('梁配筋圖讀取進度 5/15', progress_file)

    # # Step 6. 重新匯入modelspace
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         msp_beam = doc_beam.Modelspace
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 6: {e}, error_count = {error_count}.')
    # progress('梁配筋圖讀取進度 6/15', progress_file)
    
    # Step 7. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    # progress('正在遍歷梁配筋圖上的物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候', progress_file)
    
def sort_beam_cad(msp_beam,layer_config:dict,entity_config:dict, progress_file='',temp_file=''):

    rebar_layer = layer_config['rebar_layer']
    rebar_data_layer = layer_config['rebar_data_layer']
    tie_text_layer = layer_config['tie_text_layer']
    block_layer = layer_config['block_layer']
    beam_text_layer = layer_config['beam_text_layer']
    bounding_block_layer = layer_config['bounding_block_layer']
    print(f'temp_file:{temp_file}')
    coor_to_rebar_list = [] # (頭座標，尾座標，長度)
    coor_to_bend_rebar_list = [] # (直的端點，橫的端點，長度)
    coor_to_data_list = [] # (字串，座標)
    coor_to_arrow_dic = {} # 尖點座標 -> 文字連接處座標
    coor_to_tie_list = [] # (下座標，上座標，長度) 
    coor_to_tie_text_list = [] # (字串，座標)
    coor_to_block_list = [] # ((左下，右上), rebar_length_dic, tie_count_dic)
    coor_to_beam_list = [] # (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
    coor_to_bounding_block_list = [] #((左下，右上),beam_name, list of tie, tie_count_dic, list of rebar,rebar_length_dic)
    flag = 0
    error_count = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_beam.Count
            progress(f'梁配筋圖上共有{total}個物件，大約運行{int(total / 5500)}分鐘，請耐心等候', progress_file)
            for object in msp_beam:
                count += 1
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
                # 抓鋼筋的字的座標
                if object.Layer == rebar_data_layer and object.ObjectName in entity_config['rebar_data_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2))
                    coor_to_data_list.append((object.TextString, coor))
                # 抓箭頭座標
                elif object.Layer == rebar_data_layer and object.ObjectName in entity_config['rebar_data_leader_layer']:
                    # object.Coordinates 有九個參數 -> 箭頭尖點座標，直角的座標，文字接出去的座標，都有x, y, z
                    coor_to_arrow_dic[(round(object.Coordinates[0], 2), round(object.Coordinates[1], 2))] = (round(object.Coordinates[6], 2), round(object.Coordinates[7], 2))
                # 抓鋼筋本人和箍筋本人
                elif object.Layer == rebar_layer:
                    # object.Coordinates 橫的和直的有四個參數 -> 兩端點的座標，都只有x, y; 彎的有八個參數 -> 直的端點，直的轉角，橫的轉角，橫的端點
                    if (object.ObjectName == 'AcDbPolyline'):
                        if round(object.Length, 4) > 4: # 太短的是分隔線 -> 不要
                            if len(object.Coordinates) == 4 and round(object.Coordinates[1], 2) == round(object.Coordinates[3], 2): # 橫的 -> y 一樣 -> 鋼筋
                                coor_to_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4))) 
                            elif len(object.Coordinates) == 4 and round(object.Coordinates[0], 2) == round(object.Coordinates[2], 2): # 直的 -> x 一樣 -> 箍筋
                                coor_to_tie_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4))) 
                            elif len(object.Coordinates) == 8: # 彎的 -> 直的端點，橫的端點
                                coor_to_bend_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[6], 2), round(object.Coordinates[7], 2)), round(object.Length, 4))) 
                    elif (object.ObjectName == 'AcDbLine'):
                        if round(object.startPoint[1], 2) == round(object.endPoint[1], 2): # 橫的 -> y 一樣 -> 鋼筋
                            coor_to_rebar_list.append(((round(object.startPoint[0], 2), round(object.startPoint[1], 2)), (round(object.endPoint[0], 2), round(object.endPoint[1], 2)), round(object.Length, 4))) 
                        elif round(object.startPoint[0], 2) == round(object.endPoint[0], 2): # 直的 -> x 一樣 -> 箍筋
                            coor_to_tie_list.append(((round(object.startPoint[0], 2), round(object.startPoint[1], 2)), (round(object.endPoint[0], 2), round(object.endPoint[1], 2)), round(object.Length, 4))) 
                # 抓箍筋文字座標
                elif object.Layer == tie_text_layer and object.ObjectName in entity_config['tie_text_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2))
                    coor_to_tie_text_list.append((object.TextString, coor))
                    continue
                # 抓圖框
                elif object.Layer == block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    coor_to_block_list.append(((coor1, coor2), {}, {}))
                # 抓圖框
                elif object.Layer == bounding_block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    coor_to_bounding_block_list.append(((coor1, coor2),"", [],{},[], {}))
                # 抓梁的字的座標
                elif object.Layer == beam_text_layer and object.ObjectName == 'AcDbText':
                    midpoint = (round((object.GetBoundingBox()[0][0] + object.GetBoundingBox()[1][0]) / 2, 2), round((object.GetBoundingBox()[0][1] + object.GetBoundingBox()[1][1]) / 2, 2))
                    coor_to_beam_list.append([object.TextString, midpoint, [], {},(),[], {}])# (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
                if object.Layer == tie_text_layer and object.ObjectName in entity_config['tie_text_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2))
                    coor_to_tie_text_list.append((object.TextString, coor))
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 7: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 7/15', progress_file)
    save_temp_file.save_pkl({'coor_to_data_list':coor_to_data_list,
                            'coor_to_arrow_dic':coor_to_arrow_dic,
                            'coor_to_rebar_list': coor_to_rebar_list,
                            'coor_to_bend_rebar_list':coor_to_bend_rebar_list,
                            'coor_to_tie_list':coor_to_tie_list,
                            'coor_to_tie_text_list':coor_to_tie_text_list ,
                            'coor_to_block_list':coor_to_block_list,
                            'coor_to_beam_list':coor_to_beam_list,
                            'coor_to_bounding_block_list':coor_to_bounding_block_list
                            },temp_file)

#整理箭頭與直線對應
def sort_arrow_line(coor_to_arrow_dic:dict,coor_to_rebar_list:list):
    start = time.time() 
    # #method 1
    # new_coor_to_arrow_dic = {}
    # for x in coor_to_arrow_dic: #此時的coor_to_arrow_dic為尖點座標->文字端坐標
    #     arrow_coor = x
    #     min_diff = 100
    #      # 先看y是不是最近，再看x有沒有被夾到
    #     min_head_coor = ''
    #     min_length = ''
    #     min_mid_coor = ''
    #     for y in coor_to_rebar_list: # (頭座標，尾座標，長度)
    #         head_coor = y[0]
    #         tail_coor = y[1]
    #         mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])#
    #         length = y[2]
    #         y_diff = abs(mid_coor[1] - arrow_coor[1])
    #         if y_diff < min_diff and (head_coor[0] - arrow_coor[0]) * (tail_coor[0] - arrow_coor[0]) <= 0:
    #             min_diff = y_diff
    #             min_head_coor = head_coor
    #             min_tail_coor = tail_coor
    #             min_length = length
    #             min_mid_coor = mid_coor
        
    #     if min_head_coor != '':
    #         new_coor_to_arrow_dic[x] = (coor_to_arrow_dic[x], min_length, min_mid_coor) # 新的coor_to_arrow_dic為尖點座標 -> (文字端坐標，鋼筋長度，鋼筋中點座標)
    #         # coor_to_rebar_list.remove((min_head_coor, min_tail_coor, min_length))
    # print(f'Method 1:{time.time() - start}')
    
    start = time.time()
    #method 2
    new_coor_to_arrow_dic = {}
    no_arrow_line_list = []
    min_diff = 50
    for i,rebar in enumerate(coor_to_rebar_list):
        if i == 199:
            print('1')
        head_coor = rebar[0]
        tail_coor = rebar[1]
        mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
        length = rebar[2]
        arrow_dict = {k: v for k, v in coor_to_arrow_dic.items() if (head_coor[0] - k[0]) * (tail_coor[0] - k[0]) <= 0}
        if arrow_dict:
            value_pair = min(arrow_dict.items(),key=lambda x:abs(mid_coor[1] - x[0][1]))
            if(abs(value_pair[0][1] - mid_coor[1])> min_diff):
                no_arrow_line_list.append(rebar)
                continue
            new_coor_to_arrow_dic.update({value_pair[0]:(value_pair[1],length,mid_coor)})
    print(f'Method 2:{time.time() - start}')
    return new_coor_to_arrow_dic,no_arrow_line_list
    # print(new_coor_to_arrow_dic_2 == new_coor_to_arrow_dic)
    # print(set(new_coor_to_arrow_dic_2.items()) - set(new_coor_to_arrow_dic.items()))
    # print(set(new_coor_to_arrow_dic.items()) - set(new_coor_to_arrow_dic_2.items()))

#整理箭頭與鋼筋文字對應
def sort_arrow_to_word(coor_to_arrow_dic:dict,coor_to_data_list:list):
    def _get_distance(pt1,pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    # start = time.time()
    # new_coor_to_arrow_dic = {}
    # head_to_data_dic = {} # 座標 -> (number, size)
    # tail_to_data_dic = {}
    # for x in coor_to_arrow_dic: # 新的coor_to_arrow_dic為尖點座標 -> (文字端坐標，鋼筋長度，鋼筋中點座標)
    #     if len(coor_to_arrow_dic[x]) == 3:
    #         arrow_coor = coor_to_arrow_dic[x][0]
    #         length = coor_to_arrow_dic[x][1]
    #         rebar_mid_coor = coor_to_arrow_dic[x][2]
    #         min_diff = 100
    #         min_data = ''
    #         min_data_coor = ''
    #         for y in coor_to_data_list: # for 鋼筋的 (字串，座標)
    #             data = y[0]
    #             data_coor = y[1]
    #             x_diff = abs(arrow_coor[0] - data_coor[0])
    #             y_diff = abs(arrow_coor[1] - data_coor[1])
    #             total = x_diff + y_diff
    #             if total < min_diff:
    #                 min_diff = total
    #                 min_data = data
    #                 min_data_coor = data_coor
    #         if min_data != '':
    #             if '-' in min_data:
    #                 number = min_data.split('-')[0]
    #                 size =  min_data.split('-')[1]    
    #                 new_coor_to_arrow_dic[x] = (arrow_coor, length, rebar_mid_coor, number, size, min_data_coor) # 新的coor_to_arrow_dic為尖點座標 -> (箭頭文字端坐標，鋼筋長度，鋼筋中點座標，數量，尺寸，文字座標)
    #                 head_to_data_dic[(rebar_mid_coor[0] - length / 2, rebar_mid_coor[1])] = (number, size)
    #                 tail_to_data_dic[(rebar_mid_coor[0] + length / 2, rebar_mid_coor[1])] = (number, size)
    #             else:
    #                 error(f"There are no '-' in {min_data}. ")
    # print(f'Method 1:{time.time() - start}')
    start = time.time()
    min_diff = 100
    new_coor_to_arrow_dic = {}
    head_to_data_dic = {} # 座標 -> (number, size)
    tail_to_data_dic = {}
    #method 2
    for arrow_head,arrow_data in coor_to_arrow_dic.items():
        rebar_data_temp = [r for r in coor_to_data_list if '@' not in r[0]]
        text,coor = min(rebar_data_temp,key=lambda rebar_text:_get_distance(arrow_data[0],rebar_text[1]))
        arrow_tail,rebar_length,line_coor = arrow_data
        if(abs(arrow_tail[1] - coor[1])> min_diff):
            print(f'{arrow_head} / {arrow_data } cant find pair arrow')
            continue
        rebar_data = list(arrow_data)
        if '-' not in text:
            print(f'{text} not satisfied rule')
            continue
        if '@' in text:
            print(f'{text} not satisfied rule')
            continue
        number = text.split('-')[0]
        size =  text.split('-')[1]    
        rebar_data.extend([number,size ,coor])
        new_coor_to_arrow_dic.update({arrow_head:(*rebar_data  ,)})
        head_to_data_dic.update({(line_coor[0] - rebar_length/2,line_coor[1]):{'number':number,'size':size}})
        tail_to_data_dic.update({(line_coor[0] + rebar_length/2,line_coor[1]):{'number':number,'size':size}})

    print(f'Method 2:{time.time() - start}')
    # print(new_coor_to_arrow_dic_2 == new_coor_to_arrow_dic)
    # print(set(new_coor_to_arrow_dic_2.items()) - set(new_coor_to_arrow_dic.items()))
    # print(set(new_coor_to_arrow_dic.items()) - set(new_coor_to_arrow_dic_2.items()))
    return new_coor_to_arrow_dic,head_to_data_dic,tail_to_data_dic
'''
def concat_no_arrow_line(no_arrow_line_list:list,head_to_data_dic:dict,tail_to_data_dic:dict,coor_to_bend_rebar_list:list):
    coor_to_rebar_list_straight_left = [] # (頭座標，尾座標，長度，number，size)
    coor_to_rebar_list_straight_right = []
    no_concat_line_list = []
    no_concat_bend_list = []
    for head_coor, tail_coor,line_length in no_arrow_line_list: # (頭座標，尾座標，長度)
        if tail_coor in head_to_data_dic: # 座標 -> (number, size)
            coor_to_rebar_list_straight_right.append((head_coor, tail_coor, line_length, head_to_data_dic[tail_coor]['number'], 
                                                        head_to_data_dic[tail_coor]['size']))
        elif head_coor in head_to_data_dic:
            coor_to_rebar_list_straight_right.append((head_coor, tail_coor, line_length, head_to_data_dic[head_coor]['number'], 
                                            head_to_data_dic[head_coor]['size']))
        elif head_coor in tail_to_data_dic :
            coor_to_rebar_list_straight_left.append((head_coor, tail_coor,line_length, tail_to_data_dic[head_coor]['number'], 
                                                    tail_to_data_dic[head_coor]['size']))
        elif tail_coor in tail_to_data_dic:
            coor_to_rebar_list_straight_left.append((head_coor, tail_coor,line_length, tail_to_data_dic[tail_coor]['number'], 
                                                    tail_to_data_dic[tail_coor]['size']))
        else:
            no_concat_line_list.append((head_coor, tail_coor,line_length))
            print(f'{head_coor}/{tail_coor} no concat')

    new_coor_to_bend_rebar_list = [] # 新的：(直的端點，橫的端點，長度，number，size)
    for vert_coor, horz_coor,line_length in coor_to_bend_rebar_list: # (直的端點，橫的端點，長度)
        if horz_coor in head_to_data_dic:
            new_coor_to_bend_rebar_list.append((vert_coor, horz_coor,line_length, head_to_data_dic[horz_coor]['number'], 
                                                head_to_data_dic[horz_coor]['size']))
        elif horz_coor in tail_to_data_dic:
            new_coor_to_bend_rebar_list.append((vert_coor, horz_coor,line_length, tail_to_data_dic[horz_coor]['number'], 
                                                tail_to_data_dic[horz_coor]['size']))
        else:
            no_concat_bend_list.append((vert_coor, horz_coor,line_length))
            print(f'bend {vert_coor}/{horz_coor} no concat')       

    return coor_to_rebar_list_straight_left,coor_to_rebar_list_straight_right,new_coor_to_bend_rebar_list,no_concat_line_list,no_concat_bend_list 
'''
def sort_noconcat_line(no_concat_line_list,head_to_data_dic:dict,tail_to_data_dic:dict):
    start = time.time()
    coor_to_rebar_list_straight = [] # (頭座標，尾座標，長度，number，size)
    def _overlap(l1,l2):
        if l1[1] == l2[0][1]:
            return (l2[0][0] - l1[0])*(l2[1][0] - l1[0]) <= 0
        return False
    def _cal_length(pt1,pt2):
        return sqrt((pt1[0]-pt2[0])**2 +(pt1[1]-pt2[1])**2)
    def _concat_line(line_list:list):
        for line in line_list[:]:
            head_coor = min([line[0],line[1]],key=lambda l: l[0])
            tail_coor = max([line[0],line[1]],key=lambda l: l[0])
            head_rebar = {}
            tail_rebar = {}
            overlap_line = {k: v for k, v in head_to_data_dic.items() if _overlap(k,(head_coor, tail_coor))}
            if len(overlap_line.keys()) > 0:
                value_key,value_items = min(overlap_line.items(),key=lambda x:x[0][0])
                tail_coor = value_key
                tail_rebar = value_items
            overlap_line = {k: v for k, v in tail_to_data_dic.items() if _overlap(k,(head_coor, tail_coor))}
            if len(overlap_line.keys()) > 0:
                value_key,value_items =max(overlap_line.items(),key=lambda x:x[0][0])
                head_coor =value_key
                head_rebar = value_items
            if (not head_rebar) and (not tail_rebar):
                print(f'{head_coor},{tail_coor} norebar')
                continue
            elif head_rebar==tail_rebar:
                coor_to_rebar_list_straight.append((head_coor, tail_coor,_cal_length(head_coor,tail_coor),head_rebar['number'],head_rebar['size']))
            elif head_rebar!=tail_rebar:
                if head_rebar and tail_rebar:
                    print(f'{head_coor},{tail_coor} head_rebar:{head_rebar} tail_rebar:{tail_rebar}')
                elif head_rebar:
                    coor_to_rebar_list_straight.append((head_coor, tail_coor,_cal_length(head_coor,tail_coor),head_rebar['number'],head_rebar['size']))
                elif tail_rebar:
                    coor_to_rebar_list_straight.append((head_coor, tail_coor,_cal_length(head_coor,tail_coor),tail_rebar['number'],tail_rebar['size']))
            head_to_data_dic.update({head_coor:{'number':coor_to_rebar_list_straight[-1][3],'size':coor_to_rebar_list_straight[-1][4]}})
            tail_to_data_dic.update({tail_coor:{'number':coor_to_rebar_list_straight[-1][3],'size':coor_to_rebar_list_straight[-1][4]}})
            line_list.remove(line)
            print(f'{head_coor},{tail_coor} rebar:{coor_to_rebar_list_straight[-1][3]}-{coor_to_rebar_list_straight[-1][4]}')
        
    while True:
        temp_count = len(no_concat_line_list)
        _concat_line(no_concat_line_list)
        if temp_count == len(no_concat_line_list) or len(no_concat_line_list) == 0:
            break
    print(f'Method 1:{time.time() - start}')
    return coor_to_rebar_list_straight

def sort_noconcat_bend(no_concat_bend_list:list,head_to_data_dic:dict,tail_to_data_dic:dict):
    start = time.time()
    coor_to_bend_rebar_list = []
    for bend in no_concat_bend_list:
        horz_coor = bend[1]
        vert_coor = bend[0]
        line_length = bend[2]
        overlap_line = {k: v for k, v in head_to_data_dic.items() if (k[0] <= horz_coor[0]) and (k[1] == horz_coor[1]) and (k[0] >= vert_coor[0])}
        if len(overlap_line.keys()) > 0:
            value_key,value_items = min(overlap_line.items(),key=lambda x:abs(x[0][0]-horz_coor[0]))
            coor_to_bend_rebar_list.append((vert_coor, value_key,line_length - abs(value_key[0] - horz_coor[0]),value_items['number'],value_items['size']))
            print(f'{horz_coor},{vert_coor} rebar:{coor_to_bend_rebar_list[-1][3]}-{coor_to_bend_rebar_list[-1][4]}')
            continue
        overlap_line = {k: v for k, v in tail_to_data_dic.items() if (k[0] >= horz_coor[0]) and (k[1] == horz_coor[1]) and (k[0] <= vert_coor[0]) }
        if len(overlap_line.keys()) > 0:
            value_key,value_items = max(overlap_line.items(),key=lambda x:abs(x[0][0]-horz_coor[0]))
            coor_to_bend_rebar_list.append((vert_coor, value_key,line_length - abs(value_key[0] - horz_coor[0]),value_items['number'],value_items['size']))
            print(f'{horz_coor},{vert_coor} rebar:{coor_to_bend_rebar_list[-1][3]}-{coor_to_bend_rebar_list[-1][4]}')
            continue
    print(f'Method 1:{time.time() - start}')
    return coor_to_bend_rebar_list

def sort_rebar_bend_line(rebar_bend_list:list,rebar_line_list:list):
    def _between_pt(pt1,pt2,pt):
        return (pt[0] - pt1[0])*(pt[0] - pt2[0]) < 0 and pt1[1] == pt2[1] == pt[1]
    def _outer_pt(start_pt,end_pt,pt):
        if start_pt[0] < end_pt[0]:
            return pt[0] > end_pt[0]
        if start_pt[0] > end_pt[0]:
            return pt[0] < end_pt[0]
        return False
    def _overline(start_pt,end_pt,line):
        if _between_pt(start_pt,end_pt,line[0]) and _outer_pt(start_pt,end_pt,line[1]):
            return True
        if _between_pt(start_pt,end_pt,line[1]) and _outer_pt(start_pt,end_pt,line[0]):
            return True
    for bend in rebar_bend_list[:]:
        vert_coor = bend[0]
        horz_coor = bend[1]
        bend_length = bend[2]
        rebar_size = bend[4]
        rebar_number = bend[3]

        end_pt = (vert_coor[0],horz_coor[1])
        start_pt = (horz_coor[0],horz_coor[1])
        concat_line = [l for l in rebar_line_list if _overline(start_pt,end_pt,(l[0],l[1])) and l[4] == rebar_size]
        if concat_line:
            closet_line = min(concat_line,key=lambda l : int(l[3]))
            new_number = int(rebar_number) - int(closet_line[3])
            if new_number > 0:
                rebar_bend_list.remove(bend)
                rebar_bend_list.append((vert_coor,horz_coor,bend_length,str(new_number),rebar_size))
                print(f'{horz_coor} {rebar_number}-{rebar_size} => {new_number}-{rebar_size}')

def count_tie(coor_to_tie_text_list:list,coor_to_block_list:list,coor_to_tie_list):
    def extract_tie(tie:str):
        new_tie =re.findall(r'\d*-\d*#\d@\d\d',tie)
        if len(new_tie) == 0:new_tie = re.findall(r'\d#\d@\d\d',tie)
        if len(new_tie) == 0:new_tie = re.findall(r'#\d@\d\d',tie)
        if len(new_tie) == 0:return tie
        return new_tie[0]
    tie_num =''
    tie_text = ''
    count = 1 
    size =''
    coor_sorted_tie_list =[]
    for tie,coor in coor_to_tie_text_list: # (字串，座標)
        tie = extract_tie(tie=tie)
        if '-' in tie:
            tie_num = tie.split('-')[0]
            tie_text = tie.split('-')[1]
            if tie_num.isdigit(): # 已經算好有幾根就直接用
                count = int(tie_num)
                size = tie_text.split('@')[0] # 用'-'和'@'來切
                if size.split('#')[0].isdigit():
                    count *= int(size.split('#')[0])
                    size = f"#{size.split('#')[1]}"
                for block in coor_to_block_list:
                    if inblock(block=block[0],pt=coor):
                        coor_sorted_tie_list.append((tie,coor,tie_num,count,size))
                        # print(f'pt:{coor} in block:{block[0]}')
                        # y[2] 是該格的tie_count_dic: size -> number
                        if size not in block[2]:
                            block[2][size] = count
                        else:
                            block[2][size] += count
                        break

        else: # 沒算好自己算
            if not '@' in tie or not '#' in tie:
                print(f'{tie} wrong format ex:#4@20')
                continue
            size = tie.split('@')[0] # 用'@'來切
            try:
                spacing = int(tie.split('@')[1])
            except Exception as ex:
                continue 
            assert spacing != 0,f'{coor} spacing is zero'

            tie_left_list = [(bottom,top,length) for bottom,top,length in coor_to_tie_list if ( bottom[0] < coor[0]) and (bottom[1] < coor[1]) and (top[1] > coor[1])]
            tie_right_list = [(bottom,top,length) for bottom,top,length in coor_to_tie_list if ( bottom[0] > coor[0]) and (bottom[1] < coor[1]) and (top[1] > coor[1])]
        # for bottom,top,length in coor_to_tie_list: # (下座標，上座標，長度) 
        #     if bottom[0] < coor[0] and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字左邊且diff最小且文字有被上下的y夾住
        #         left_diff = x[1][0] - y[0][0]
        #         min_left_coor = y[0]
        #     elif y[0][0] > x[1][0] and y[0][0] - x[1][0] < right_diff and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字右邊且diff最小且文字有被上下的y夾住
        #         right_diff = y[0][0] - x[1][0]
        #         min_right_coor = y[0]
            if not (tie_left_list and tie_right_list):
                print(f'{tie} {coor} no line bounded')
                continue

            left_tie = min(tie_left_list,key=lambda t:abs(t[0][0] - coor[0]))
            right_tie = min(tie_right_list,key=lambda t:abs(t[0][0] - coor[0]))

            
            count = int(abs(left_tie[0][0] - right_tie[0][0]) / spacing)
            if size.split('#')[0].isdigit():
                count *= int(size.split('#')[0])
                size = f"#{size.split('#')[1]}"
            for block in coor_to_block_list:
                if inblock(block=block[0],pt=coor):
                    coor_sorted_tie_list.append((tie,coor,tie_num,count,size))
                    # print(f'pt:{coor} in block:{block[0]}')
                    # y[2] 是該格的tie_count_dic: size -> number
                    if size not in block[2]:
                        block[2][size] = count
                    else:
                        block[2][size] += count
                    break
    return coor_sorted_tie_list

## 組合手動框選與梁文字
def combine_beam_boundingbox(coor_to_beam_list:list,coor_to_bounding_block_list:list,class_beam_list:list[Beam]):
    def _get_distance(pt1,pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    for beam in coor_to_beam_list:
        bounding_box = [block for block in coor_to_bounding_block_list if inblock(block[0],beam[1])]
        if len(bounding_box)==0:continue
        nearest_block = min(bounding_box,key=lambda b:_get_distance(b[0][0],beam[1]))
        # nearest_block[1] = beam[0]
        beam[4] = nearest_block[0]

    for beam in class_beam_list:
        bounding_box = [block for block in coor_to_bounding_block_list if inblock(block[0],beam.get_coor())]
        if len(bounding_box)==0:continue
        nearest_block = min(bounding_box,key=lambda b:_get_distance(b[0][0],beam.get_coor()))
        beam.set_bounding_box(nearest_block[0][0][0],nearest_block[0][0][1],nearest_block[0][1][0],nearest_block[0][1][1])

## 組合箍筋與梁文字
def combine_beam_tie(coor_sorted_tie_list:list,coor_to_beam_list:list,class_beam_list:list[Beam]):
    #((左下，右上),beam_name, list of tie, tie_count_dic, list of rebar,rebar_length_dic)
    def _get_distance(pt1,pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1]) 
    for tie,coor,tie_num,count,size in coor_sorted_tie_list:
        bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=coor)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < coor[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],coor))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],coor))
        nearest_beam[2].append(tie)
        if size in nearest_beam[3]:
            nearest_beam[3][size] += count
        else:
            nearest_beam[3][size] = count

    for tie,coor,tie_num,count,size in coor_sorted_tie_list:
        bounding_box = [beam for beam in class_beam_list if inblock(block=beam.get_bounding_box(),pt=coor)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in class_beam_list if beam.coor.y < coor[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b.get_coor(),coor))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b.get_coor(),coor))
        nearest_beam.add_tie(tie,coor,tie_num,count,size)

## 組合主筋與梁文字
def combine_beam_rebar(coor_to_arrow_dic:dict,coor_to_rebar_list_straight:list,coor_to_bend_rebar_list:list,coor_to_beam_list:list,class_beam_list:list[Beam]):
    def _get_distance(pt1,pt2):
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    for arrow_head,arrow_item in coor_to_arrow_dic.items():
        tail_coor,length,line_head_coor,number,size,line_tail_coor= arrow_item
        bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=arrow_head)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < arrow_head[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],arrow_head))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],arrow_head))
        nearest_beam[5].append({f'{number}-{size}':length})
        if size in nearest_beam[6]:
            if 'E.F' in size:
                nearest_beam[6][size] += length
            else:
                nearest_beam[6][size] += int(number)*length
        else:
            if 'E.F' in size:
                nearest_beam[6][size] = length
            else:
                nearest_beam[6][size] = int(number)*length

    for arrow_head,arrow_item in coor_to_arrow_dic.items():
        tail_coor,length,line_mid_coor,number,size,text_coor= arrow_item
        bounding_box = [beam for beam in class_beam_list if inblock(block=beam.get_bounding_box(),pt=arrow_head)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in class_beam_list if beam.coor.y < arrow_head[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b.get_coor(),arrow_head))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b.get_coor(),arrow_head))
        nearest_beam.add_rebar(start_pt = line_mid_coor, end_pt = line_mid_coor,
                                length = length, number = number, size = size,text=f'{number}-{size}')

    for rebar_line in coor_to_rebar_list_straight:# (頭座標，尾座標，長度，number，size)
        head_coor,tail_coor,length,number,size= rebar_line
        mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
        bounding_box = [beam for beam in class_beam_list if inblock(block=beam.get_bounding_box(),pt=mid_pt)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in class_beam_list if beam.coor.y < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b.get_coor(),mid_pt))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b.get_coor(),mid_pt))
        nearest_beam.add_rebar(start_pt = mid_pt, end_pt = mid_pt,
                                length = length, number = number, size = size,text=f'{number}-{size}',add_up='straight')

    for rebar_line in coor_to_rebar_list_straight:# (頭座標，尾座標，長度，number，size)
        head_coor,tail_coor,length,number,size= rebar_line
        mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
        bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=mid_pt)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],mid_pt))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],mid_pt))
        nearest_beam[5].append({f'支承端{number}-{size}':length})
        if rebar_line[4] in nearest_beam[6]:
            nearest_beam[6][size] += int(number)*length
        else:
            nearest_beam[6][size] = int(number)*length

    for bend_line in coor_to_bend_rebar_list:# (直的端點，橫的端點，長度，number，size)
        head_coor,tail_coor,length,number,size= bend_line
        # mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
        mid_pt = head_coor
        bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=mid_pt)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],mid_pt))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],mid_pt))
        nearest_beam[5].append({f'彎鉤{number}-{size}':length})
        if bend_line[4] in nearest_beam[6]:
            if 'E.F' in size:
                nearest_beam[6][size] += length
            else:
                nearest_beam[6][size] += int(number)*length
        else:
            if 'E.F' in size:
                nearest_beam[6][size] = length
            else:
                nearest_beam[6][size] = int(number)*length

    for bend_line in coor_to_bend_rebar_list:# (直的端點，橫的端點，長度，number，size)
        head_coor,tail_coor,length,number,size= bend_line
        # mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
        mid_pt = head_coor
        bounding_box = [beam for beam in class_beam_list if inblock(block=beam.get_bounding_box(),pt=arrow_head)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [beam for beam in class_beam_list  if beam.coor.y < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:continue
            nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b.get_coor(),mid_pt))
        else:
            nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b.get_coor(),arrow_head))
        nearest_beam.add_rebar(start_pt = mid_pt, end_pt = mid_pt,
                                length = length, number = number, size = size,text=f'{number}-{size}',add_up='bend')

## 輸出每隻梁的結果    
def count_each_beam_rebar_tie(coor_to_beam_list:list,output_txt='test.txt'):
    # (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
    lines=[]
    total_tie = {}
    # total_tie_count = {}
    total_rebar = {}
    floor_rebar = {}
    floor_concrete = {}
    floor_formwork = {}
    total_concrete = {}
    total_formwork={}
    def _add_total(size,number,total):
        if size in total:
            total[size] += number
        else:
            total[size] = number

    for beam in coor_to_beam_list:
        count_floor = beam[0].split(' ')[0]
        if count_floor not in floor_rebar:
            floor_rebar.update({count_floor:{}})
            floor_concrete.update({count_floor:0})
            floor_formwork.update({count_floor:0})
        # temp_dict = floor_rebar[count_floor]
        matches= re.findall(r"\((.*?)\)",beam[0],re.MULTILINE)
        # if len(matches) == 0 or 'X' not in matches[0]:return
        if len(matches) == 0 or len(re.findall(r"X|x",matches[0],re.MULTILINE)) == 0:continue
        split_char = re.findall(r"X|x",matches[0])[0]
        tie=0
        rebar=0
        total_tie_count = 0
        try:
            depth = int(matches[0].split(split_char)[1])
            width = int(matches[0].split(split_char)[0])
        except:
            depth = 0
            width = 0 
        tie_count = beam[3]
        rebar_count = beam[6]
        for size,count in tie_count.items():
            tie += count * RebarInfo(size) * ((depth - 10)+(width-10))*2
            total_tie_count += count
            _add_total(size=size,number=count * RebarInfo(size) * ((depth - 10)+(width-10))*2,total=total_tie)
            _add_total(size=size,number=count * RebarInfo(size) * ((depth - 10)+(width-10))*2,total=floor_rebar[count_floor])
            # _add_total(size=size,number=count,total=total_tie_count)
        for size,length in rebar_count.items():
            rebar += RebarInfo(size) * length
            _add_total(size=size,number=RebarInfo(size) * length,total=total_rebar)
            _add_total(size=size,number=RebarInfo(size) * length,total=floor_rebar[count_floor])
        lines.append('\n梁{}:'.format(beam[0]))
        lines.append('\n寬度:{}、高度:{}'.format(width,depth))
        lines.append('\n主筋為:{}'.format(beam[5]))
        lines.append('\n箍筋為:{}'.format(beam[2]))
        lines.append('\n箍筋個數為:{}'.format(total_tie_count))
        lines.append('\n主筋量為:{}'.format(rebar))
        lines.append('\n箍筋量為:{}'.format(tie))
        lines.append(f'==================================')
    with open(output_txt, 'w',encoding= 'utf-8') as f:
        for floor,item in floor_rebar.items():
            lines.append('\n{0} 鋼筋量 :{1}'.format(floor,item))
        lines.append('\n箍筋總量{}:'.format(total_tie))
        lines.append('\n主筋總量{}'.format(total_rebar))
        lines.append('\n混凝土體積已扣除與版共構區域')
        lines.append('\n模板已扣除與版共構區域')
        f.write('\n'.join(lines))
    pass
def count_floor_total_beam_rebar_tie(class_to_beam_list:list[Beam],output_txt='test.txt'):
    lines=[]
    total_tie = {}
    # total_tie_count = {}
    total_rebar = {}
    floor_rebar = {}
    floor_concrete = {}
    floor_formwork = {}
    total_concrete = 0
    total_formwork = 0
    def _add_total(size,number,total):
        if size in total:
            total[size] += number
        else:
            total[size] = number
    for beam in class_to_beam_list:
        matches= re.findall(r"\((.*?)\)",beam.serial,re.MULTILINE)
        if len(matches) == 0 or len(re.findall(r"X|x",beam.serial,re.MULTILINE)) == 0:continue
        count_floor = beam.floor
        if count_floor not in floor_rebar:
            floor_rebar.update({count_floor:{}})
            floor_concrete.update({count_floor:0})
            floor_formwork.update({count_floor:0})
        floor_concrete[count_floor] += beam.concrete
        floor_formwork[count_floor] += beam.formwork
        lines.append('\n梁{}:'.format(beam.serial))
        lines.append('\n寬度:{}、深度:{}'.format(beam.width,beam.depth))
        lines.append('\n主筋為:{}'.format(beam.get_rebar_list()))
        lines.append('\n箍筋為:{}'.format(beam.get_tie_list()))
        lines.append('\n主筋量(g)為:{}'.format(beam.get_rebar_weight()))
        lines.append('\n箍筋量(g)為:{}'.format(beam.get_tie_weight()))
        lines.append(f'==================================')
        for size,weight in beam.rebar_count.items():
            _add_total(size=size,number=weight,total=total_rebar)
            _add_total(size=size,number=weight,total=floor_rebar[count_floor])
        for size,weight in beam.tie_count.items():
            _add_total(size=size,number=weight,total=total_tie)
            _add_total(size=size,number=weight,total=floor_rebar[count_floor])

    with open(output_txt, 'w',encoding= 'utf-8') as f:
        for floor,item in floor_rebar.items():
            total_concrete +=floor_concrete[count_floor]
            total_formwork +=floor_formwork[count_floor]
            lines.append('\n{0} 鋼筋量(g) :{1}'.format(floor,item))
            lines.append('\n{0} 混凝土體積(cm3) :{1}'.format(floor,floor_concrete[count_floor]))
            lines.append('\n{0} 模板量(cm2) :{1}'.format(floor,floor_formwork[count_floor]))
            lines.append(f'==================================')
        lines.append('\n箍筋總量(g):{}'.format(total_tie))
        lines.append('\n主筋總量(g):{}'.format(total_rebar))
        lines.append('\n混凝土總體積(cm3):{}'.format(total_concrete))
        lines.append('\n模板總量(cm2):{}'.format(total_formwork))
        lines.append('\n混凝土體積已扣除與版共構區域')
        lines.append('\n模板僅考慮梁兩側及底面')
        f.write('\n'.join(lines))
    pass
def count_rebar_in_block(coor_to_arrow_dic:dict,coor_to_block_list:list,coor_to_rebar_list_straight,coor_to_bend_rebar_list):
    for x in coor_to_arrow_dic: # 新的coor_to_arrow_dic為尖點座標 -> (箭頭文字端坐標，鋼筋長度，鋼筋中點座標，數量，尺寸，文字座標)
    # 先找在哪個block裡面
        for y in coor_to_block_list: # ((左下，右上), rebar_length_dic, tie_count_dic)
            if x[0] > y[0][0][0] and x[0] < y[0][1][0] and x[1] > y[0][0][1] and x[1] < y[0][1][1]:
                # y[1] 是該格的rebar_length_dic: size -> length * number
                if coor_to_arrow_dic[x][4] not in y[1]:
                    y[1][coor_to_arrow_dic[x][4]] = float(coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
                else:
                    y[1][coor_to_arrow_dic[x][4]] += float(coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
    for rebar_line in coor_to_rebar_list_straight:# (頭座標，尾座標，長度，number，size)
        for block in coor_to_block_list:# ((左下，右上), rebar_length_dic, tie_count_dic)
            if inblock(pt=rebar_line[0],block=block[0]) and inblock(pt=rebar_line[1],block=block[0]):
                if rebar_line[4] not in block[1]:
                    block[1][rebar_line[4]] = float(rebar_line[2]) * float(rebar_line[3])
                else:
                    block[1][rebar_line[4]] += float(rebar_line[2]) * float(rebar_line[3])
    for rebar_bend in coor_to_bend_rebar_list:# (直的端點，橫的端點，長度，number，size)
        for block in coor_to_block_list:# ((左下，右上), rebar_length_dic, tie_count_dic)
            if inblock(pt=rebar_bend[0],block=block[0]) and inblock(pt=rebar_bend[1],block=block[0]):
                if rebar_bend[4] not in block[1]:
                    block[1][rebar_bend[4]] = float(rebar_bend[2]) * float(rebar_bend[3])
                else:
                    block[1][rebar_bend[4]] += float(rebar_bend[2]) * float(rebar_bend[3])
def summary_block_rebar_tie(coor_to_block_list):
    rebar_length_dic = {}
    tie_count_dic = {}
    with open(output_txt_2, 'w',encoding= 'utf-8') as f:
    # f = open(output_txt, "w", encoding = 'utf-8')
    
        for x in coor_to_block_list:
            if len(x[1]) != 0 or len(x[2]) != 0:
                f.write(f'統計左下角為{x[0][0]}，右上角為{x[0][1]}的框框內結果：\n')
                if len(x[1]) != 0:
                    f.write('鋼筋計算：\n')
                    for y in x[1]:
                        f.write(f'{y}: 總長度(長度*數量)為 {x[1][y]}\n')
                        if y in rebar_length_dic:
                            rebar_length_dic[y] += x[1][y] * RebarInfo(y)
                        else:
                            rebar_length_dic[y] = x[1][y] * RebarInfo(y)
                else:
                    f.write('此圖框內沒有鋼筋\n')
                    
                if len(x[2]) != 0:    
                    f.write('箍筋計算：\n')
                    for y in x[2]:
                        f.write(f'{y}: 總數量為 {x[2][y]}\n')
                        if y in tie_count_dic:
                            tie_count_dic[y] += x[2][y]
                        else:
                            tie_count_dic[y] = x[2][y]
                else:
                    f.write('此圖框內沒有箍筋\n')
                    
                f.write('\n')
                    
        f.write(f'統計所有結果：\n')
        f.write('鋼筋計算：\n')
        for y in rebar_length_dic:
            f.write(f'{y}: 總長度(長度*數量)為 {rebar_length_dic[y]}\n')
            
        f.write('箍筋計算：\n')
        for y in tie_count_dic:
            f.write(f'{y}: 總數量為 {tie_count_dic[y]}\n')
    pass
def inblock(block:tuple,pt:tuple):
    pt_x = pt[0]
    pt_y = pt[1]
    if len(block) == 0:return False
    if (pt_x - block[0][0])*(pt_x - block[1][0])<0 and (pt_y - block[0][1])*(pt_y - block[1][1])<0:
        return True
    return False

def cal_beam_rebar(data={},progress_file=''):
    # output_txt = f'{output_folder}{project_name}'
    if not data:
        return
    coor_to_rebar_list = data['coor_to_rebar_list'] # (頭座標，尾座標，長度)
    coor_to_bend_rebar_list = data['coor_to_bend_rebar_list'] # (直的端點，橫的端點，長度)
    coor_to_data_list = data['coor_to_data_list'] # (字串，座標)
    coor_to_arrow_dic = data['coor_to_arrow_dic'] # 尖點座標 -> 文字連接處座標
    coor_to_tie_list = data['coor_to_tie_list'] # (下座標，上座標，長度) 
    coor_to_tie_text_list = data['coor_to_tie_text_list'] # (字串，座標)
    coor_to_block_list = data['coor_to_block_list'] # ((左下，右上), rebar_length_dic, tie_count_dic)
    coor_to_beam_list = data['coor_to_beam_list'] # (string, midpoint, list of tie, tie_count_dic)
    coor_to_bounding_block_list = data['coor_to_bounding_block_list']
    class_beam_list = []
    # Step 8. 對應箭頭跟鋼筋
    coor_to_arrow_dic,no_arrow_line_list = sort_arrow_line(coor_to_arrow_dic,coor_to_rebar_list)
    progress('梁配筋圖讀取進度 8/15', progress_file)
        
    # Step 9. 對應箭頭跟文字，並完成head_to_data_dic, tail_to_data_dic
    coor_to_arrow_dic,head_to_data_dic,tail_to_data_dic = sort_arrow_to_word(coor_to_arrow_dic=coor_to_arrow_dic,
                                                                            coor_to_data_list=coor_to_data_list)
    progress('梁配筋圖讀取進度 9/15', progress_file)
    
    # Step 10. 統計目前的type跟size
    
    progress('梁配筋圖讀取進度 10/15', progress_file)
    # coor_to_rebar_list_straight_left,coor_to_rebar_list_straight_right, coor_to_bend_rebar_list,no_concat_line_list,no_concat_bend_list=concat_no_arrow_line(no_arrow_line_list=no_arrow_line_list,
    #                                                                                                                 head_to_data_dic=head_to_data_dic,
    #                                                                                                                 tail_to_data_dic=tail_to_data_dic,
    #                                                                                                                 coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 12. 拿彎的去找跟誰接在一起
    coor_to_rebar_list_straight = sort_noconcat_line(no_concat_line_list=no_arrow_line_list,head_to_data_dic=head_to_data_dic,tail_to_data_dic=tail_to_data_dic)
    coor_to_bend_rebar_list = sort_noconcat_bend(no_concat_bend_list=coor_to_bend_rebar_list,head_to_data_dic=head_to_data_dic,tail_to_data_dic=tail_to_data_dic)
    sort_rebar_bend_line(rebar_bend_list=coor_to_bend_rebar_list,rebar_line_list=coor_to_rebar_list_straight)
    #截斷處重複計算
    progress('梁配筋圖讀取進度 12/15', progress_file)

    # Step 14-15 和 16 為箍筋部分，14-15在算框框內的數量，16在算每個梁的總長度，兩者獨立
    count_rebar_in_block(coor_to_arrow_dic,coor_to_block_list,coor_to_rebar_list_straight=coor_to_rebar_list_straight,coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 14. 算箍筋
    coor_sorted_tie_list = count_tie(coor_to_tie_text_list=coor_to_tie_text_list,coor_to_block_list=coor_to_block_list,coor_to_tie_list=coor_to_tie_list)
    add_beam_to_list(coor_to_beam_list=coor_to_beam_list,class_beam_list=class_beam_list)
    combine_beam_boundingbox(coor_to_beam_list=coor_to_beam_list,coor_to_bounding_block_list=coor_to_bounding_block_list,class_beam_list=class_beam_list)
    combine_beam_tie(coor_sorted_tie_list=coor_sorted_tie_list,coor_to_beam_list=coor_to_beam_list,class_beam_list=class_beam_list)
    combine_beam_rebar(coor_to_arrow_dic=coor_to_arrow_dic,coor_to_rebar_list_straight = coor_to_rebar_list_straight,
                        coor_to_bend_rebar_list=coor_to_bend_rebar_list,coor_to_beam_list=coor_to_beam_list,class_beam_list=class_beam_list)
    sort_beam(class_beam_list=class_beam_list)
    return class_beam_list
def create_report(class_beam_list:list[Beam],output_folder:str,project_name:str,floor_parameter_xlsx:str):
    # output_txt =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar.txt')
    # output_txt_2 =os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_rebar_floor.txt')
    excel_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Count.xlsx'
    )

    beam_df = output_beam(class_beam_list=class_beam_list,output_folder=output_folder,project_name=project_name)
    # count_each_beam_rebar_tie(coor_to_beam_list=coor_to_beam_list,output_txt=output_txt)
    # count_floor_total_beam_rebar_tie(class_to_beam_list=class_beam_list,output_txt=output_txt)
    floor_list = floor_parameter(beam_list=class_beam_list,floor_parameter_xlsx=floor_parameter_xlsx)
    rebar_df,concrete_df,coupler_df,formwork_df  = summary_floor_rebar(floor_list=floor_list,item_type='beam')
    bs_list = create_beam_scan()
    scan_df = beam_check(beam_list=class_beam_list,beam_scan_list=bs_list)
    OutputExcel(df=scan_df,file_path=excel_filename,sheet_name='梁檢核表',auto_fit_columns=[1],auto_fit_rows=[1],
        columns_list=range(2,len(scan_df.columns)+2),rows_list=range(2,len(scan_df.index)+2))
    OutputExcel(df=beam_df,file_path=excel_filename,sheet_name='梁統整表')
    OutputExcel(df=rebar_df,file_path=excel_filename,sheet_name='鋼筋統計表')
    OutputExcel(df=concrete_df,file_path=excel_filename,sheet_name='混凝土統計表')
    OutputExcel(df=formwork_df,file_path=excel_filename,sheet_name='模板統計表')

    progress('梁配筋圖讀取進度 14/15', progress_file)
    # Step 15. 印出每個框框的結果然後加在一起
    
    # Step 16. 把箍筋跟beam字串綁在一起
    return excel_filename

def add_beam_to_list(coor_to_beam_list:list,class_beam_list:list):
    for beam in coor_to_beam_list:
        class_beam_list.append(Beam(beam[0],beam[1][0],beam[1][1]))
    pass
    # DEBUG # 畫線把文字跟左右的線連在一起
    # coor_list1 = [min_left_coor[0], min_left_coor[1], 0, x[1][0], x[1][1], 0]
    # coor_list2 = [min_right_coor[0], min_right_coor[1], 0, x[1][0], x[1][1], 0]
    # points1 = vtFloat(coor_list1)
    # points2 = vtFloat(coor_list2)
    # line1 = msp_beam.AddPolyline(points1)
    # line2 = msp_beam.AddPolyline(points2)
    # line1.SetWidth(0, 2, 2)
    # line2.SetWidth(0, 2, 2)
    # line1.color = 101
    # line2.color = 101
def draw_rebar_line(class_beam_list:list[Beam],msp_beam:object,doc_beam:object,output_folder:str,project_name:str):
    output_dwg = os.path.join(output_folder,f'{project_name}_{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_Markon.dwg')
    for beam in class_beam_list:
        for pos,rebar_list in beam.rebar.items():
            for rebar in rebar_list:
                x = rebar.start_pt.x
                y = rebar.start_pt.y
                coor_list1 = [beam.coor.x, beam.coor.y, 0, x, y, 0]
                coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
                points1 = vtFloat(coor_list1)
                points2 = vtFloat(coor_list2)
                line1 = msp_beam.AddPolyline(points1)
                line2 = msp_beam.AddPolyline(points2)
                text1 = msp_beam.AddMText(vtPnt((x + rebar.end_pt.x)/2, (y + rebar.end_pt.y)/2),10,rebar.text)
                text1.Height = 5
                line1.SetWidth(0, 1, 1)
                line2.SetWidth(0, 1, 1)
                line1.color = 101
                line2.color = 101
        for pos,tie in beam.tie.items():
            if(not tie):continue
            coor_list1 = [beam.coor.x, beam.coor.y, 0, tie.start_pt.x, tie.start_pt.y, 0]
            points1 = vtFloat(coor_list1)
            line1 = msp_beam.AddPolyline(points1)
            line1.SetWidth(0, 1, 1)
            line1.color = 1
        for middle_tie in beam.middle_tie:
            coor_list1 = [beam.coor.x, beam.coor.y, 0, middle_tie.start_pt.x, middle_tie.start_pt.y, 0]
            points1 = vtFloat(coor_list1) 
            line1 =msp_beam.AddPolyline(points1)
            line1.SetWidth(0, 1, 1)
            line1.color = 240

            # coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
        # coor_list2 = [min_right_coor[0], min_right_coor[1], 0, x[1][0], x[1][1], 0]
    doc_beam.SaveAs(output_dwg)
    doc_beam.Close(SaveChanges=True)
    return output_dwg
def sort_beam(class_beam_list:list[Beam]):
    for beam in class_beam_list:
        beam.sort_beam_rebar()
        beam.sort_beam_tie()
    return
def output_beam(class_beam_list:list[Beam],output_folder:str,project_name:str):
    import pandas as pd
    import numpy as np
    header_info_1 = [('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', '')]
    header_rebar = [('主筋', ''),('主筋', '左'),('主筋', '中'),('主筋', '右'),('主筋長度', '左'),('主筋長度', '中'),('主筋長度', '右')]

    header_sidebar = [('腰筋', '')]

    header_stirrup = [
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')
    ]

    header_info_2 = [
        ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'),
        ('主筋量', 'g'), ('箍筋量', 'g'),('模板', 'cm2'),('混凝土', 'cm3')
    ]
    header = pd.MultiIndex.from_tuples(
        header_info_1 + header_rebar + header_sidebar + header_stirrup + header_info_2)
    beam = pd.DataFrame(np.empty([len(class_beam_list)*4,len(header)],dtype='<U16'),columns=header)
    
    # beam = pd.DataFrame(np.empty([len(etabs_design.groupby(['Story', 'BayID'])) * 4, len(header)], dtype='<U16'), columns=header)
    row = 0
    temp_x = 0 
    rebar_pos ={
        'top_first':0,
        'top_second':1,
        'bot_first':3,
        'bot_second':2
    }
    min_diff = 1
    for b in class_beam_list:
        b.cal_rebar()
        b.sort_rebar_table()
        beam.at[row,('樓層','')] = b.floor
        beam.at[row,('編號','')] = b.serial
        beam.at[row,('RC 梁寬','')] = b.width
        beam.at[row,('RC 梁深','')] = b.depth
        beam.at[row,('腰筋', '')] = b.get_middle_tie()
        for rebar_text,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                if abs(rebar.start_pt.x - b.start_pt.x) < min_diff :
                    beam.at[row + rebar_pos[rebar_text],('主筋', '左')] = rebar.text
                if abs(rebar.end_pt.x - b.end_pt.x)< min_diff:
                    beam.at[row + rebar_pos[rebar_text],('主筋', '右')] = rebar.text
                if (abs(rebar.start_pt.x - b.start_pt.x) >= min_diff and abs(rebar.end_pt.x - b.end_pt.x)>= min_diff) or (rebar.start_pt.x == b.start_pt.x and rebar.end_pt.x == b.end_pt.x):
                    beam.at[row + rebar_pos[rebar_text],('主筋', '中')] = rebar.text
            for rebar in rebar_list:
                if abs(rebar.start_pt.x - b.start_pt.x) < min_diff:
                    beam.at[row + rebar_pos[rebar_text],('主筋長度', '左')] = rebar.length
                    continue
                if abs(rebar.end_pt.x - b.end_pt.x)< min_diff:
                    beam.at[row + rebar_pos[rebar_text],('主筋長度', '右')] = rebar.length
                    continue
                if (abs(rebar.start_pt.x - b.start_pt.x) >= min_diff and abs(rebar.end_pt.x - b.end_pt.x)>= min_diff):
                    beam.at[row + rebar_pos[rebar_text],('主筋長度', '中')] = rebar.length
                    continue
        # for tie_text,tie in b.tie.items():
        if b.tie_list:
            beam.at[row,('箍筋', '左')] = b.tie['left'].text
            beam.at[row,('箍筋', '中')] = b.tie['middle'].text
            beam.at[row,('箍筋', '右')] = b.tie['right'].text
            beam.at[row,('箍筋長度', '左')] = b.length/4
            beam.at[row,('箍筋長度', '中')] = b.length/2
            beam.at[row,('箍筋長度', '右')] = b.length/4
        beam.at[row,('主筋量', 'g')]=b.get_rebar_weight()
        beam.at[row,('箍筋量', 'g')]=b.get_tie_weight()
        beam.at[row,('模板', 'cm2')]=b.get_formwork()
        beam.at[row,('混凝土', 'cm3')]=b.get_concrete()
        row += 4
    # excel_filename = (
    #         f'{output_folder}/'
    #         f'{project_name}_'
    #         f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
    #         f'Count.xlsx'
    #     )
    # writer = pd.ExcelWriter(excel_filename)
    # beam.to_excel(writer,'Count')
    # writer.save()
    return beam

def count_beam_main(beam_filename,layer_config,temp_file='temp_1221_1F.pkl',output_folder='',project_name='',template_name=''):
    progress_file = './result/tmp'
    start = time.time()
    msp_beam,doc_beam = read_beam_cad(beam_filename=beam_filename,progress_file=progress_file)
    sort_beam_cad(msp_beam=msp_beam,layer_config=layer_config,entity_config=get_template(template_name),progress_file=progress_file,temp_file=temp_file)
    output_txt,output_txt_2,output_excel,class_beam_list = cal_beam_rebar(data=save_temp_file.read_temp(temp_file),output_folder=output_folder,project_name=project_name,progress_file=progress_file)
    output_dwg = draw_rebar_line(class_beam_list=class_beam_list,msp_beam=msp_beam,doc_beam=doc_beam,output_folder=output_folder,project_name=project_name)
    print(f'Total Time:{time.time() - start}')
    return os.path.basename(output_txt),os.path.basename(output_txt_2),os.path.basename(output_excel),os.path.basename(output_dwg)


def get_template(name:str):

    if name == '公司1':
        return  {
         'rebar_layer':['AcDbPolyline'],
         'rebar_data_layer':['AcDbMText'],
         'rebar_data_leader_layer':['AcDbLeader'],
         'tie_text_layer':['AcDbText']
        }
    if name == '公司2':
        return {
            'rebar_layer':['AcDbLine'],
            'rebar_data_layer':['AcDbText','AcDbMText'],
            'rebar_data_leader_layer':['AcDbPolyline'],
            'tie_text_layer':['AcDbMText']
        }

def floor_parameter(beam_list:list[Beam],floor_parameter_xlsx:str):
    parameter_df:pd.DataFrame
    floor_list:list[Floor]
    floor_list = []
    parameter_df = read_parameter_df(floor_parameter_xlsx,'參數表')
    parameter_df.set_index(['樓層'],inplace=True)
    for floor_name in parameter_df.index:
        temp_floor = Floor(floor_name)
        floor_list.append(temp_floor)
        temp_floor.set_prop(parameter_df.loc[floor_name])
        temp_floor.add_beam([b for b in beam_list if b.floor == temp_floor.floor_name])
    return floor_list
if __name__=='__main__':
    # from multiprocessing import Process, Pool
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # beam_filename = r"D:\Desktop\BeamQC\TEST\INPUT\2022-11-18-17-16temp-XS-BEAM.dwg"#sys.argv[1] # XS-BEAM的路徑
    beam_filename = r"D:\Desktop\BeamQC\TEST\2023-0131\XS-BEAM-新市.dwg"
    # beam_filename = r"D:\Desktop\BeamQC\TEST\DEMO\數量計算\Other-大梁\2F-大梁 - Rec.dwg"
    progress_file = './result/tmp'#sys.argv[14]
    rebar_file = './result/0107-rebar_wu2.txt' # rebar.txt的路徑 -> 計算鋼筋和箍筋總量
    tie_file = './result/0107-tie_wu2.txt' # rebar.txt的路徑 -> 把箍筋跟梁綁在一起
    output_folder ='D:/Desktop/BeamQC/TEST/OUTPUT/'
    floor_parameter_xlsx = r'D:\Desktop\BeamQC\TEST\柱樓層參數.xlsx'
    project_name = 'test_1F'
    # 在beam裡面自訂圖層
    layer_config = {
        'rebar_data_layer':'S-LEADER', # 箭頭和鋼筋文字的塗層
        'rebar_layer':'S-REINF', # 鋼筋和箍筋的線的塗層
        'tie_text_layer':'S-TEXT', # 箍筋文字圖層
        'block_layer':'DwFm', # 框框的圖層
        'beam_text_layer' :'S-RC', # 梁的字串圖層
        'bounding_block_layer':'S-ARCH'
    }
    # layer_config = {
    #     'rebar_data_layer':'NBAR', # 箭頭和鋼筋文字的塗層
    #     'rebar_layer':'RBAR', # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer':'NBAR', # 箍筋文字圖層
    #     'block_layer':'DEFPOINTS', # 框框的圖層
    #     'beam_text_layer' :'TITLE', # 梁的字串圖層
    #     'bounding_block_layer':'S-ARCH'
    # }
    entity_type ={
        'rebar_layer':['AcDbPolyline'],
        'rebar_data_layer':['AcDbMText'],
        'rebar_data_leader_layer':['AcDbLeader'],
        'tie_text_layer':['AcDbText']
    }
    
    # entity_type ={
    #     'rebar_layer':['AcDbLine'],
    #     'rebar_data_layer':['AcDbText','AcDbMText'],
    #     'rebar_data_leader_layer':['AcDbPolyline'],
    #     'tie_text_layer':['AcDbMText']
    # }
    # def test(l:list):
    #     l2 = l
    #     l2[0] = 9
        # for a in l2:
        #     if a <= 2 :
        #         # print(a) 
        #         l2.remove(a)
        #         # print(l.pop())
        #     # print(a,l)
    # l = list(range(1,10))
    # test(l)
    # print(l)
    start = time.time()
    # msp_beam,doc_beam = read_beam_cad(beam_filename=beam_filename,progress_file=progress_file)
    # sort_beam_cad(msp_beam=msp_beam,layer_config=layer_config,entity_config=entity_type,progress_file=progress_file,temp_file='temp_0201_1F.pkl')
    
    class_beam_list = cal_beam_rebar(data=save_temp_file.read_temp(r'temp_0201_1F.pkl'),progress_file=progress_file)
    create_report(class_beam_list=class_beam_list,output_folder=output_folder,project_name=project_name,floor_parameter_xlsx=floor_parameter_xlsx)
    # draw_rebar_line(class_beam_list=class_beam_list,msp_beam=msp_beam,doc_beam=doc_beam,output_folder=output_folder,project_name=project_name)
    print(f'Total Time:{time.time() - start}')
    # output_beam([Beam('1F B1-1',0,0)])
    # data=save_temp_file.read_temp('temp_1216.pkl')
    # import pprint
    # pprint.pprint(data['coor_to_bounding_block_list'])
