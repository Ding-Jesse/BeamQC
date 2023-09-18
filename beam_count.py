from __future__ import annotations
import time
import pythoncom
import win32com.client
import re
import os
import save_temp_file
import pandas as pd
import numpy as np
import copy
from math import sqrt, ceil
from item.beam import Beam, BeamType, RebarType
from item.rebar import RebarInfo
from item.floor import Floor, read_parameter_df, summary_floor_rebar, summary_floor_rebar_ratio
from item.excel import AddExcelDataBar, AddBorderLine
from item.excepteions import NoRebarDataError, BeamFloorNameError
from beam_scan import create_beam_scan, beam_check, create_sbeam_scan, create_fbeam_scan, output_detail_scan_report, output_ng_ratio
from main import OutputExcel, Add_Row_Title
from multiprocessing.pool import ThreadPool as Pool
from collections import Counter
from item.rebar import isRebarSize, readRebarExcel
from item.pdf import create_scan_pdf
error_file = './result/error_log.txt'  # error_log.txt的路徑


def vtFloat(l):  # 要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)


def vtPnt(x, y, z=0):
    """座標點轉化爲浮點數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def error(error_message):  # 把錯誤訊息印到error.log裡面
    # f = open(error_file, 'a', encoding='utf-8')
    localtime = time.asctime(time.localtime(time.time()))
    with open(error_file, 'a', encoding='utf-8') as f:
        f.write(f'{localtime} | {error_message}\n')
    # f.close()
    return


def progress(message, progress_file):  # 把進度印到progress裡面，在app.py會對這個檔案做事
    f = open(progress_file, 'a', encoding='utf-8')
    f.write(f'{message}\n')
    f.close()
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
            error(
                f'read_beam error in step 1: {e}, error_count = {error_count}.')
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
            error(
                f'read_beam error in step 2: {e}, error_count = {error_count}.')
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
            error(
                f'read_beam error in step 3: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 3/15', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False

    return msp_beam, doc_beam
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


def sort_beam_cad(msp_beam, layer_config: dict, entity_config: dict, progress_file='', temp_file=''):

    rebar_layer = layer_config['rebar_layer']
    rebar_data_layer = layer_config['rebar_data_layer']
    tie_text_layer = layer_config['tie_text_layer']
    block_layer = layer_config['block_layer']
    beam_text_layer = layer_config['beam_text_layer']
    bounding_block_layer = layer_config['bounding_block_layer']
    beam_layer = layer_config['rc_block_layer']
    s_dim_layer = layer_config['s_dim_layer']
    print(f'temp_file:{temp_file}')
    coor_to_rebar_list = []  # (頭座標，尾座標，長度)
    coor_to_bend_rebar_list = []  # (直的端點，橫的端點，長度)
    coor_to_data_list = []  # (字串，座標)
    coor_to_arrow_dic = {}  # 尖點座標 -> 文字連接處座標
    coor_to_tie_list = []  # (下座標，上座標，長度)
    coor_to_tie_text_list = []  # (字串，座標)
    coor_to_block_list = []  # ((左下，右上), rebar_length_dic, tie_count_dic)
    # (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
    coor_to_beam_list = []
    # ((左下，右上),beam_name, list of tie, tie_count_dic, list of rebar,rebar_length_dic)
    coor_to_bounding_block_list = []
    coor_to_rc_block_list = []
    coor_to_dim_list = []
    count = 0
    total = msp_beam.Count
    progress(
        f'{temp_file}:梁配筋圖上共有{total}個物件，大約運行{int(total / 5500)}分鐘，請耐心等候', progress_file)
    for object in msp_beam:
        error_count = 0
        while error_count < 10:
            try:
                # print(f'{object.ObjectName}:{object.Layer}')
                count += 1
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
                # 抓鋼筋的字的座標
                if object.Layer in rebar_data_layer and object.ObjectName in entity_config['rebar_data_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))
                    coor_to_data_list.append((object.TextString, coor))
                # 抓箭頭座標
                elif object.Layer in rebar_data_layer and object.ObjectName in entity_config['rebar_data_leader_layer']:
                    # object.Coordinates 有九個參數 -> 箭頭尖點座標，直角的座標，文字接出去的座標，都有x, y, z
                    if hasattr(object, 'Coordinates'):
                        coor_to_arrow_dic[(round(object.Coordinates[0], 2), round(object.Coordinates[1], 2))] = (
                            round(object.Coordinates[6], 2), round(object.Coordinates[7], 2))
                    if hasattr(object, 'startPoint'):
                        coor_to_arrow_dic[(round(object.startPoint[0], 2), round(object.startPoint[1], 2))] = (
                            round(object.endPoint[0], 2), round(object.endPoint[1], 2))
                # 抓鋼筋本人和箍筋本人
                elif object.Layer in rebar_layer:
                    # object.Coordinates 橫的和直的有四個參數 -> 兩端點的座標，都只有x, y; 彎的有八個參數 -> 直的端點，直的轉角，橫的轉角，橫的端點
                    if (object.ObjectName == 'AcDbPolyline'):
                        if round(object.Length, 4) > 4:  # 太短的是分隔線 -> 不要
                            # 橫的 -> y 一樣 -> 鋼筋
                            if len(object.Coordinates) == 4 and round(object.Coordinates[1], 2) == round(object.Coordinates[3], 2):
                                coor_to_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(
                                    object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4)))
                            # 直的 -> x 一樣 -> 箍筋
                            elif len(object.Coordinates) == 4 and round(object.Coordinates[0], 2) == round(object.Coordinates[2], 2):
                                coor_to_tie_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(
                                    object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4)))
                            elif len(object.Coordinates) == 8:  # 彎的 -> 直的端點，橫的端點
                                coor_to_bend_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(
                                    object.Coordinates[6], 2), round(object.Coordinates[7], 2)), round(object.Length, 4)))
                    elif (object.ObjectName == 'AcDbLine'):
                        # 橫的 -> y 一樣 -> 鋼筋
                        if round(object.startPoint[1], 2) == round(object.endPoint[1], 2):
                            coor_to_rebar_list.append(((round(object.startPoint[0], 2), round(object.startPoint[1], 2)), (round(
                                object.endPoint[0], 2), round(object.endPoint[1], 2)), round(object.Length, 4)))
                        # 直的 -> x 一樣 -> 箍筋
                        elif round(object.startPoint[0], 2) == round(object.endPoint[0], 2):
                            coor_to_tie_list.append(((round(object.startPoint[0], 2), round(object.startPoint[1], 2)), (round(
                                object.endPoint[0], 2), round(object.endPoint[1], 2)), round(object.Length, 4)))
                # 抓箍筋文字座標
                elif object.Layer in tie_text_layer and object.ObjectName in entity_config['tie_text_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))
                    coor_to_tie_text_list.append((object.TextString, coor))
                    break
                # 抓圖框
                elif object.Layer in block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    coor_to_block_list.append(((coor1, coor2), {}, {}))
                # 抓圖框
                elif object.Layer in bounding_block_layer and (object.EntityName == "AcDbBlockReference" or object.EntityName == "AcDbPolyline"):
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    coor_to_bounding_block_list.append(
                        ((coor1, coor2), "", [], {}, [], {}))
                # 抓梁的字的座標
                elif object.Layer in beam_text_layer and object.ObjectName == 'AcDbText':
                    midpoint = (round((object.GetBoundingBox()[0][0] + object.GetBoundingBox()[1][0]) / 2, 2),
                                round((object.GetBoundingBox()[0][1] + object.GetBoundingBox()[1][1]) / 2, 2))
                    # (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
                    coor_to_beam_list.append(
                        [object.TextString, midpoint, [], {}, (), [], {}])
                elif object.Layer in beam_text_layer and object.ObjectName == 'AcDbMText':
                    midpoint = (round((object.GetBoundingBox()[0][0] + object.GetBoundingBox()[1][0]) / 2, 2),
                                round((object.GetBoundingBox()[0][1] + object.GetBoundingBox()[1][1]) / 2, 2))
                    coor_to_beam_list.append(
                        [object.TextString, midpoint, [], {}, (), [], {}])
                # 抓箍筋文字
                if object.Layer in tie_text_layer and object.ObjectName in entity_config['tie_text_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))
                    coor_to_tie_text_list.append((object.TextString, coor))
                if object.Layer in beam_layer and object.ObjectName in ['AcDbPolyline']:
                    if len(object.Coordinates) >= 8:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_rc_block_list.append(((coor1, coor2), ''))
                if object.Layer in s_dim_layer and hasattr(object, 'TextPosition') and hasattr(object, 'Measurement'):
                    if len(object.GetBoundingBox()) >= 2:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        if hasattr(object, 'ExtLine1Point') and hasattr(object, 'ExtLine2Point'):
                            coor1 = object.ExtLine1Point
                            coor2 = object.ExtLine2Point
                        coor_to_dim_list.append(
                            (object.TextPosition, object.Measurement, (coor1, coor2)))
                break
            except Exception:
                # raise
                # print(f'error:{error_count}')
                error_count += 1
                time.sleep(5)
    progress('梁配筋圖讀取進度 7/15', progress_file)
    save_temp_file.save_pkl({'coor_to_data_list': coor_to_data_list,
                            'coor_to_arrow_dic': coor_to_arrow_dic,
                             'coor_to_rebar_list': coor_to_rebar_list,
                             'coor_to_bend_rebar_list': coor_to_bend_rebar_list,
                             'coor_to_tie_list': coor_to_tie_list,
                             'coor_to_tie_text_list': coor_to_tie_text_list,
                             'coor_to_block_list': coor_to_block_list,
                             'coor_to_beam_list': coor_to_beam_list,
                             'coor_to_bounding_block_list': coor_to_bounding_block_list,
                             'coor_to_rc_block_list': coor_to_rc_block_list,
                             'coor_to_dim_list': coor_to_dim_list
                             }, temp_file)

# 整理箭頭與直線對應


def sort_arrow_line(coor_to_arrow_dic: dict,
                    coor_to_rebar_list: list,
                    coor_to_dim_list: list[tuple[tuple[float, float], float, tuple[tuple[float, float], tuple[float, float]]]],
                    coor_to_bend_rebar_list: list[tuple[tuple[float, float], tuple[float, float], float]],
                    progress_file: str):
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
    # method 2
    new_coor_to_arrow_dic = {}
    no_arrow_line_list = []
    no_arrow_dim_list = []
    min_diff = 1

    # method 3
    # for i,dim in enumerate(coor_to_dim_list):
    #     text_postion, text_value, line_1_point,line_2_point = dim
    #     arrow_dict = {k: v for k, v in coor_to_arrow_dic.items() if (line_1_point[0] - k[0]) * (line_2_point[0] - k[0]) <= 0}
    #     if arrow_dict:
    #         value_pair = min(arrow_dict.items(),key=lambda x:abs(text_postion[1] - x[0][1]))
    #         if(abs(value_pair[0][1] - text_postion[1])> min_diff):
    #             no_arrow_dim_list.append(dim)
    #             continue
    #         key,value = arrow_dict.items()
    #         rebar_coor1 = (text_postion[0],key[1])
    #         rebar_coor2 = (value[0],value[1])
    #         new_coor_to_arrow_dic.update({rebar_coor1:(rebar_coor2,text_value,text_postion)})
    #         arrow_dict.pop(key)
    # return new_coor_to_arrow_dic,no_arrow_line_list

    for i, rebar in enumerate(coor_to_rebar_list):

        head_coor = rebar[0]
        tail_coor = rebar[1]
        mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
        length = rebar[2]
        arrow_dict = {k: v for k, v in coor_to_arrow_dic.items() if (
            head_coor[0] - k[0]) * (tail_coor[0] - k[0]) <= 0}
        if arrow_dict:
            value_pair = min(arrow_dict.items(),
                             key=lambda x: abs(mid_coor[1] - x[0][1]))
            if (abs(value_pair[0][1] - mid_coor[1]) > min_diff):
                no_arrow_line_list.append(rebar)
                continue
            for key, value in {k: v for k, v in arrow_dict.items() if abs(k[1] - value_pair[0][1]) < min_diff}.items():
                with_dim = False
                mid_coor = (
                    round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
                length = rebar[2]
                # try:
                #     assert (key[0] != -66308.08)
                # except:
                #     print('')
                rebar_coor1 = key
                rebar_coor2 = value
                # 下層筋
                if key[1] > value[1]:
                    dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < key[1]) and (
                        dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
                # 上層筋
                else:
                    dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] > key[1]) and (
                        dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
                if dim_list:
                    dim = min(dim_list, key=lambda dim: abs(
                        dim[0][1] - key[1]))
                    text_postion, text_value, (line_1_point,
                                               line_2_point) = dim
                    if abs(text_postion[1] - key[1]) < 150:
                        rebar_coor1 = (
                            (line_1_point[0] + line_2_point[0])/2, head_coor[1])
                        rebar_coor2 = (value[0], value[1])
                        length = text_value
                        with_dim = True
                        mid_coor = (
                            (line_1_point[0] + line_2_point[0])/2, head_coor[1])
                new_coor_to_arrow_dic.update(
                    {rebar_coor1: (rebar_coor2, length, mid_coor, with_dim)})
    for bend_rebar in coor_to_bend_rebar_list:
        arrow_dict = {}
        bend_coor = bend_rebar[0]
        line_coor = bend_rebar[1]
        mid_coor = (round((bend_coor[0] + line_coor[0]) / 2, 2), line_coor[1])
        length = rebar[2]
        if length <= 100:
            continue
        arrow_dict = {k: v for k, v in coor_to_arrow_dic.items() if (
            bend_coor[0] - k[0]) * (line_coor[0] - k[0]) <= 0}
        if not arrow_dict:
            continue
        value_pair = min(arrow_dict.items(),
                         key=lambda x: abs(mid_coor[1] - x[0][1]))
        if (abs(value_pair[0][1] - mid_coor[1]) <= min_diff):
            with_dim = False
            mid_coor = (
                round((bend_coor[0] + line_coor[0]) / 2, 2), line_coor[1])
            length = rebar[2]
            for key, value in {k: v for k, v in arrow_dict.items() if abs(k[1] - value_pair[0][1]) < min_diff}.items():
                with_dim = False
                mid_coor = (
                    round((bend_coor[0] + line_coor[0]) / 2, 2), line_coor[1])
                length = rebar[2]
                # try:
                #     assert (key[0] != -66308.08)
                # except:
                #     print('')
                rebar_coor1 = key
                rebar_coor2 = value
                # 下層筋
                if key[1] > value[1]:
                    dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < key[1]) and (
                        dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
                # 上層筋
                else:
                    dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] > key[1]) and (
                        dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
                if dim_list:
                    dim = min(dim_list, key=lambda dim: abs(
                        dim[0][1] - key[1]))
                    text_postion, text_value, (line_1_point,
                                               line_2_point) = dim
                    if abs(text_postion[1] - key[1]) < 150:
                        rebar_coor1 = (
                            (line_1_point[0] + line_2_point[0])/2, line_coor[1])
                        rebar_coor2 = (value[0], value[1])
                        length = text_value
                        with_dim = True
                        mid_coor = (
                            (line_1_point[0] + line_2_point[0])/2, line_coor[1])
                new_coor_to_arrow_dic.update(
                    {rebar_coor1: (rebar_coor2, length, mid_coor, with_dim)})
    progress(f'sort arrow to line:{time.time() - start}', progress_file)
    return new_coor_to_arrow_dic, no_arrow_line_list
    #         # new_coor_to_arrow_dic.update({value_pair[0]:(value_pair[1],length,mid_coor)})
    # print(f'Method 2:{time.time() - start}')
    # print(new_coor_to_arrow_dic_2 == new_coor_to_arrow_dic)
    # print(set(new_coor_to_arrow_dic_2.items()) - set(new_coor_to_arrow_dic.items()))
    # print(set(new_coor_to_arrow_dic.items()) - set(new_coor_to_arrow_dic_2.items()))

# 整理箭頭與鋼筋文字對應


def sort_arrow_to_word(coor_to_arrow_dic: dict,
                       coor_to_data_list: list,
                       progress_file: str):
    def _get_distance(pt1, pt2):
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
    min_diff = 1000
    new_coor_to_arrow_dic = {}
    head_to_data_dic = {}  # 座標 -> (number, size)
    tail_to_data_dic = {}
    text: str

    # method 2
    # 腰筋會抓到上層筋
    for arrow_head, arrow_data in coor_to_arrow_dic.items():
        arrow_tail, rebar_length, line_coor, with_dim = arrow_data
        rebar_data_temp = []
        if arrow_head[1] > arrow_tail[1]:
            # 箭頭朝上
            rebar_data_temp = [
                r for r in coor_to_data_list if '@' not in r[0] and r[1][1] <= arrow_head[1]]
        if arrow_head[1] < arrow_tail[1]:
            # 箭頭朝下
            rebar_data_temp = [
                r for r in coor_to_data_list if '@' not in r[0] and r[1][1] >= arrow_head[1]]
        if not rebar_data_temp:
            progress(
                f'arrow head:{arrow_head} has no rebar data', progress_file)
            # print(f'arrow head:{arrow_head} has no rebar data')
            rebar_data_temp = [r for r in coor_to_data_list if '@' not in r[0]]
        if not rebar_data_temp:
            raise NoRebarDataError

        text, coor = min(rebar_data_temp, key=lambda rebar_text: _get_distance(
            arrow_data[0], rebar_text[1]))
        if (abs(arrow_tail[1] - coor[1]) > min_diff):
            progress(
                f'{arrow_head} / {arrow_data} cant find pair arrow', progress_file)
            # print(f'{arrow_head} / {arrow_data} cant find pair arrow')
            continue
        rebar_data = list(arrow_data)
        if '-' not in text:
            progress(f'{text} not satisfied  rebar rule', progress_file)
            # print(f'{text} not satisfied rule')
            continue
        if '@' in text:
            progress(f'{text} not satisfied tie rule', progress_file)
            # print(f'{text} not satisfied rule')
            continue
        number = text.split('-')[0]
        size = text.split('-')[1]
        if not isRebarSize(size):
            progress(f'{size} not satisfied rebar rule', progress_file)
            # print(size)
            continue
        if not number.isdigit():
            progress(f'{text} not satisfied rebar rule', progress_file)
            # print(text)
            continue
        rebar_data.extend([number, size, coor])
        new_coor_to_arrow_dic.update({arrow_head: (*rebar_data,)})
        head_to_data_dic.update({(line_coor[0] - rebar_length/2, line_coor[1]): {
                                'number': number, 'size': size, 'dim': with_dim}})
        tail_to_data_dic.update({(line_coor[0] + rebar_length/2, line_coor[1]): {
                                'number': number, 'size': size, 'dim': with_dim}})

    progress(f'sort arrow to word:{time.time() - start}', progress_file)
    return new_coor_to_arrow_dic, head_to_data_dic, tail_to_data_dic


def sort_noconcat_line(no_concat_line_list, head_to_data_dic: dict, tail_to_data_dic: dict):
    # start = time.time()
    coor_to_rebar_list_straight = []  # (頭座標，尾座標，長度，number，size)

    def _overlap(l1, l2):
        if l1[1] == l2[0][1]:
            return round(l2[0][0] - l1[0], 2)*round(l2[1][0] - l1[0], 2) <= 0
        return False

    def _cal_length(pt1, pt2):
        return sqrt((pt1[0]-pt2[0])**2 + (pt1[1]-pt2[1])**2)

    def _concat_line(line_list: list):
        for line in line_list[:]:
            head_coor = min([line[0], line[1]], key=lambda l: l[0])
            tail_coor = max([line[0], line[1]], key=lambda l: l[0])
            head_rebar = {}
            tail_rebar = {}
            overlap_line = {k: v for k, v in head_to_data_dic.items(
            ) if _overlap(k, (head_coor, tail_coor))}
            if len(overlap_line.keys()) > 0:
                value_key, value_items = min(
                    overlap_line.items(), key=lambda x: x[0][0])
                tail_coor = value_key
                tail_rebar = value_items
            overlap_line = {k: v for k, v in tail_to_data_dic.items(
            ) if _overlap(k, (head_coor, tail_coor))}
            if len(overlap_line.keys()) > 0:
                value_key, value_items = max(
                    overlap_line.items(), key=lambda x: x[0][0])
                head_coor = value_key
                head_rebar = value_items
            if (not head_rebar) and (not tail_rebar):
                # print(f'{head_coor},{tail_coor} norebar')
                continue
            elif head_rebar == tail_rebar:
                coor_to_rebar_list_straight.append((head_coor, tail_coor,
                                                    _cal_length(
                                                        head_coor, tail_coor),
                                                    head_rebar['number'], head_rebar['size'], False))
            elif head_rebar != tail_rebar:
                if head_rebar and tail_rebar:
                    if head_rebar['dim']:
                        coor_to_rebar_list_straight.append((head_coor, tail_coor,
                                                            _cal_length(
                                                                head_coor, tail_coor),
                                                            tail_rebar['number'], tail_rebar['size'], False))
                    if tail_rebar['dim']:
                        coor_to_rebar_list_straight.append((head_coor,
                                                            tail_coor,
                                                            _cal_length(
                                                                head_coor, tail_coor),
                                                            head_rebar['number'],
                                                            head_rebar['size'], False))
                    # print(f'{head_coor},{tail_coor} head_rebar:{head_rebar} tail_rebar:{tail_rebar}')
                elif head_rebar:
                    coor_to_rebar_list_straight.append((head_coor,
                                                        tail_coor,
                                                        _cal_length(
                                                            head_coor, tail_coor),
                                                        head_rebar['number'],
                                                        head_rebar['size'], False))
                elif tail_rebar:
                    coor_to_rebar_list_straight.append((head_coor, tail_coor,
                                                        _cal_length(
                                                            head_coor, tail_coor),
                                                        tail_rebar['number'],
                                                        tail_rebar['size'],
                                                        False))
            head_to_data_dic.update({head_coor: {'number': coor_to_rebar_list_straight[-1][3],
                                                 'size': coor_to_rebar_list_straight[-1][4],
                                                 'dim': coor_to_rebar_list_straight[-1][5]}})
            tail_to_data_dic.update({tail_coor: {'number': coor_to_rebar_list_straight[-1][3],
                                                 'size': coor_to_rebar_list_straight[-1][4],
                                                 'dim': coor_to_rebar_list_straight[-1][5]}})
            line_list.remove(line)
            # print(f'{head_coor},{tail_coor} rebar:{coor_to_rebar_list_straight[-1][3]}-{coor_to_rebar_list_straight[-1][4]}')

    while True:
        temp_count = len(no_concat_line_list)
        _concat_line(no_concat_line_list)
        if temp_count == len(no_concat_line_list) or len(no_concat_line_list) == 0:
            break
    # print(f'Method 1:{time.time() - start}')
    return coor_to_rebar_list_straight


def sort_noconcat_bend(no_concat_bend_list: list, head_to_data_dic: dict, tail_to_data_dic: dict):
    # start = time.time()
    coor_to_bend_rebar_list = []
    for bend in no_concat_bend_list:
        horz_coor = bend[1]
        vert_coor = bend[0]
        line_length = bend[2]
        overlap_line = {k: v for k, v in head_to_data_dic.items() if (
            k[0] <= horz_coor[0]) and (k[1] == horz_coor[1]) and (k[0] >= vert_coor[0])}
        if len(overlap_line.keys()) > 0:
            value_key, value_items = min(
                overlap_line.items(), key=lambda x: abs(x[0][0]-horz_coor[0]))
            coor_to_bend_rebar_list.append((vert_coor, value_key, line_length - abs(
                value_key[0] - horz_coor[0]), value_items['number'], value_items['size']))
            # print(f'{horz_coor},{vert_coor} rebar:{coor_to_bend_rebar_list[-1][3]}-{coor_to_bend_rebar_list[-1][4]}')
            continue
        overlap_line = {k: v for k, v in tail_to_data_dic.items() if (
            k[0] >= horz_coor[0]) and (k[1] == horz_coor[1]) and (k[0] <= vert_coor[0])}
        if len(overlap_line.keys()) > 0:
            value_key, value_items = max(
                overlap_line.items(), key=lambda x: abs(x[0][0]-horz_coor[0]))
            coor_to_bend_rebar_list.append((vert_coor, value_key, line_length - abs(
                value_key[0] - horz_coor[0]), value_items['number'], value_items['size']))
            # print(f'{horz_coor},{vert_coor} rebar:{coor_to_bend_rebar_list[-1][3]}-{coor_to_bend_rebar_list[-1][4]}')
            continue
    # print(f'Method 1:{time.time() - start}')
    return coor_to_bend_rebar_list


def sort_rebar_bend_line(rebar_bend_list: list, rebar_line_list: list):
    def _between_pt(pt1, pt2, pt):
        return (pt[0] - pt1[0])*(pt[0] - pt2[0]) < 0 and pt1[1] == pt2[1] == pt[1]

    def _outer_pt(start_pt, end_pt, pt):
        if start_pt[0] < end_pt[0]:
            return pt[0] > end_pt[0]
        if start_pt[0] > end_pt[0]:
            return pt[0] < end_pt[0]
        return False

    def _overline(start_pt, end_pt, line):
        if _between_pt(start_pt, end_pt, line[0]) and _outer_pt(start_pt, end_pt, line[1]):
            return True
        if _between_pt(start_pt, end_pt, line[1]) and _outer_pt(start_pt, end_pt, line[0]):
            return True
    for bend in rebar_bend_list[:]:
        vert_coor = bend[0]
        horz_coor = bend[1]
        bend_length = bend[2]
        rebar_size = bend[4]
        rebar_number = bend[3]

        end_pt = (vert_coor[0], horz_coor[1])
        start_pt = (horz_coor[0], horz_coor[1])
        concat_line = [l for l in rebar_line_list if _overline(
            start_pt, end_pt, (l[0], l[1])) and l[4] == rebar_size]
        if concat_line:
            closet_line = min(concat_line, key=lambda l: int(l[3]))
            new_number = int(rebar_number) - int(closet_line[3])
            if new_number > 0:
                rebar_bend_list.remove(bend)
                rebar_bend_list.append(
                    (vert_coor, horz_coor, bend_length, str(new_number), rebar_size))
                # print(f'{horz_coor} {rebar_number}-{rebar_size} => {new_number}-{rebar_size}')


def count_tie(coor_to_tie_text_list: list, coor_to_block_list: list, coor_to_tie_list):
    def extract_tie(tie: str):
        new_tie = re.findall(r'\d*-\d*#\d@\d\d', tie)
        if len(new_tie) == 0:
            new_tie = re.findall(r'\d#\d@\d\d', tie)
        if len(new_tie) == 0:
            new_tie = re.findall(r'#\d@\d\d', tie)
        if len(new_tie) == 0:
            return tie
        return new_tie[0]
    tie_num = ''
    tie_text = ''
    count = 1
    size = ''
    coor_sorted_tie_list = []
    for tie, coor in coor_to_tie_text_list:  # (字串，座標)
        tie = extract_tie(tie=tie)
        if '-' in tie:
            tie_num = tie.split('-')[0]
            tie_text = tie.split('-')[1]
            if tie_num.isdigit():  # 已經算好有幾根就直接用
                count = int(tie_num)
                size = tie_text.split('@')[0]  # 用'-'和'@'來切
                if size.split('#')[0].isdigit():
                    count *= int(size.split('#')[0])
                    size = f"#{size.split('#')[1]}"
                coor_sorted_tie_list.append((tie, coor, tie_num, count, size))
                # for block in coor_to_block_list:
                #     if inblock(block=block[0],pt=coor):
                #         # print(f'pt:{coor} in block:{block[0]}')
                #         # y[2] 是該格的tie_count_dic: size -> number
                #         if size not in block[2]:
                #             block[2][size] = count
                #         else:
                #             block[2][size] += count
                #         break

        else:  # 沒算好自己算
            if not '@' in tie or not '#' in tie:
                print(f'{tie} wrong format ex:#4@20')
                continue
            size = tie.split('@')[0]  # 用'@'來切
            try:
                spacing = int(tie.split('@')[1])
            except Exception:
                continue
            assert spacing != 0, f'{coor} spacing is zero'

            tie_left_list = [(bottom, top, length) for bottom, top, length in coor_to_tie_list if (
                bottom[0] < coor[0]) and (min(bottom[1], top[1]) < coor[1]) and (max(bottom[1], top[1]) > coor[1])]
            tie_right_list = [(bottom, top, length) for bottom, top, length in coor_to_tie_list if (
                bottom[0] > coor[0]) and (min(bottom[1], top[1]) < coor[1]) and (max(bottom[1], top[1]) > coor[1])]
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

            left_tie = min(tie_left_list, key=lambda t: abs(t[0][0] - coor[0]))
            right_tie = min(
                tie_right_list, key=lambda t: abs(t[0][0] - coor[0]))

            count = int(abs(left_tie[0][0] - right_tie[0][0]) / spacing)
            if size.split('#')[0].isdigit():
                count *= int(size.split('#')[0])
                size = f"#{size.split('#')[1]}"
            coor_sorted_tie_list.append((tie, coor, tie_num, count, size))
            # for block in coor_to_block_list:
            #     if inblock(block=block[0],pt=coor):
            #         # print(f'pt:{coor} in block:{block[0]}')
            #         # y[2] 是該格的tie_count_dic: size -> number
            #         if size not in block[2]:
            #             block[2][size] = count
            #         else:
            #             block[2][size] += count
            #         break
    return coor_sorted_tie_list

# 組合手動框選與梁文字


def combine_beam_boundingbox(coor_to_block_list: list[tuple[tuple[tuple, tuple], dict, dict]],
                             coor_to_bounding_block_list: list,
                             class_beam_list: list[Beam],
                             coor_to_rc_block_list: list):
    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])

    def _get_distance_block(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])*3
    # for beam in coor_to_beam_list:
    #     bounding_box = [block for block in coor_to_bounding_block_list if inblock(block[0],beam[1])]
    #     if len(bounding_box)==0:continue
    #     nearest_block = min(bounding_box,key=lambda b:_get_distance(b[0][0],beam[1]))
    #     # nearest_block[1] = beam[0]
    #     beam[4] = nearest_block[0]
    same_block_rc_block_list = []

    for beam in class_beam_list:
        if beam.beam_type == BeamType.FB:
            accept_y_diff = 50
        else:
            accept_y_diff = 0
        # 先找圖框
        outer_block = [block for block in coor_to_block_list if inblock(
            block[0], beam.get_coor())]
        bounding_box = [block for block in coor_to_bounding_block_list if inblock(
            block[0], beam.get_coor())]
        if len(bounding_box) == 0:
            if outer_block:
                same_block_rc_block_list = [block for block in coor_to_rc_block_list if inblock(
                    outer_block[0][0], block[0][0])]

            left_bounding_box_list = [block for block in same_block_rc_block_list if block[0][1][0] <= beam.get_coor(
            )[0] and block[0][1][1] > beam.get_coor()[1] - accept_y_diff]
            right_bounding_box_list = [block for block in same_block_rc_block_list if block[0][0][0] >= beam.get_coor(
            )[0] and block[0][0][1] > beam.get_coor()[1] - accept_y_diff]
            if len(left_bounding_box_list) and len(right_bounding_box_list):
                left_bounding_box = min(
                    left_bounding_box_list, key=lambda b: _get_distance_block(b[0][0], beam.get_coor()))
                right_bounding_box = min(
                    right_bounding_box_list, key=lambda b: _get_distance_block(b[0][0], beam.get_coor()))
                top_bounding = max(left_bounding_box[0][0][1], left_bounding_box[0]
                                   [1][1], right_bounding_box[0][0][1], right_bounding_box[0][1][1])
                bot_bounding = min(left_bounding_box[0][0][1], left_bounding_box[0]
                                   [1][1], right_bounding_box[0][0][1], right_bounding_box[0][1][1])
                # 邊框距離編號過近
                if abs(beam.get_coor()[1] - max(top_bounding, bot_bounding)) < 100:
                    left_bounding_box_list = [
                        new_block for new_block in left_bounding_box_list if new_block[0][1][1] > top_bounding + 15]
                    right_bounding_box_list = [
                        new_block for new_block in right_bounding_box_list if new_block[0][1][1] > top_bounding + 15]
                    new_top_bound = []
                    if not len(left_bounding_box_list) and not len(right_bounding_box_list):
                        print(
                            f'{beam.floor}{beam.serial} no bounding, use outer block')
                        top_bounding = outer_block[0][0][1][1]
                        beam.set_bounding_box(left_bounding_box[0][1][0], beam.get_coor()[
                                              1], right_bounding_box[0][0][0], top_bounding)
                        continue
                    if left_bounding_box_list:
                        new_left_bounding_box = min(
                            left_bounding_box_list, key=lambda b: _get_distance_block(b[0][0], beam.get_coor()))
                        new_top_bound.append(
                            max(new_left_bounding_box[0][0][1], new_left_bounding_box[0][1][1]))
                    if right_bounding_box_list:
                        new_right_bounding_box = min(
                            right_bounding_box_list, key=lambda b: _get_distance_block(b[0][0], beam.get_coor()))
                        new_top_bound.append(
                            max(new_right_bounding_box[0][0][1], new_right_bounding_box[0][1][1]))
                    top_bounding = min(new_top_bound)
                beam.set_bounding_box(left_bounding_box[0][1][0], beam.get_coor()[
                                      1], right_bounding_box[0][0][0], top_bounding)
        else:
            nearest_block = min(
                bounding_box, key=lambda b: _get_distance(b[0][0], beam.get_coor()))
            beam.set_bounding_box(
                nearest_block[0][0][0], nearest_block[0][0][1], nearest_block[0][1][0], nearest_block[0][1][1])

# 組合箍筋與梁文字


def combine_beam_tie(coor_sorted_tie_list: list, coor_to_beam_list: list, class_beam_list: list[Beam]):
    # ((左下，右上),beam_name, list of tie, tie_count_dic, list of rebar,rebar_length_dic)
    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    # for tie,coor,tie_num,count,size in coor_sorted_tie_list:
    #     bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=coor)]
    #     if len(bounding_box) == 0:
    #         coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < coor[1]]
    #         if len(coor_sorted_beam_list) == 0:continue
    #         nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],coor))
    #     else:
    #         nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],coor))
    #     nearest_beam[2].append(tie)
    #     if size in nearest_beam[3]:
    #         nearest_beam[3][size] += count
    #     else:
    #         nearest_beam[3][size] = count

    for tie, coor, tie_num, count, size in coor_sorted_tie_list:
        bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=coor)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [
                beam for beam in class_beam_list if beam.coor.y < coor[1]]
            if len(coor_sorted_beam_list) == 0:
                continue
            nearest_beam = min(coor_sorted_beam_list,
                               key=lambda b: _get_distance(b.get_coor(), coor))
        else:
            nearest_beam = min(
                bounding_box, key=lambda b: _get_distance(b.get_coor(), coor))
        match_obj = re.search(r'(\d*)([#|D]\d+)[@](\d+)', tie)
        if match_obj:
            nearest_beam.add_tie(tie, coor, tie_num, count, size)
        else:
            print(f'{tie} does not match tie type')
# 截斷主筋


def break_down_rebar(coor_to_arrow_dic: dict, class_beam_list: list[Beam]):
    def _get_distance(pt1, pt2):
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    add_list = []

    for arrow_head in list(coor_to_arrow_dic.keys()):
        nearest_beam = None
        # try:
        #     assert arrow_head != (7439.0, 2504.36)
        # except:
        #     print('')
        right_nearest_beam = None
        left_nearest_beam = None
        arrow_item = coor_to_arrow_dic[arrow_head]

        tail_coor, length, line_mid_coor, with_dim, number, size, text_coor = arrow_item
        line_left_pt = (line_mid_coor[0] - length/2, line_mid_coor[1])
        line_right_pt = (line_mid_coor[0] + length/2, line_mid_coor[1])

        bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=arrow_head)]
        if len(bounding_box) > 1:
            nearest_beam = min(bounding_box, key=lambda b: _get_distance(
                b.get_coor(), arrow_head))
        elif bounding_box:
            nearest_beam = bounding_box[0]
        if not nearest_beam:
            # print(f'{arrow_head} with no bounding')
            continue
        left_nearest_beam = nearest_beam
        right_nearest_beam = nearest_beam

        left_bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=line_left_pt)]
        if len(left_bounding_box) > 1:
            left_nearest_beam = min(
                left_bounding_box, key=lambda b: _get_distance(b.get_coor(), line_left_pt))
        elif left_bounding_box:
            left_nearest_beam = left_bounding_box[0]

        right_bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=line_right_pt)]
        if len(right_bounding_box) > 1:
            right_nearest_beam = min(
                right_bounding_box, key=lambda b: _get_distance(b.get_coor(), line_right_pt))
        elif right_bounding_box:
            right_nearest_beam = right_bounding_box[0]

        if nearest_beam.serial and left_nearest_beam and right_nearest_beam:
            left_bound = nearest_beam.get_bounding_box()[0]
            right_bound = nearest_beam.get_bounding_box()[1]
            if (nearest_beam.serial != left_nearest_beam.serial or
                nearest_beam.serial != right_nearest_beam.serial or
                abs(line_left_pt[0] - left_bound[0]) > 20 or
                    abs(line_right_pt[0] - right_bound[0]) > 20) and not with_dim:
                new_left_pt_x = max(line_left_pt[0], left_bound[0])
                new_right_pt_x = min(line_right_pt[0], right_bound[0])
                if new_right_pt_x < new_left_pt_x:
                    new_right_pt_x = max(line_right_pt[0], right_bound[0])
                coor_to_arrow_dic.pop(arrow_head)
                new_mid_coor = (
                    (new_left_pt_x + new_right_pt_x)/2, line_left_pt[1])
                coor_to_arrow_dic.update({arrow_head: ((new_left_pt_x, line_left_pt[1]),
                                                       abs(new_left_pt_x -
                                                           new_right_pt_x),
                                                       new_mid_coor,
                                                       False,
                                                       number,
                                                       size,
                                                       text_coor)})
# 組合主筋與梁文字


def combine_beam_rebar(coor_to_arrow_dic: dict, coor_to_rebar_list_straight: list, coor_to_bend_rebar_list: list, coor_to_beam_list: list, class_beam_list: list[Beam]):
    # 以箭頭的頭為搜尋中心
    def _get_distance(pt1, pt2):
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    # for arrow_head,arrow_item in coor_to_arrow_dic.items():
    #     tail_coor,length,line_head_coor,number,size,line_tail_coor= arrow_item
    #     try:
    #         number = int(number)
    #     except:
    #         continue
    #     bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=arrow_head)]
    #     if len(bounding_box) == 0:
    #         coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < arrow_head[1]]
    #         if len(coor_sorted_beam_list) == 0:continue
    #         nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],arrow_head))
    #     else:
    #         nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],arrow_head))
    #     nearest_beam[5].append({f'{number}-{size}':length})
    #     if size in nearest_beam[6]:
    #         if 'E.F' in size:
    #             nearest_beam[6][size] += length
    #         else:
    #             nearest_beam[6][size] += int(number)*length
    #     else:
    #         if 'E.F' in size:
    #             nearest_beam[6][size] = length
    #         else:
    #             nearest_beam[6][size] = int(number)*length

    for arrow_head, arrow_item in coor_to_arrow_dic.items():
        tail_coor, length, line_mid_coor, with_dim, number, size, text_coor = arrow_item
        bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=arrow_head)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [
                beam for beam in class_beam_list if beam.coor.y < arrow_head[1]]
            if len(coor_sorted_beam_list) == 0:
                continue
            nearest_beam = min(coor_sorted_beam_list, key=lambda b: _get_distance(
                b.get_coor(), arrow_head))
        else:
            nearest_beam = min(bounding_box, key=lambda b: _get_distance(
                b.get_coor(), arrow_head))
        # if nearest_beam.serial == 'B2-4':
        #     print(f'b8-1:{arrow_head} {arrow_item}')
        nearest_beam.add_rebar(start_pt=line_mid_coor, end_pt=line_mid_coor,
                               length=length, number=number,
                               size=size, text=f'{number}-{size}', arrow_coor=(arrow_head, text_coor), with_dim=with_dim)

    for rebar_line in coor_to_rebar_list_straight:  # (頭座標，尾座標，長度，number，size)
        head_coor, tail_coor, length, number, size, with_dim = rebar_line
        mid_pt = ((head_coor[0] + tail_coor[0])/2,
                  (head_coor[1] + tail_coor[1])/2)
        bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=mid_pt)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [
                beam for beam in class_beam_list if beam.coor.y < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:
                continue
            nearest_beam = min(coor_sorted_beam_list,
                               key=lambda b: _get_distance(b.get_coor(), mid_pt))
        else:
            nearest_beam = min(
                bounding_box, key=lambda b: _get_distance(b.get_coor(), mid_pt))
        nearest_beam.add_rebar(start_pt=mid_pt, end_pt=mid_pt,
                               length=length, number=number,
                               size=size, text=f'{number}-{size}',
                               add_up='straight', arrow_coor=(arrow_head, tail_coor), with_dim=with_dim)

    # for rebar_line in coor_to_rebar_list_straight:# (頭座標，尾座標，長度，number，size)
    #     head_coor,tail_coor,length,number,size= rebar_line
    #     mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
    #     bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=mid_pt)]
    #     if len(bounding_box) == 0:
    #         coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < mid_pt[1]]
    #         if len(coor_sorted_beam_list) == 0:continue
    #         nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],mid_pt))
    #     else:
    #         nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],mid_pt))
    #     nearest_beam[5].append({f'支承端{number}-{size}':length})
    #     if rebar_line[4] in nearest_beam[6]:
    #         nearest_beam[6][size] += int(number)*length
    #     else:
    #         nearest_beam[6][size] = int(number)*length

    # for bend_line in coor_to_bend_rebar_list:# (直的端點，橫的端點，長度，number，size)
    #     head_coor,tail_coor,length,number,size= bend_line
    #     # mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
    #     mid_pt = head_coor
    #     bounding_box = [block for block in coor_to_beam_list if inblock(block=block[4],pt=mid_pt)]
    #     if len(bounding_box) == 0:
    #         coor_sorted_beam_list = [beam for beam in coor_to_beam_list if beam[1][1] < mid_pt[1]]
    #         if len(coor_sorted_beam_list) == 0:continue
    #         nearest_beam = min(coor_sorted_beam_list,key=lambda b:_get_distance(b[1],mid_pt))
    #     else:
    #         nearest_beam  = min(bounding_box,key=lambda b:_get_distance(b[1],mid_pt))
    #     nearest_beam[5].append({f'彎鉤{number}-{size}':length})
    #     if bend_line[4] in nearest_beam[6]:
    #         if 'E.F' in size:
    #             nearest_beam[6][size] += length
    #         else:
    #             nearest_beam[6][size] += int(number)*length
    #     else:
    #         if 'E.F' in size:
    #             nearest_beam[6][size] = length
    #         else:
    #             nearest_beam[6][size] = int(number)*length

    for bend_line in coor_to_bend_rebar_list:  # (直的端點，橫的端點，長度，number，size)
        head_coor, tail_coor, length, number, size = bend_line
        # mid_pt = ((head_coor[0] + tail_coor[0])/2,(head_coor[1] +tail_coor[1])/2)
        mid_pt = head_coor
        bounding_box = [beam for beam in class_beam_list if inblock(
            block=beam.get_bounding_box(), pt=mid_pt)]
        if len(bounding_box) == 0:
            coor_sorted_beam_list = [
                beam for beam in class_beam_list if beam.coor.y < mid_pt[1]]
            if len(coor_sorted_beam_list) == 0:
                print(f'{head_coor} bend rebar no beam')
                continue
            nearest_beam = min(coor_sorted_beam_list,
                               key=lambda b: _get_distance(b.get_coor(), mid_pt))
        else:
            nearest_beam = min(
                bounding_box, key=lambda b: _get_distance(b.get_coor(), mid_pt))
        nearest_beam.add_rebar(start_pt=mid_pt, end_pt=mid_pt,
                               length=length, number=number,
                               size=size, text=f'{number}-{size}',
                               add_up='bend', arrow_coor=(arrow_head, tail_coor), with_dim=with_dim)

# 輸出每隻梁的結果
# def count_each_beam_rebar_tie(coor_to_beam_list:list,output_txt='test.txt'):
#     # (string, midpoint, list of tie, tie_count_dic,(左下，右上),list of rebar,rebar count dict)
#     lines=[]
#     total_tie = {}
#     # total_tie_count = {}
#     total_rebar = {}
#     floor_rebar = {}
#     floor_concrete = {}
#     floor_formwork = {}
#     total_concrete = {}
#     total_formwork={}
#     def _add_total(size,number,total):
#         if size in total:
#             total[size] += number
#         else:
#             total[size] = number

#     for beam in coor_to_beam_list:
#         count_floor = beam[0].split(' ')[0]
#         if count_floor not in floor_rebar:
#             floor_rebar.update({count_floor:{}})
#             floor_concrete.update({count_floor:0})
#             floor_formwork.update({count_floor:0})
#         # temp_dict = floor_rebar[count_floor]
#         matches= re.findall(r"\((.*?)\)",beam[0],re.MULTILINE)
#         # if len(matches) == 0 or 'X' not in matches[0]:return
#         if len(matches) == 0 or len(re.findall(r"X|x",matches[0],re.MULTILINE)) == 0:continue
#         split_char = re.findall(r"X|x",matches[0])[0]
#         tie=0
#         rebar=0
#         total_tie_count = 0
#         try:
#             depth = int(matches[0].split(split_char)[1])
#             width = int(matches[0].split(split_char)[0])
#         except:
#             depth = 0
#             width = 0
#         tie_count = beam[3]
#         rebar_count = beam[6]
#         for size,count in tie_count.items():
#             tie += count * RebarInfo(size) * ((depth - 10)+(width-10))*2
#             total_tie_count += count
#             _add_total(size=size,number=count * RebarInfo(size) * ((depth - 10)+(width-10))*2,total=total_tie)
#             _add_total(size=size,number=count * RebarInfo(size) * ((depth - 10)+(width-10))*2,total=floor_rebar[count_floor])
#             # _add_total(size=size,number=count,total=total_tie_count)
#         for size,length in rebar_count.items():
#             rebar += RebarInfo(size) * length
#             _add_total(size=size,number=RebarInfo(size) * length,total=total_rebar)
#             _add_total(size=size,number=RebarInfo(size) * length,total=floor_rebar[count_floor])
#         lines.append('\n梁{}:'.format(beam[0]))
#         lines.append('\n寬度:{}、高度:{}'.format(width,depth))
#         lines.append('\n主筋為:{}'.format(beam[5]))
#         lines.append('\n箍筋為:{}'.format(beam[2]))
#         lines.append('\n箍筋個數為:{}'.format(total_tie_count))
#         lines.append('\n主筋量為:{}'.format(rebar))
#         lines.append('\n箍筋量為:{}'.format(tie))
#         lines.append(f'==================================')
#     with open(output_txt, 'w',encoding= 'utf-8') as f:
#         for floor,item in floor_rebar.items():
#             lines.append('\n{0} 鋼筋量 :{1}'.format(floor,item))
#         lines.append('\n箍筋總量{}:'.format(total_tie))
#         lines.append('\n主筋總量{}'.format(total_rebar))
#         lines.append('\n混凝土體積已扣除與版共構區域')
#         lines.append('\n模板已扣除與版共構區域')
#         f.write('\n'.join(lines))

# def count_floor_total_beam_rebar_tie(class_to_beam_list:list[Beam],output_txt='test.txt'):
#     lines=[]
#     total_tie = {}
#     # total_tie_count = {}
#     total_rebar = {}
#     floor_rebar = {}
#     floor_concrete = {}
#     floor_formwork = {}
#     total_concrete = 0
#     total_formwork = 0
#     def _add_total(size,number,total):
#         if size in total:
#             total[size] += number
#         else:
#             total[size] = number
#     for beam in class_to_beam_list:
#         matches= re.findall(r"\((.*?)\)",beam.serial,re.MULTILINE)
#         if len(matches) == 0 or len(re.findall(r"X|x",beam.serial,re.MULTILINE)) == 0:continue
#         count_floor = beam.floor
#         if count_floor not in floor_rebar:
#             floor_rebar.update({count_floor:{}})
#             floor_concrete.update({count_floor:0})
#             floor_formwork.update({count_floor:0})
#         floor_concrete[count_floor] += beam.concrete
#         floor_formwork[count_floor] += beam.formwork
#         lines.append('\n梁{}:'.format(beam.serial))
#         lines.append('\n寬度:{}、深度:{}'.format(beam.width,beam.depth))
#         lines.append('\n主筋為:{}'.format(beam.get_rebar_list()))
#         lines.append('\n箍筋為:{}'.format(beam.get_tie_list()))
#         lines.append('\n主筋量(g)為:{}'.format(beam.get_rebar_weight()))
#         lines.append('\n箍筋量(g)為:{}'.format(beam.get_tie_weight()))
#         lines.append(f'==================================')
#         for size,weight in beam.rebar_count.items():
#             _add_total(size=size,number=weight,total=total_rebar)
#             _add_total(size=size,number=weight,total=floor_rebar[count_floor])
#         for size,weight in beam.tie_count.items():
#             _add_total(size=size,number=weight,total=total_tie)
#             _add_total(size=size,number=weight,total=floor_rebar[count_floor])

#     with open(output_txt, 'w',encoding= 'utf-8') as f:
#         for floor,item in floor_rebar.items():
#             total_concrete +=floor_concrete[count_floor]
#             total_formwork +=floor_formwork[count_floor]
#             lines.append('\n{0} 鋼筋量(g) :{1}'.format(floor,item))
#             lines.append('\n{0} 混凝土體積(cm3) :{1}'.format(floor,floor_concrete[count_floor]))
#             lines.append('\n{0} 模板量(cm2) :{1}'.format(floor,floor_formwork[count_floor]))
#             lines.append(f'==================================')
#         lines.append('\n箍筋總量(g):{}'.format(total_tie))
#         lines.append('\n主筋總量(g):{}'.format(total_rebar))
#         lines.append('\n混凝土總體積(cm3):{}'.format(total_concrete))
#         lines.append('\n模板總量(cm2):{}'.format(total_formwork))
#         lines.append('\n混凝土體積已扣除與版共構區域')
#         lines.append('\n模板僅考慮梁兩側及底面')
#         f.write('\n'.join(lines))
#     pass


def count_rebar_in_block(coor_to_arrow_dic: dict, coor_to_block_list: list, coor_to_rebar_list_straight, coor_to_bend_rebar_list):
    # 新的coor_to_arrow_dic為尖點座標 -> (箭頭文字端坐標，鋼筋長度，鋼筋中點座標，數量，尺寸，文字座標)
    for x in coor_to_arrow_dic:
        # 先找在哪個block裡面
        # ((左下，右上), rebar_length_dic, tie_count_dic)
        for y in coor_to_block_list:
            if x[0] > y[0][0][0] and x[0] < y[0][1][0] and x[1] > y[0][0][1] and x[1] < y[0][1][1]:
                # y[1] 是該格的rebar_length_dic: size -> length * number
                if coor_to_arrow_dic[x][4] not in y[1]:
                    y[1][coor_to_arrow_dic[x][4]] = float(
                        coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
                else:
                    y[1][coor_to_arrow_dic[x][4]
                         ] += float(coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
    for rebar_line in coor_to_rebar_list_straight:  # (頭座標，尾座標，長度，number，size)
        # ((左下，右上), rebar_length_dic, tie_count_dic)
        for block in coor_to_block_list:
            if inblock(pt=rebar_line[0], block=block[0]) and inblock(pt=rebar_line[1], block=block[0]):
                if rebar_line[4] not in block[1]:
                    block[1][rebar_line[4]] = float(
                        rebar_line[2]) * float(rebar_line[3])
                else:
                    block[1][rebar_line[4]
                             ] += float(rebar_line[2]) * float(rebar_line[3])
    for rebar_bend in coor_to_bend_rebar_list:  # (直的端點，橫的端點，長度，number，size)
        # ((左下，右上), rebar_length_dic, tie_count_dic)
        for block in coor_to_block_list:
            if inblock(pt=rebar_bend[0], block=block[0]) and inblock(pt=rebar_bend[1], block=block[0]):
                if rebar_bend[4] not in block[1]:
                    block[1][rebar_bend[4]] = float(
                        rebar_bend[2]) * float(rebar_bend[3])
                else:
                    block[1][rebar_bend[4]
                             ] += float(rebar_bend[2]) * float(rebar_bend[3])
# def summary_block_rebar_tie(coor_to_block_list):
#     rebar_length_dic = {}
#     tie_count_dic = {}
#     output_txt_2 = ''
#     with open(output_txt_2, 'w',encoding= 'utf-8') as f:
#     # f = open(output_txt, "w", encoding = 'utf-8')

#         for x in coor_to_block_list:
#             if len(x[1]) != 0 or len(x[2]) != 0:
#                 f.write(f'統計左下角為{x[0][0]}，右上角為{x[0][1]}的框框內結果：\n')
#                 if len(x[1]) != 0:
#                     f.write('鋼筋計算：\n')
#                     for y in x[1]:
#                         f.write(f'{y}: 總長度(長度*數量)為 {x[1][y]}\n')
#                         if y in rebar_length_dic:
#                             rebar_length_dic[y] += x[1][y] * RebarInfo(y)
#                         else:
#                             rebar_length_dic[y] = x[1][y] * RebarInfo(y)
#                 else:
#                     f.write('此圖框內沒有鋼筋\n')

#                 if len(x[2]) != 0:
#                     f.write('箍筋計算：\n')
#                     for y in x[2]:
#                         f.write(f'{y}: 總數量為 {x[2][y]}\n')
#                         if y in tie_count_dic:
#                             tie_count_dic[y] += x[2][y]
#                         else:
#                             tie_count_dic[y] = x[2][y]
#                 else:
#                     f.write('此圖框內沒有箍筋\n')

#                 f.write('\n')

#         f.write(f'統計所有結果：\n')
#         f.write('鋼筋計算：\n')
#         for y in rebar_length_dic:
#             f.write(f'{y}: 總長度(長度*數量)為 {rebar_length_dic[y]}\n')

#         f.write('箍筋計算：\n')
#         for y in tie_count_dic:
#             f.write(f'{y}: 總數量為 {tie_count_dic[y]}\n')
#     pass


def inblock(block: tuple, pt: tuple):
    pt_x = pt[0]
    pt_y = pt[1]
    if len(block) == 0:
        return False
    if (pt_x - block[0][0])*(pt_x - block[1][0]) < 0 and (pt_y - block[0][1])*(pt_y - block[1][1]) < 0:
        return True
    return False


def cal_beam_rebar(data={}, progress_file='', rebar_parameter_excel=''):
    # output_txt = f'{output_folder}{project_name}'
    progress(f'================start cal_beam_rebar================', progress_file)
    if not data:
        return
    cad_data = {}
    coor_to_rebar_list = data['coor_to_rebar_list']  # (頭座標，尾座標，長度)
    coor_to_bend_rebar_list = data['coor_to_bend_rebar_list']  # (直的端點，橫的端點，長度)
    coor_to_data_list = data['coor_to_data_list']  # (字串，座標)
    coor_to_arrow_dic = data['coor_to_arrow_dic']  # 尖點座標 -> 文字連接處座標
    coor_to_tie_list = data['coor_to_tie_list']  # (下座標，上座標，長度)
    coor_to_tie_text_list = data['coor_to_tie_text_list']  # (字串，座標)
    # ((左下，右上), rebar_length_dic, tie_count_dic)
    coor_to_block_list = data['coor_to_block_list']
    # (string, midpoint, list of tie, tie_count_dic)
    coor_to_beam_list = data['coor_to_beam_list']
    coor_to_bounding_block_list = data['coor_to_bounding_block_list']
    coor_to_rc_block_list = data['coor_to_rc_block_list']
    coor_to_dim_list = data['coor_to_dim_list']
    cad_data = {
        '直線鋼筋': len(coor_to_rebar_list),
        '彎鉤鋼筋': len(coor_to_bend_rebar_list),
        '主筋文字': len(coor_to_data_list),
        '鋼筋指示線': len(coor_to_arrow_dic),
        '箍筋': len(coor_to_tie_list),
        '箍筋文字': len(coor_to_tie_text_list),
        '圖框': len(coor_to_block_list),
        '梁編號': len(coor_to_beam_list),
        '梁圖框': len(coor_to_bounding_block_list)
    }
    class_beam_list = []

    readRebarExcel(file_path=rebar_parameter_excel)

    # 2023-0505 add floor xlsx to found floor
    parameter_df = read_parameter_df(rebar_parameter_excel, '梁參數表')
    floor_list = parameter_df['樓層'].tolist()

    def fix_floor_list(floor):
        if floor[-1] != 'F':
            floor += 'F'
        return floor
    floor_list = list(map(lambda f: fix_floor_list(f), floor_list))

    # parameter_df.set_index(['樓層'],inplace=True)
    # Step 8. 對應箭頭跟鋼筋
    start = time.time()
    coor_to_arrow_dic, no_arrow_line_list = sort_arrow_line(coor_to_arrow_dic,
                                                            coor_to_rebar_list,
                                                            coor_to_dim_list=coor_to_dim_list,
                                                            coor_to_bend_rebar_list=coor_to_bend_rebar_list,
                                                            progress_file=progress_file)
    progress(f'整理線段與鋼筋標示箭頭:{time.time() - start}s', progress_file)

    # Step 9. 對應箭頭跟文字，並完成head_to_data_dic, tail_to_data_dic
    start = time.time()
    try:
        coor_to_arrow_dic, head_to_data_dic, tail_to_data_dic = sort_arrow_to_word(coor_to_arrow_dic=coor_to_arrow_dic,
                                                                                   coor_to_data_list=coor_to_data_list,
                                                                                   progress_file=progress_file)
    except NoRebarDataError:
        print('NoRebarData')
    progress(f'整理鋼筋文字與鋼筋標示箭頭:{time.time() - start}s', progress_file)

    # Step 10. 統計目前的type跟size

    # progress('梁配筋圖讀取進度 10/15', progress_file)
    # coor_to_rebar_list_straight_left,coor_to_rebar_list_straight_right, coor_to_bend_rebar_list,no_concat_line_list,no_concat_bend_list=concat_no_arrow_line(no_arrow_line_list=no_arrow_line_list,
    #                                                                                                                 head_to_data_dic=head_to_data_dic,
    #                                                                                                                 tail_to_data_dic=tail_to_data_dic,
    #                                                                                                                 coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 12. 拿彎的去找跟誰接在一起
    start = time.time()
    coor_to_rebar_list_straight = sort_noconcat_line(no_concat_line_list=no_arrow_line_list,
                                                     head_to_data_dic=head_to_data_dic,
                                                     tail_to_data_dic=tail_to_data_dic)
    coor_to_bend_rebar_list = sort_noconcat_bend(no_concat_bend_list=coor_to_bend_rebar_list,
                                                 head_to_data_dic=head_to_data_dic,
                                                 tail_to_data_dic=tail_to_data_dic)
    sort_rebar_bend_line(rebar_bend_list=coor_to_bend_rebar_list,
                         rebar_line_list=coor_to_rebar_list_straight)
    # 截斷處重複計算
    progress(f'整理彎折鋼筋與直線鋼筋:{time.time() - start}s', progress_file)

    # Step 14-15 和 16 為箍筋部分，14-15在算框框內的數量，16在算每個梁的總長度，兩者獨立
    # count_rebar_in_block(coor_to_arrow_dic,coor_to_block_list,coor_to_rebar_list_straight=coor_to_rebar_list_straight,coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 14. 算箍筋
    start = time.time()
    coor_sorted_tie_list = count_tie(coor_to_tie_text_list=coor_to_tie_text_list,
                                     coor_to_block_list=coor_to_block_list, coor_to_tie_list=coor_to_tie_list)
    add_beam_to_list(coor_to_beam_list=coor_to_beam_list,
                     class_beam_list=class_beam_list,
                     floor_list=floor_list)
    progress(f'新增梁:{time.time() - start}s', progress_file)
    start = time.time()
    combine_beam_boundingbox(coor_to_block_list=coor_to_block_list,
                             coor_to_bounding_block_list=coor_to_bounding_block_list,
                             class_beam_list=class_beam_list,
                             coor_to_rc_block_list=coor_to_rc_block_list)
    progress(f'整理梁邊界與邊框:{time.time() - start}s', progress_file)
    start = time.time()
    break_down_rebar(coor_to_arrow_dic=coor_to_arrow_dic,
                     class_beam_list=class_beam_list)
    progress(f'截斷直線鋼筋:{time.time() - start}s', progress_file)
    start = time.time()
    combine_beam_tie(coor_sorted_tie_list=coor_sorted_tie_list,
                     coor_to_beam_list=coor_to_beam_list,
                     class_beam_list=class_beam_list)
    combine_beam_rebar(coor_to_arrow_dic=coor_to_arrow_dic,
                       coor_to_rebar_list_straight=coor_to_rebar_list_straight,
                       coor_to_bend_rebar_list=coor_to_bend_rebar_list,
                       coor_to_beam_list=coor_to_beam_list,
                       class_beam_list=class_beam_list)
    progress(f'配對梁與主筋、箍筋:{time.time() - start}s', progress_file)
    start = time.time()
    compare_line_with_dim(class_beam_list=class_beam_list,
                          coor_to_dim_list=coor_to_dim_list,
                          coor_to_block_list=coor_to_block_list,
                          progress_file=progress_file)
    progress(f'配對梁主筋與標註線:{time.time() - start}s', progress_file)
    start = time.time()
    sort_beam(class_beam_list=class_beam_list)
    progress(f'整理梁配筋:{time.time() - start}s', progress_file)
    return class_beam_list, cad_data


def create_report(class_beam_list: list[Beam],
                  output_folder: str,
                  project_name: str,
                  floor_parameter_xlsx: str,
                  cad_data: Counter,
                  progress_file: str):
    progress(f'產生報表', progress_file)
    excel_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Count.xlsx'
    )
    excel_filename_rcad = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Rcad.xlsx'
    )
    output_file_list = []
    pdf_GB_file = ''
    pdf_FB_file = ''
    pdf_SB_file = ''

    cad_df = pd.DataFrame.from_dict(
        data=cad_data, orient='index', columns=['數量'])
    floor_list = floor_parameter(
        beam_list=class_beam_list, floor_parameter_xlsx=floor_parameter_xlsx)
    beam_df = output_beam(class_beam_list=class_beam_list)
    fbeam_list, sbeam_list, beam_list = seperate_beam(
        class_beam_list=class_beam_list)

    bs_list = create_beam_scan()
    sb_bs_list = create_sbeam_scan()
    fb_bs_list = create_fbeam_scan()
    # if beam_list:
    #     pdf_GB_file = create_scan_report(floor_list=floor_list,
    #                                      beam_list=beam_list,
    #                                      bs_list=bs_list,
    #                                      excel_filename=excel_filename,
    #                                      beam_type=BeamType.Grider,
    #                                      project_name=project_name,
    #                                      output_folder=output_folder)
    #     output_file_list.append(pdf_GB_file)
    # if fbeam_list:
    #     pdf_FB_file = create_scan_report(floor_list=floor_list,
    #                                      beam_list=fbeam_list,
    #                                      bs_list=fb_bs_list,
    #                                      excel_filename=excel_filename,
    #                                      beam_type=BeamType.FB,
    #                                      project_name=project_name,
    #                                      output_folder=output_folder)
    #     output_file_list.append(pdf_FB_file)
    if sbeam_list:
        pdf_SB_file = create_scan_report(floor_list=floor_list,
                                         beam_list=sbeam_list,
                                         bs_list=sb_bs_list,
                                         excel_filename=excel_filename,
                                         beam_type=BeamType.SB,
                                         project_name=project_name,
                                         output_folder=output_folder)
        output_file_list.append(pdf_SB_file)
    rebar_df, concrete_df, coupler_df, formwork_df = summary_floor_rebar(
        floor_list=floor_list, item_type='beam')
    # header_list,ratio_dict,ratio_df = summary_floor_rebar_ratio(floor_list=floor_list)

    # enoc_df,code_df = beam_check(beam_list=beam_list,beam_scan_list=bs_list)
    # sb_enoc_df, sb_code_df= beam_check(beam_list=sbeam_list,beam_scan_list=sb_bs_list)
    # fb_enoc_df, fb_code_df= beam_check(beam_list=fbeam_list,beam_scan_list=fb_bs_list)
    # ng_df = output_detail_scan_report(beam_list=beam_list + sbeam_list + fbeam_list)

    rcad_df = output_rcad_beam(class_beam_list=class_beam_list)
    # beam_ng_df,sum_df = output_ng_ratio(code_df)
    # sbeam_ng_df,sb_sum_df = output_ng_ratio(sb_code_df)
    # fbeam_ng_df,fb_sum_df = output_ng_ratio(fb_code_df)
    # OutputExcel(df_list=[code_df,enoc_df],file_path=excel_filename,sheet_name='梁檢核表',auto_fit_columns=[1],auto_fit_rows=[1],
    #     columns_list=range(2,len(code_df.columns)+2),rows_list=range(2,len(code_df.index)+ 10 + len(enoc_df.index)),df_spacing= 3)
    # OutputExcel(df_list=[sb_code_df,sb_enoc_df],file_path=excel_filename,sheet_name='小梁檢核表',auto_fit_columns=[1],auto_fit_rows=[1],
    #     columns_list=range(2,len(sb_code_df.columns)+2),rows_list=range(2,len(sb_code_df.index)+ 10 + len(sb_enoc_df.index)),df_spacing= 3)
    # OutputExcel(df_list=[fb_code_df,fb_enoc_df],file_path=excel_filename,sheet_name='地梁檢核表',auto_fit_columns=[1],auto_fit_rows=[1],
    #     columns_list=range(2,len(fb_code_df.columns)+2),rows_list=range(2,len(fb_code_df.index)+ 10 + len(fb_enoc_df.index)),df_spacing= 3)
    OutputExcel(df_list=[beam_df], file_path=excel_filename, sheet_name='梁統整表')
    # Add_Row_Title(file_path=excel_filename,sheet_name='梁檢核表',i=len(code_df.index) + 4,j=1,title_text='經濟性檢核',font_size= 20)
    # Add_Row_Title(file_path=excel_filename,sheet_name='小梁檢核表',i=len(sb_code_df.index) + 4,j=1,title_text='經濟性檢核',font_size= 20)
    # Add_Row_Title(file_path=excel_filename,sheet_name='地梁檢核表',i=len(fb_code_df.index) + 4,j=1,title_text='經濟性檢核',font_size= 20)
    OutputExcel(df_list=[rebar_df],
                file_path=excel_filename, sheet_name='鋼筋統計表')
    OutputExcel(df_list=[concrete_df],
                file_path=excel_filename, sheet_name='混凝土統計表')
    OutputExcel(df_list=[formwork_df],
                file_path=excel_filename, sheet_name='模板統計表')
    # OutputExcel(df_list=[ng_df],file_path=excel_filename,sheet_name='詳細檢核表')
    # OutputExcel(df_list=[ratio_df],
    #             file_path=excel_filename,
    #             sheet_name='鋼筋比統計表',
    #             columns_list=range(2,len(ratio_df.columns) + 2),
    #             rows_list=range(2,len(ratio_df.index) + 2)
    #             )
    # AddExcelDataBar(workbook_path=excel_filename,
    #                 sheet_name='鋼筋比統計表',
    #                 start_col=4,
    #                 start_row=4,
    #                 end_col=len(ratio_df.columns) + 4,
    #                 end_row=len(ratio_df.index) + 4)
    # AddBorderLine(workbook_path=excel_filename,
    #                 sheet_name='鋼筋比統計表',
    #                 start_col=4,
    #                 start_row=4,
    #                 end_col=len(ratio_df.columns) + 4,
    #                 end_row=len(ratio_df.index) + 4,
    #                 step_row= 2,
    #                 step_col= 3)
    OutputExcel(df_list=[cad_df], file_path=excel_filename,
                sheet_name='CAD統計表')
    OutputExcel(df_list=[rcad_df],
                file_path=excel_filename_rcad, sheet_name='RCAD撿料')
    output_file_list.append(excel_filename)
    output_file_list.append(excel_filename_rcad)
    # return excel_filename,excel_filename_rcad,pdf_FB_file,pdf_GB_file,pdf_SB_file
    return output_file_list


def create_scan_report(floor_list: list[Floor],
                       beam_list: list[Beam],
                       bs_list: list,
                       excel_filename: str,
                       beam_type: BeamType,
                       project_name: str,
                       output_folder: str):
    '''
    Create scan pdf report with beam list \n
    Args:
        floor_list:list of Floor Object
        beam_list:list of Beam object
        bs_list:list of BeamScan Object
    '''
    if beam_type == BeamType.FB:
        item_name = '地梁'
    if beam_type == BeamType.Grider:
        item_name = '大梁'
    if beam_type == BeamType.SB:
        item_name = '小梁'

    pdf_report = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'{item_name}_report.pdf'
    )

    enoc_df, code_df = beam_check(beam_list=beam_list, beam_scan_list=bs_list)
    header_list, ratio_dict, ratio_df = summary_floor_rebar_ratio(floor_list=floor_list,
                                                                  beam_type=beam_type)
    rebar_df, concrete_df, coupler_df, formwork_df = summary_floor_rebar(floor_list=floor_list,
                                                                         item_type='beam',
                                                                         beam_type=beam_type)
    ng_df = output_detail_scan_report(beam_list=beam_list)
    beam_ng_df, sum_df = output_ng_ratio(code_df)
    create_scan_pdf(scan_list=bs_list,
                    scan_df=ng_df.copy(),
                    rebar_df=rebar_df,
                    concrete_df=concrete_df,
                    formwork_df=formwork_df,
                    ng_sum_df=sum_df,
                    beam_ng_df=beam_ng_df,
                    project_prop={
                        "專案名稱:": f'{project_name}_{item_name}',
                        "測試日期:": time.strftime("%Y/%m/%d %H:%M:%S", time.localtime())
                    },
                    pdf_filename=pdf_report,
                    header_list=header_list,
                    ratio_dict=ratio_dict,
                    report_type='beam',
                    item_name=item_name)
    OutputExcel(df_list=[code_df, enoc_df], file_path=excel_filename,
                sheet_name=f'{item_name}檢核表', auto_fit_columns=[1], auto_fit_rows=[1],
                columns_list=range(2, len(code_df.columns)+2),
                rows_list=range(2, len(code_df.index) +
                                10 + len(enoc_df.index)),
                df_spacing=3)
    Add_Row_Title(file_path=excel_filename, sheet_name=f'{item_name}檢核表',
                  i=len(code_df.index) + 4, j=1,
                  title_text='經濟性檢核', font_size=20)
    OutputExcel(df_list=[ng_df], file_path=excel_filename,
                sheet_name=f'{item_name}詳細檢核表')
    OutputExcel(df_list=[rebar_df], file_path=excel_filename,
                sheet_name=f'{item_name}鋼筋統計表')
    OutputExcel(df_list=[ratio_df],
                file_path=excel_filename,
                sheet_name=f'{item_name}鋼筋比統計表',
                columns_list=range(2, len(ratio_df.columns) + 2),
                rows_list=range(2, len(ratio_df.index) + 2)
                )
    AddExcelDataBar(workbook_path=excel_filename,
                    sheet_name=f'{item_name}鋼筋比統計表',
                    start_col=4,
                    start_row=4,
                    end_col=len(ratio_df.columns) + 4,
                    end_row=len(ratio_df.index) + 4)
    AddBorderLine(workbook_path=excel_filename,
                  sheet_name=f'{item_name}鋼筋比統計表',
                  start_col=4,
                  start_row=4,
                  end_col=len(ratio_df.columns) + 4,
                  end_row=len(ratio_df.index) + 4,
                  step_row=2,
                  step_col=3)
    return pdf_report


def seperate_beam(class_beam_list: list[Beam]):
    '''
    Seperate Beam into GB,SB,FB
    '''
    return [b for b in class_beam_list if b.beam_type == BeamType.FB],\
        [b for b in class_beam_list if b.beam_type == BeamType.SB],\
        [b for b in class_beam_list if b.beam_type == BeamType.Grider]


def add_beam_to_list(coor_to_beam_list: list, class_beam_list: list, floor_list: list):
    floor_pattern = r'(\d+F)|(R\d+)|(PR)|(B\d+)|(MF)'
    for beam in coor_to_beam_list:
        try:
            b = Beam(beam[0], beam[1][0], beam[1][1])
            b.get_beam_info(floor_list=floor_list)
        except BeamFloorNameError:
            print(f'{beam[0]} beam serial error')
            continue
        if re.search(floor_pattern, b.floor) and b.serial != '':
            # print(f'{b.floor} {b.serial} accept')
            class_beam_list.append(b)
        else:
            pass
            # print(b.floor)
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


def draw_rebar_line(class_beam_list: list[Beam], msp_beam: object, doc_beam: object, output_folder: str, project_name: str):
    error_count = 0
    date = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    output_dwg = os.path.join(
        output_folder, f'{project_name}_{date}_Markon.dwg')
    while error_count <= 10:
        try:
            layer_beam = doc_beam.Layers.Add(f"S-CLOUD_{date}")
            doc_beam.ActiveLayer = layer_beam
            layer_beam.color = 10
            layer_beam.Linetype = "Continuous"
            layer_beam.Lineweight = 0.5
            break
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'write_beam error in step 4, {e}, error_count = {error_count}.')

    for beam in class_beam_list:
        for pos, rebar_list in beam.rebar.items():
            for rebar in rebar_list:
                error_count = 0
                while error_count < 3:
                    try:
                        x = rebar.start_pt.x
                        y = rebar.start_pt.y
                        coor_list1 = [beam.coor.x, beam.coor.y, 0, x, y, 0]
                        coor_list2 = [beam.coor.x, beam.coor.y,
                                      0, rebar.end_pt.x, rebar.end_pt.y, 0]
                        points1 = vtFloat(coor_list1)
                        points2 = vtFloat(coor_list2)
                        line1 = msp_beam.AddPolyline(points1)
                        line2 = msp_beam.AddPolyline(points2)
                        text1 = msp_beam.AddMText(
                            vtPnt((x + rebar.end_pt.x)/2, (y + rebar.end_pt.y)/2), 10, rebar.text)
                        text1.Height = 5
                        line1.SetWidth(0, 1, 1)
                        line2.SetWidth(0, 1, 1)
                        line1.color = 101
                        line2.color = 101
                        break
                    except Exception:
                        print('write rebar error')
                        error_count += 1
                        time.sleep(5)
        for pos, tie in beam.tie.items():
            if (not tie):
                continue
            error_count = 0
            while error_count < 3:
                try:
                    coor_list1 = [beam.coor.x, beam.coor.y,
                                  0, tie.start_pt.x, tie.start_pt.y, 0]
                    points1 = vtFloat(coor_list1)
                    line1 = msp_beam.AddPolyline(points1)
                    line1.SetWidth(0, 1, 1)
                    line1.color = 1
                    break
                except Exception:
                    print('write tie error')
                    error_count += 1
                    time.sleep(5)
        for middle_tie in beam.middle_tie:
            error_count = 0
            while error_count < 3:
                try:
                    coor_list1 = [beam.coor.x, beam.coor.y, 0,
                                  middle_tie.start_pt.x, middle_tie.start_pt.y, 0]
                    points1 = vtFloat(coor_list1)
                    line1 = msp_beam.AddPolyline(points1)
                    line1.SetWidth(0, 1, 1)
                    line1.color = 240
                    break
                except Exception:
                    print('write middle tie error')
                    error_count += 1
                    time.sleep(5)
        try:
            beam_bounding_box = beam.get_bounding_box()
            block_list = [beam_bounding_box[0][0], beam_bounding_box[0][1], 0,
                          beam_bounding_box[1][0], beam_bounding_box[0][1], 0,
                          beam_bounding_box[1][0], beam_bounding_box[1][1], 0,
                          beam_bounding_box[0][0], beam_bounding_box[1][1], 0,
                          beam_bounding_box[0][0], beam_bounding_box[0][1], 0,]
            points1 = vtFloat(block_list)
            block = msp_beam.AddPolyline(points1)
            block.SetWidth(0, 1, 1)
            block.color = 50
        except Exception:
            print('write block error')
            error_count += 1
            time.sleep(5)
            # coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
        # coor_list2 = [min_right_coor[0], min_right_coor[1], 0, x[1][0], x[1][1], 0]
    try:
        doc_beam.SaveAs(output_dwg)
        doc_beam.Close(SaveChanges=True)
    except:
        print(project_name)
    return output_dwg


def sort_beam(class_beam_list: list[Beam]):
    for beam in class_beam_list:
        beam.sort_beam_rebar()
        beam.sort_beam_tie()
    for beam in class_beam_list[:]:
        for floor_text in beam.multi_floor[1:]:
            new_beam = copy.deepcopy(beam)
            new_beam.floor = floor_text
            class_beam_list.append(new_beam)
    for beam in class_beam_list[:]:
        for serial_text in beam.multi_serial[1:]:
            new_beam = copy.deepcopy(beam)
            new_beam.serial = serial_text
            class_beam_list.append(new_beam)
    return


def output_beam(class_beam_list: list[Beam]):
    def output_rebar_pos(row, i, pos, text):
        if abs(i) >= 2:
            beam.at[row + i, pos] = f'{beam.at[row + i,pos]}+{text}'
        else:
            beam.at[row + i, pos] = text
    header_info_1 = [('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', '')]
    header_rebar = [('主筋', ''), ('主筋', '左'), ('主筋', '中'),
                    ('主筋', '右'), ('主筋長度', '左'), ('主筋長度', '中'), ('主筋長度', '右')]

    header_sidebar = [('腰筋', '')]

    header_stirrup = [
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')
    ]

    header_info_2 = [
        ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'),
        ('主筋量', 'g'), ('箍筋量', 'g'), ('模板', 'cm2'), ('混凝土', 'cm3'),
        ('搭接長度', '左端'), ('搭接長度', '中央'), ('搭接長度', '右端')
    ]
    header = pd.MultiIndex.from_tuples(
        header_info_1 + header_rebar + header_sidebar + header_stirrup + header_info_2)
    beam = pd.DataFrame(
        np.empty([len(class_beam_list)*4, len(header)], dtype='<U16'), columns=header)

    # beam = pd.DataFrame(np.empty([len(etabs_design.groupby(['Story', 'BayID'])) * 4, len(header)], dtype='<U16'), columns=header)
    row = 0
    # temp_x = 0
    rebar_pos = {
        RebarType.Top: 0,
        RebarType.Bottom: 3,
        RebarType.Left: '左端',
        RebarType.Middle: '中央',
        RebarType.Right: '右端'
    }
    # min_diff = 10
    for b in class_beam_list:
        try:
            b.cal_rebar()
            b.sort_rebar_table()
            beam.at[row, ('樓層', '')] = b.floor
            beam.at[row, ('編號', '')] = b.serial
            beam.at[row, ('RC 梁寬', '')] = b.width
            beam.at[row, ('RC 梁深', '')] = b.depth
            beam.at[row, ('腰筋', '')] = b.get_middle_tie()
            for i, rebar in enumerate(b.rebar_table[RebarType.Top.value][RebarType.Left.value], start=0):
                output_rebar_pos(row, i, ('主筋', '左'), rebar.text)
                # beam.at[row + i,('主筋', '左')] = rebar.text
            for i, rebar in enumerate(b.rebar_table[RebarType.Bottom.value][RebarType.Left.value], start=0):
                output_rebar_pos(row + 3, -i, ('主筋', '左'), rebar.text)
                # beam.at[row + 3 - i,('主筋', '左')] = rebar.text
            for i, rebar in enumerate(b.rebar_table[RebarType.Top.value][RebarType.Middle.value], start=0):
                output_rebar_pos(row, i, ('主筋', '中'), rebar.text)
                # beam.at[row + i,('主筋', '中')] = rebar.text
            for i, rebar in enumerate(b.rebar_table[RebarType.Bottom.value][RebarType.Middle.value], start=0):
                output_rebar_pos(row + 3, -i, ('主筋', '中'), rebar.text)
            for i, rebar in enumerate(b.rebar_table[RebarType.Top.value][RebarType.Right.value], start=0):
                output_rebar_pos(row, i, ('主筋', '右'), rebar.text)
                # beam.at[row + i,('主筋', '右')] = rebar.text
            for i, rebar in enumerate(b.rebar_table[RebarType.Bottom.value][RebarType.Right.value], start=0):
                output_rebar_pos(row + 3, -i, ('主筋', '右'), rebar.text)
                # beam.at[row + 3 - i,('主筋', '右')] = rebar.text
            if b.rebar_table['top_length'][RebarType.Left.value]:
                beam.at[row, ('主筋長度', '左')
                        ] = b.rebar_table['top_length'][RebarType.Left.value][0]
            if b.rebar_table['top_length'][RebarType.Middle.value]:
                beam.at[row, ('主筋長度', '中')
                        ] = b.rebar_table['top_length'][RebarType.Middle.value][0]
            if b.rebar_table['top_length'][RebarType.Right.value]:
                beam.at[row, ('主筋長度', '右')
                        ] = b.rebar_table['top_length'][RebarType.Right.value][0]
            if b.rebar_table['bottom_length'][RebarType.Left.value]:
                beam.at[row + 3, ('主筋長度', '左')
                        ] = b.rebar_table['bottom_length'][RebarType.Left.value][0]
            if b.rebar_table['bottom_length'][RebarType.Middle.value]:
                beam.at[row + 3, ('主筋長度', '中')
                        ] = b.rebar_table['bottom_length'][RebarType.Middle.value][0]
            if b.rebar_table['bottom_length'][RebarType.Right.value]:
                beam.at[row + 3, ('主筋長度', '右')
                        ] = b.rebar_table['bottom_length'][RebarType.Right.value][0]
            # for rebar_text,rebar_list in b.rebar.items():
            #     for rebar in rebar_list:
            #         if abs(rebar.start_pt.x - b.start_pt.x) < min_diff :
            #             beam.at[row + rebar_pos[rebar_text],('主筋', '左')] = rebar.text
            #         if abs(rebar.end_pt.x - b.end_pt.x)< min_diff:
            #             beam.at[row + rebar_pos[rebar_text],('主筋', '右')] = rebar.text
            #         if (abs(rebar.start_pt.x - b.start_pt.x) >= min_diff and abs(rebar.end_pt.x - b.end_pt.x)>= min_diff) or (rebar.start_pt.x == b.start_pt.x and rebar.end_pt.x == b.end_pt.x):
            #             beam.at[row + rebar_pos[rebar_text],('主筋', '中')] = rebar.text
            #     for rebar in rebar_list:
            #         if abs(rebar.start_pt.x - b.start_pt.x) < min_diff:
            #             beam.at[row + rebar_pos[rebar_text],('主筋長度', '左')] = rebar.length
            #             continue
            #         if abs(rebar.end_pt.x - b.end_pt.x)< min_diff:
            #             beam.at[row + rebar_pos[rebar_text],('主筋長度', '右')] = rebar.length
            #             continue
            #         if (abs(rebar.start_pt.x - b.start_pt.x) >= min_diff and abs(rebar.end_pt.x - b.end_pt.x)>= min_diff):
            #             beam.at[row + rebar_pos[rebar_text],('主筋長度', '中')] = rebar.length
            #             continue
            # for tie_text,tie in b.tie.items():
            for pos, ld in b.ld_table.items():
                beam.at[row + rebar_pos[pos[0]],
                        ('搭接長度', rebar_pos[pos[1]])] = ld
            if b.tie_list:
                beam.at[row, ('箍筋', '左')] = b.tie['left'].text
                beam.at[row, ('箍筋', '中')] = b.tie['middle'].text
                beam.at[row, ('箍筋', '右')] = b.tie['right'].text
                beam.at[row, ('箍筋長度', '左')] = b.length/4
                beam.at[row, ('箍筋長度', '中')] = b.length/2
                beam.at[row, ('箍筋長度', '右')] = b.length/4

            beam.at[row, ('主筋量', 'g')] = b.get_rebar_weight()
            beam.at[row, ('箍筋量', 'g')] = b.get_tie_weight()
            beam.at[row, ('模板', 'cm2')] = b.get_formwork()
            beam.at[row, ('混凝土', 'cm3')] = b.get_concrete()
        except Exception as e:
            # raise
            print(f'{b.floor}:{b.serial}')
            print(e)
            pass
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


def output_rcad_beam(class_beam_list: list[Beam]):

    header_info_1 = [('樓層', ''), ('梁編號', ''), ('型式', ''), ('梁尺寸', '樑寬B '), ('梁尺寸', '樑深H '),
                     ('上層筋', '左'), ('上層筋', '中'), ('上層筋',
                                                  '右'), ('下層筋', '左'), ('下層筋', '中'), ('下層筋', '右'),
                     ('腰筋', '單側'), ('箍筋', '左'), ('箍筋',
                                                 '中'), ('箍筋', '右'), ('垂直繫筋', ''),
                     ('左斷筋點', 'lcut_neg'), ('左斷筋點', 'lcut_pos'), ('右斷筋點', 'lcut_neg'), ('右斷筋點', 'lcut_pos')]
    header = pd.MultiIndex.from_tuples(header_info_1)
    rcad_beam = pd.DataFrame(
        np.empty([len(class_beam_list), len(header)], dtype='<U16'), columns=header)
    row = 1
    for b in class_beam_list:
        left_neg = 0
        left_pos = 0
        right_neg = 0
        right_pos = 0
        rcad_beam.at[row, ('樓層', '')] = b.floor
        rcad_beam.at[row, ('梁編號', '')] = b.serial
        rcad_beam.at[row, ('型式', '')] = 0
        rcad_beam.at[row, ('梁尺寸', '樑寬B ')] = f'{b.width}.'
        rcad_beam.at[row, ('梁尺寸', '樑深H ')] = f'{b.depth}.'
        rcad_beam.at[row, ('上層筋', '左')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Top.value][RebarType.Left.value])))
        rcad_beam.at[row, ('上層筋', '中')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Top.value][RebarType.Middle.value])))
        rcad_beam.at[row, ('上層筋', '右')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Top.value][RebarType.Right.value])))
        rcad_beam.at[row, ('下層筋', '左')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Bottom.value][RebarType.Left.value])))
        rcad_beam.at[row, ('下層筋', '中')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Bottom.value][RebarType.Middle.value])))
        rcad_beam.at[row, ('下層筋', '右')] = ','.join(list(map(lambda r: r.text.replace(
            '-', ''), b.rebar_table[RebarType.Bottom.value][RebarType.Right.value])))
        rcad_beam.at[row, ('腰筋', '單側')] = b.middle_tie[0].text.replace(
            '-', '').replace('E.F', '') if b.middle_tie else '0'
        rcad_beam.at[row, ('箍筋', '左')
                     ] = b.tie[RebarType.Left.value].text if b.tie[RebarType.Left.value] else ''
        rcad_beam.at[row, ('箍筋', '中')
                     ] = b.tie[RebarType.Middle.value].text if b.tie[RebarType.Middle.value] else ''
        rcad_beam.at[row, ('箍筋', '右')
                     ] = b.tie[RebarType.Right.value].text if b.tie[RebarType.Right.value] else ''
        if b.rebar_table["top_length"][RebarType.Left.value]:
            left_neg = ceil(
                b.rebar_table["top_length"][RebarType.Left.value][0])
        if b.rebar_table["bottom_length"][RebarType.Left.value]:
            left_pos = ceil(
                b.rebar_table["bottom_length"][RebarType.Left.value][0])
        if b.rebar_table["top_length"][RebarType.Right.value]:
            right_neg = ceil(
                b.rebar_table["top_length"][RebarType.Right.value][0])
        if b.rebar_table["bottom_length"][RebarType.Right.value]:
            right_pos = ceil(
                b.rebar_table["bottom_length"][RebarType.Right.value][0])
        rcad_beam.at[row, ('左斷筋點', 'lcut_neg')] = f'{left_neg}.,{left_pos}.'
        rcad_beam.at[row, ('左斷筋點', 'lcut_pos')] = f'{left_neg}.,{left_pos}.'
        rcad_beam.at[row, ('右斷筋點', 'lcut_neg')] = f'{right_neg}.,{right_pos}.'
        rcad_beam.at[row, ('右斷筋點', 'lcut_pos')] = f'{right_neg}.,{right_pos}.'
        row += 1
    return rcad_beam


def count_beam_main(beam_filename, layer_config, temp_file='temp_1221_1F.pkl', output_folder='', project_name='', template_name=''):
    progress_file = './result/tmp'
    start = time.time()
    msp_beam, doc_beam = read_beam_cad(
        beam_filename=beam_filename, progress_file=progress_file)
    sort_beam_cad(msp_beam=msp_beam, layer_config=layer_config, entity_config=get_template(
        template_name), progress_file=progress_file, temp_file=temp_file)
    output_txt, output_txt_2, output_excel, class_beam_list = cal_beam_rebar(data=save_temp_file.read_temp(
        temp_file), output_folder=output_folder, project_name=project_name, progress_file=progress_file)
    output_dwg = draw_rebar_line(class_beam_list=class_beam_list, msp_beam=msp_beam,
                                 doc_beam=doc_beam, output_folder=output_folder, project_name=project_name)
    print(f'Total Time:{time.time() - start}')
    return os.path.basename(output_txt), os.path.basename(output_txt_2), os.path.basename(output_excel), os.path.basename(output_dwg)


def count_beam_multiprocessing(beam_filenames: list,
                               layer_config: dict,
                               temp_file='temp_1221_1F.pkl',
                               output_folder='',
                               project_name='',
                               template_name='',
                               floor_parameter_xlsx='',
                               progress_file=''):
    if progress_file == '':
        progress_file = './result/tmp'
    cad_counter = Counter()
    output_dwg_list = []
    output_dwg = ''

    def read_beam_multi(beam_filename, temp_file):
        msp_beam, doc_beam = read_beam_cad(
            beam_filename=beam_filename, progress_file=progress_file)
        sort_beam_cad(msp_beam=msp_beam,
                      layer_config=layer_config,
                      entity_config=get_template(template_name),
                      temp_file=temp_file,
                      progress_file=progress_file)
        output_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(temp_file),
                                                    progress_file=progress_file,
                                                    rebar_parameter_excel=floor_parameter_xlsx)
        output_dwg = draw_rebar_line(class_beam_list=output_beam_list,
                                     msp_beam=msp_beam,
                                     doc_beam=doc_beam,
                                     output_folder=output_folder,
                                     project_name=project_name)
        return output_beam_list, cad_data, os.path.basename(output_dwg)
    start = time.time()  # 開始測量執行時間
    with Pool(processes=10) as p:
        jobs = []
        beam_list = []
        for i, filename in enumerate(beam_filenames):
            temp_new = os.path.splitext(temp_file)[0]
            beam_temp = f'{temp_new}-{i}.pkl'
            jobs.append(p.apply_async(read_beam_multi, (filename, beam_temp)))
        for job in jobs:
            output_beam_list, cad_data, output_dwg = job.get()
            cad_counter.update(cad_data)
            output_dwg_list.append(output_dwg)
            beam_list.extend(output_beam_list)
    save_temp_file.save_pkl(beam_list, f'{temp_new}-beam_list.pkl')
    save_temp_file.save_pkl(cad_counter, f'{temp_new}-cad_data.pkl')
    output_file_list = create_report(class_beam_list=beam_list,
                                     floor_parameter_xlsx=floor_parameter_xlsx,
                                     output_folder=output_folder,
                                     project_name=project_name,
                                     cad_data=cad_counter,
                                     progress_file=progress_file)
    end = time.time()
    progress(f'執行時間：{end - start}s', progress_file)
    print("執行時間：%f 秒" % (end - start))
    # return os.path.basename(excel_filename),os.path.basename(excel_filename_rcad),output_dwg_list
    output_file_list = [os.path.basename(file) for file in output_file_list]
    return output_file_list, output_dwg_list


def get_template(name: str):

    if name == '公司2':
        return {
            'rebar_layer': ['AcDbPolyline'],
            'rebar_data_layer': ['AcDbMText'],
            'rebar_data_leader_layer': ['AcDbLeader'],
            'tie_text_layer': ['AcDbText']
        }
    if name == '公司1':
        return {
            'rebar_layer': ['AcDbLine'],
            'rebar_data_layer': ['AcDbText', 'AcDbMText'],
            'rebar_data_leader_layer': ['AcDbPolyline'],
            'tie_text_layer': ['AcDbMText']
        }
    if name == '公司3':
        return {
            'rebar_layer': ['AcDbPolyline'],
            'rebar_data_layer': ['AcDbText'],
            'rebar_data_leader_layer': ['AcDbPolyline'],
            'tie_text_layer': ['AcDbText']
        }


def floor_parameter(beam_list: list[Beam], floor_parameter_xlsx: str):
    parameter_df: pd.DataFrame
    floor_list: list[Floor]
    floor_list = []
    parameter_df = read_parameter_df(floor_parameter_xlsx, '梁參數表')
    parameter_df.set_index(['樓層'], inplace=True)
    for floor_name in parameter_df.index:
        temp_floor = Floor(str(floor_name))
        floor_list.append(temp_floor)
        temp_floor.set_beam_prop(parameter_df.loc[floor_name])
        temp_floor.add_beam(
            [b for b in beam_list if b.floor == temp_floor.floor_name])
    return floor_list
# combine dim with text arrow


def compare_line_with_dim(coor_to_dim_list: list[tuple[tuple[float, float], float, tuple[tuple[float, float], tuple[float, float]]]],
                          class_beam_list: list[Beam],
                          coor_to_block_list: list[tuple, dict[str, str], dict[str, str]],
                          progress_file: str):
    # coor_to_arrow_dic:{head:tail,float,mid}
    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])*3
    min_diff = 100
    if True:
        for block in coor_to_block_list:
            dim_list = [
                dim for dim in coor_to_dim_list if inblock(block[0], dim[0])]
            beam_list = [beam for beam in class_beam_list if inblock(
                block[0], beam.get_coor())]
            for beam in beam_list:
                for rebar in beam.rebar_list:
                    try:
                        rebar_dim = [dim for dim in dim_list if ((rebar.start_pt.x - dim[0][0]) * (rebar.end_pt.x - dim[0][0]) <= 0
                                                                 and abs(dim[0][1] - rebar.arrow_coor[0][1]) < 150)]
                        # 下層筋
                        if rebar.arrow_coor[0][1] <= rebar.arrow_coor[1][1]:
                            rebar_dim = [dim for dim in rebar_dim if (
                                dim[0][1] > rebar.arrow_coor[0][1])]
                        else:
                            rebar_dim = [dim for dim in rebar_dim if (
                                dim[0][1] < rebar.arrow_coor[0][1])]

                        if not rebar_dim:
                            continue

                        min_dim = min(rebar_dim, key=lambda dim: _get_distance(
                            dim[0], rebar.arrow_coor[0]))
                        if abs(rebar.length - min_dim[1]) > 1:
                            progress(
                                f'{rebar.start_pt} :{rebar.length} <> {min_dim[0]}:{min_dim[1]}', progress_file)
                            if abs(rebar.length - min_dim[1]) < 1:
                                continue
                            if (min_dim[2][0][0] - rebar.arrow_coor[0][0]) * (min_dim[2][1][0] - rebar.arrow_coor[0][0]) > 0:
                                rebar.length -= min_dim[1]
                                if abs(rebar.end_pt.x - min_dim[0][0]) < abs(rebar.start_pt.x - min_dim[0][0]):
                                    rebar.end_pt.x = rebar.start_pt.x + rebar.length
                                else:
                                    rebar.start_pt.x = rebar.end_pt.x - rebar.length
                            else:
                                rebar.length = min_dim[1]
                                if abs(rebar.end_pt.x - min_dim[0][0]) > abs(rebar.start_pt.x - min_dim[0][0]):
                                    rebar.end_pt.x = rebar.start_pt.x + rebar.length
                                else:
                                    rebar.start_pt.x = rebar.end_pt.x - rebar.length
                    except:
                        progress(
                            f'{beam.floor}:{beam.serial} compare_line_with_dim error', progress_file)
                        # new_rebar_arrow,_ = sort_arrow_line(coor_to_rebar_list=[((origin_start_pt_x,rebar.start_pt.y),
                        #                                     (origin_end_pt_x,rebar.end_pt.y),
                        #                                     abs(origin_length - rebar.length))],
                        #                                   coor_to_arrow_dic=coor_to_arrow_dic)
                        # new_coor_to_arrow_dic, _ , _ = sort_arrow_to_word(coor_to_arrow_dic=new_rebar_arrow,
                        #                                                   coor_to_data_list= coor_to_data_list)
                        # arrow_head,arrow_item = zip(*new_coor_to_arrow_dic.items())
                        # tail_coor,length,line_mid_coor,number,size,text_coor= arrow_item[0]
                        # beam.add_rebar(start_pt = line_mid_coor, end_pt = line_mid_coor,
                        #                length = length, number = number,
                        #                size = size,text=f'{number}-{size}',arrow_coor= (arrow_head[0],tail_coor))


if __name__ == '__main__':
    from main import GetAllFiles
    # from multiprocessing import Process, Pool
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # beam_filename = r"D:\Desktop\BeamQC\TEST\INPUT\2022-11-18-17-16temp-XS-BEAM.dwg"#sys.argv[1] # XS-BEAM的路徑
    beam_filename = r"D:\Desktop\BeamQC\TEST\2023-0814\Drawing2.dwg"
    beam_filenames = [r"D:\Desktop\BeamQC\TEST\2023-0617\FB.dwg",
                      r"D:\Desktop\BeamQC\TEST\2023-0617\GB.dwg",
                      r"D:\Desktop\BeamQC\TEST\2023-0617\GB2.dwg",
                      r"D:\Desktop\BeamQC\TEST\2023-0617\SB.dwg"]
    # beam_filenames = [r"D:\Desktop\BeamQC\TEST\2023-0320\東仁\XS-BEAM(北基地)-B-1.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0320\東仁\XS-BEAM(北基地)-B-2.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-0320\東仁\XS-BEAM(北基地)-B-3.dwg",
    #                   ]
    # beam_filenames = GetAllFiles(r"D:\Desktop\BeamQC\TEST\2023-0324\東仁")
    #                 r"D:\Desktop\BeamQC\TEST\INPUT\1-2023-02-24-17-37-sb-test.dwg",
    #                 r"D:\Desktop\BeamQC\TEST\INPUT\1-2023-02-24-17-37-fb-test.dwg"]
    # beam_filename = r"D:\Desktop\BeamQC\TEST\DEMO\數量計算\Other-大梁\2F-大梁 - Rec.dwg"
    progress_file = './result/tmp'  # sys.argv[14]
    rebar_file = './result/0107-rebar_wu2.txt'  # rebar.txt的路徑 -> 計算鋼筋和箍筋總量
    tie_file = './result/0107-tie_wu2.txt'  # rebar.txt的路徑 -> 把箍筋跟梁綁在一起
    # output_folder ='D:/Desktop/BeamQC/TEST/OUTPUT/'
    output_folder = r'D:\Desktop\BeamQC\TEST\2023-0831'
    # floor_parameter_xlsx = r'D:\Desktop\BeamQC\file\樓層參數_floor.xlsx'
    floor_parameter_xlsx = r'D:\Desktop\BeamQC\TEST\2023-0831\P2022-09A 中德建設楠梓區15FB4-2023-08-31-11-15-floor.xlsx'
    project_name = '0831-楠梓'
    # 在beam裡面自訂圖層
    # layer_config = {
    #     'rebar_data_layer': ['S-LEADER'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['S-REINF'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['S-TEXT'],  # 箍筋文字圖層
    #     'block_layer': ['0'],  # 框框的圖層
    #     'beam_text_layer': ['S-RC'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['S-RC'],  # 支承端圖層
    #     's_dim_layer': ['S-DIM']  # 標註線圖層
    # }
    layer_config = {
        'rebar_data_layer': ['P1'],  # 箭頭和鋼筋文字的塗層
        'rebar_layer': ['P2'],  # 鋼筋和箍筋的線的塗層
        'tie_text_layer': ['P1'],  # 箍筋文字圖層
        'block_layer': ['0'],  # 框框的圖層
        'beam_text_layer': ['P2'],  # 梁的字串圖層
        'bounding_block_layer': ['S-ARCH'],
        'rc_block_layer': ['P1'],  # 支承端圖層
        's_dim_layer': ['P1']  # 標註線圖層
    }

    # layer_config = {
    #     'rebar_data_layer':['NBAR'], # 箭頭和鋼筋文字的塗層
    #     'rebar_layer':['RBAR'], # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer':['NBAR'], # 箍筋文字圖層
    #     'block_layer':['DEFPOINTS'], # 框框的圖層
    #     'beam_text_layer' :['TITLE'], # 梁的字串圖層
    #     'bounding_block_layer':['S-ARCH'],
    #     'rc_block_layer':['OLINE']
    # }

    # layer_config = {
    #     'rebar_data_layer':['E6DIM'], # 箭頭和鋼筋文字的塗層
    #     'rebar_layer':['MAINBAR','WEB','STIRRUP'], # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer':['E6DIM'], # 箍筋文字圖層
    #     'block_layer':['0'], # 框框的圖層
    #     'beam_text_layer' :['BeamName'], # 梁的字串圖層
    #     'bounding_block_layer':['S-ARCH'],
    #     'rc_block_layer':['MEMBER'],
    #     's_dim_layer':['S-DIM'] # 標註線圖層
    # }

    # layer_config = {
    #     'rebar_data_layer': ['7', '5'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['2', '5'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['7'],  # 箍筋文字圖層
    #     'block_layer': ['Frame'],  # 框框的圖層
    #     'beam_text_layer': ['7'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['7'],
    #     's_dim_layer': ['5']  # 標註線圖層
    # }

    # layer_config ={
    #     'rebar_data_layer': ['主筋文字引線','主筋文字','扭力筋文字','扭力筋文字引線'],
    #     'rebar_layer': ['主筋線','剪力筋線','扭力筋線'],
    #     'tie_text_layer': ['剪力筋文字'],
    #     'block_layer': ['SHEET'],
    #     'beam_text_layer': ['梁跨名稱'],
    #     'bounding_block_layer': [''],
    #     'rc_block_layer': ['梁柱截斷記號','邊界線-梁支撐外線']}

    # entity_type ={
    #     'rebar_layer':['AcDbPolyline'],
    #     'rebar_data_layer':['AcDbText'],
    #     'rebar_data_leader_layer':['AcDbPolyline'],
    #     'tie_text_layer':['AcDbText']
    # }

    # Elements
    entity_type = {
        'rebar_layer': ['AcDbPolyline'],
        'rebar_data_layer': ['AcDbMText'],
        'rebar_data_leader_layer': ['AcDbLeader'],
        'tie_text_layer': ['AcDbText']
    }

    # entity_type ={
    #     'rebar_layer':['AcDbLine'],
    #     'rebar_data_layer':['AcDbText'],
    #     'rebar_data_leader_layer':['AcDbLine'],
    #     'tie_text_layer':['AcDbText']
    # }

    # entity_type ={
    #     'rebar_layer':['AcDbLine'],
    #     'rebar_data_layer':['AcDbText','AcDbMText'],
    #     'rebar_data_leader_layer':['AcDbPolyline'],
    #     'tie_text_layer':['AcDbMText']
    # }

    entity_type = {
        'rebar_layer': ['AcDbLine'],
        'rebar_data_layer': ['AcDbText'],
        'rebar_data_leader_layer': ['AcDbLine'],
        'tie_text_layer': ['AcDbText']
    }

    start = time.time()
    # msp_beam, doc_beam = read_beam_cad(
    #     beam_filename=beam_filename, progress_file=progress_file)
    # sort_beam_cad(msp_beam=msp_beam,
    #               layer_config=layer_config,
    #               entity_config=entity_type,
    #               progress_file=progress_file,
    #               temp_file=r'D:\Desktop\BeamQC\TEST\2023-0814\temp.pkl')
    # count_beam_multiprocessing(beam_filenames=beam_filenames,
    #                            layer_config=layer_config,
    #                            temp_file='0617_Wuku.pkl',
    #                            project_name=project_name,
    #                            output_folder=output_folder,
    #                            template_name='公司2',
    #                            floor_parameter_xlsx=floor_parameter_xlsx)
    # class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(r'D:\Desktop\BeamQC\TEST\2023-0714\台南隆興段-2023-07-14-10-25-temp-0.pkl'),
    #                                            progress_file=progress_file,
    #                                            rebar_parameter_excel=floor_parameter_xlsx)
    # save_temp_file.save_pkl(class_beam_list,tmp_file=r'D:\Desktop\BeamQC\TEST\2023-0607\0607-GB-list.pkl')
    # save_temp_file.save_pkl(cad_data,tmp_file=r'D:\Desktop\BeamQC\TEST\2023-0607\0607-GB-cad.pkl')
    class_beam_list = save_temp_file.read_temp(
        r'D:\Desktop\BeamQC\TEST\2023-0831\P2022-09A 中德建設楠梓區15FB4-2023-08-31-11-15-temp-beam_list.pkl')
    # class_beam_list.extend(save_temp_file.read_temp(r'D:\Desktop\BeamQC\0617_Wuku-cad_data.pkl'))
    cad_data = save_temp_file.read_temp(
        r'D:\Desktop\BeamQC\TEST\2023-0831\P2022-09A 中德建設楠梓區15FB4-2023-08-31-11-15-temp-cad_data.pkl')
    create_report(class_beam_list=class_beam_list,
                  output_folder=output_folder,
                  project_name=project_name,
                  floor_parameter_xlsx=floor_parameter_xlsx,
                  cad_data=cad_data,
                  progress_file=progress_file)
    # draw_rebar_line(class_beam_list=class_beam_list,
    #                 msp_beam=msp_beam,
    #                 doc_beam=doc_beam,
    #                 output_folder=output_folder,
    #                 project_name=project_name)
    print(f'Total Time:{time.time() - start}')
