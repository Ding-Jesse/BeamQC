from __future__ import annotations
import time
import win32com.client
import re
import os
import pandas as pd
import numpy as np
import copy
import pythoncom
import json
import src.save_temp_file as save_temp_file
from typing import Literal
from math import sqrt, ceil
from item.beam import Beam, BeamType, RebarType
from item.floor import Floor, read_parameter_df, summary_floor_rebar, summary_floor_rebar_ratio
from item.excel import AddExcelDataBar, AddBorderLine
from item.excepteions import NoRebarDataError, BeamFloorNameError
from src.beam_scan import create_beam_scan, beam_check, create_sbeam_scan, create_fbeam_scan, output_detail_scan_report, output_ng_ratio
from src.plan_count import sort_plan_count
from src.main import OutputExcel, Add_Row_Title
from multiprocessing.pool import ThreadPool as Pool
from collections import Counter
from item.rebar import isRebarSize, readRebarExcel
from item.pdf import create_scan_pdf
from src.logger import setup_custom_logger
from src.beam_draw import draw_beam_rebar_dxf
error_file = './result/error_log.txt'  # error_log.txt的路徑
one_side = False
global main_logger


def vtFloat(l):  # 要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)


def vtPnt(x, y, z=0):
    """座標點轉化爲浮點數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def error(error_message):  # 把錯誤訊息印到error.log裡面
    try:
        main_logger.error(error_message)
    except:
        print(error_message)


def progress(message):  # 把進度印到progress裡面，在app.py會對這個檔案做事
    try:
        main_logger.info(message)
    except:
        print(message)


def read_beam_cad(beam_filename):
    error_count = 0
    pythoncom.CoInitialize()
    progress('開始讀取梁配筋圖')
    # Step 1. 打開應用程式
    wincad_beam = None
    while wincad_beam is None and error_count <= 10:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")

        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 1: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 1/3')

    # Step 2. 匯入檔案
    doc_beam = None
    while wincad_beam and doc_beam is None and error_count <= 10:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)

        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 2: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 2/3')

    # Step 3. 匯入modelspace
    msp_beam = None
    while doc_beam and msp_beam is None and error_count <= 10:
        try:
            msp_beam = doc_beam.Modelspace

        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 3: {e}, error_count = {error_count}.')

    progress('梁配筋圖讀取進度 3/3')

    if msp_beam is None:
        try:
            if doc_beam:
                doc_beam.Close(SaveChanges=False)
        except:
            error(f'fail while closing file:{os.path.basename(beam_filename)}')
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


def sort_beam_cad(msp_beam,
                  doc_beam,
                  layer_config: dict,
                  entity_config: dict,
                  temp_file='',
                  redraw=False,
                  break_rebar_point={}):
    '''
    If AcDbLine and AcDbText in "dim_layer" entity type
    If AcDbLine in "rebar_data_leader_layer" entity type
    '''
    # define custom layer
    rebar_layer = layer_config['rebar_layer']
    rebar_data_layer = layer_config['rebar_data_layer']
    tie_text_layer = layer_config['tie_text_layer']
    block_layer = layer_config['block_layer']
    beam_text_layer = layer_config['beam_text_layer']
    bounding_block_layer = layer_config['bounding_block_layer']
    beam_layer = layer_config['rc_block_layer']
    s_dim_layer = layer_config['s_dim_layer']
    burst_layer_list = layer_config['burst_layer_list']

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
    coor_to_dim_text = []
    coor_to_dim_line = []
    coor_to_break_point = []
    coor_to_arrow_line = []

    count = 0
    total = msp_beam.Count
    progress(
        f'梁配筋圖上共有{total}個物件，大約運行{int(total / 5500)}分鐘，請耐心等候')
    for msp_object in msp_beam:
        object_list = []
        error_count = 0
        while error_count <= 3 and not object_list:
            try:
                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference" and msp_object.Layer in burst_layer_list:
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read error in step 7: {ex}, error_count = {error_count}.')
        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                count += 1
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件')
                # 抓鋼筋的字的座標
                if object.Layer in rebar_data_layer and object.ObjectName in entity_config['rebar_data_layer']:
                    coor = (round(object.InsertionPoint[0], 2), round(
                        object.InsertionPoint[1], 2))
                    coor_to_data_list.append((object.TextString, coor))
                # 抓箭頭座標
                elif object.Layer in rebar_data_layer and \
                        object.ObjectName in entity_config['rebar_data_leader_layer']:
                    # object.Coordinates 有九個參數 -> 箭頭尖點座標，直角的座標，文字接出去的座標，都有x, y, z
                    if object.ObjectName == 'AcDbLine':
                        coor1 = (round(object.startPoint[0], 2), round(
                            object.startPoint[1], 2))
                        coor2 = (round(object.endPoint[0], 2), round(
                            object.endPoint[1], 2))
                        coor_to_arrow_line.append((coor1, coor2))
                    elif hasattr(object, 'Coordinates') and len(object.Coordinates) >= 8:
                        coor_to_arrow_dic[(round(object.Coordinates[0], 2), round(object.Coordinates[1], 2))] = (
                            round(object.Coordinates[6], 2), round(object.Coordinates[7], 2))
                    elif hasattr(object, 'Coordinates'):
                        coor_to_arrow_dic[(round(object.Coordinates[0], 2), round(object.Coordinates[1], 2))] = (
                            round(object.Coordinates[-2], 2), round(object.Coordinates[-1], 2))
                    elif hasattr(object, 'startPoint'):
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
                    if hasattr(object, 'TextOverride'):
                        coor = (round(object.TextPosition[0], 2), round(
                            object.TextPosition[1], 2))
                        coor_to_tie_text_list.append(
                            (object.TextOverride, coor))
                    else:
                        coor = (round(object.InsertionPoint[0], 2), round(
                            object.InsertionPoint[1], 2))
                        coor_to_tie_text_list.append((object.TextString, coor))
                    continue
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
                    if hasattr(object, 'TextOverride'):
                        coor = (round(object.TextPosition[0], 2), round(
                            object.TextPosition[1], 2))
                        coor_to_tie_text_list.append(
                            (object.TextOverride, coor))
                    else:
                        coor = (round(object.InsertionPoint[0], 2), round(
                            object.InsertionPoint[1], 2))
                        coor_to_tie_text_list.append((object.TextString, coor))

                if object.Layer in beam_layer and object.ObjectName in entity_config['rc_block_layer']:
                    if hasattr(object, 'Coordinates') and len(object.Coordinates) >= 8:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        coor_to_rc_block_list.append(((coor1, coor2), ''))
                    elif object.ObjectName in ['AcDbLine']:
                        coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        if coor1[0] == coor2[0]:
                            coor_to_rc_block_list.append(((coor1, coor2), ''))

                if object.Layer in s_dim_layer and object.ObjectName in entity_config['dim_layer'] and\
                        hasattr(object, 'TextPosition') and hasattr(object, 'Measurement'):
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
                # 如標註線非標註形式
                elif object.Layer in s_dim_layer and \
                        object.ObjectName in ['AcDbText'] and \
                        object.ObjectName in entity_config['dim_layer']:
                    midpoint = (round((object.GetBoundingBox()[0][0] + object.GetBoundingBox()[1][0]) / 2, 2),
                                round((object.GetBoundingBox()[0][1] + object.GetBoundingBox()[1][1]) / 2, 2))
                    data = [object.TextString, midpoint]

                    coor_to_dim_text.append(data)
                # 如標註線非標註形式
                elif object.Layer in s_dim_layer and \
                    object.ObjectName in ['AcDbLine'] and \
                        object.ObjectName in entity_config['dim_layer']:
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                        object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                        object.GetBoundingBox()[1][1], 2))
                    if coor1[1] == coor2[1]:  # only horztion line can be dim line
                        coor_to_dim_line.append([(coor1, coor2), '', []])

                # 取斷筋點
                if break_rebar_point:
                    if object.Layer in break_rebar_point['layer'] \
                            and object.ObjectName in break_rebar_point['entity']:
                        if hasattr(object, 'StartPoint') and hasattr(object, 'EndPoint'):
                            coor1 = object.StartPoint
                            coor2 = object.EndPoint
                            if coor1[1] != coor2[1] and coor1[0] != coor2[0]:
                                coor_to_break_point.append([coor1, coor2])
                        elif object.GetBoundingBox():
                            coor1 = (round(object.GetBoundingBox()[0][0], 2), round(
                                object.GetBoundingBox()[0][1], 2))
                            coor2 = (round(object.GetBoundingBox()[1][0], 2), round(
                                object.GetBoundingBox()[1][1], 2))
                            if coor1[1] != coor2[1] and coor1[0] != coor2[0]:
                                coor_to_break_point.append([coor1, coor2])
                continue
            except Exception as ex:
                # print(f'error:{error_count}')
                object_list.append(object)
                error_count += 1
                time.sleep(5)
                error(f'error while {ex}')
    progress('梁配筋圖讀取進度 7/15')
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
                             'coor_to_dim_list': coor_to_dim_list,
                             'coor_to_dim_line': coor_to_dim_line,
                             'coor_to_dim_text': coor_to_dim_text,
                             'coor_to_break_point': coor_to_break_point,
                             'coor_to_arrow_line': coor_to_arrow_line
                             }, temp_file)
    try:
        if not redraw:
            doc_beam.Close(SaveChanges=False)
    except:
        error('Cant Close Dwg File')

# 整理箭頭與直線對應


def break_down_line(coor_to_rebar_list: list,
                    coor_to_break_point: list[list]):

    remove_item = []
    for i, rebar in enumerate(coor_to_rebar_list[:]):
        head_coor = rebar[0]
        tail_coor = rebar[1]
        mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
        length = rebar[2]

        # find the break point inside the line
        break_points = [break_point for break_point in coor_to_break_point
                        if (round(break_point[0][1], 2) == mid_coor[1] or round(break_point[1][1], 2) == mid_coor[1]) and
                        (break_point[0][0] - head_coor[0]) * (break_point[0][0] - tail_coor[0]) < 0]
        if not break_points:
            continue
        break_points.sort(key=lambda c: c[0][0])
        # remove the origin one
        remove_item.append(i)

        # split the line
        for break_point in break_points:
            contact_point = None
            coor1, coor2 = break_point
            if round(coor1[1], 2) == head_coor[1]:
                contact_point = coor1
            if round(coor2[1], 2) == head_coor[1]:
                contact_point = coor2

            split_x = round(contact_point[0], 2)

            assert split_x > head_coor[0], 'Error'

            coor_to_rebar_list.append(
                (head_coor, (split_x, head_coor[1]), abs(split_x - head_coor[0])))

            head_coor = (split_x, head_coor[1])

        if head_coor[0] < tail_coor[0]:
            split_x = round(tail_coor[0], 2)
            coor_to_rebar_list.append(
                (head_coor, (split_x, head_coor[1]), abs(split_x - head_coor[0])))
    remove_item.reverse()
    for index in remove_item:
        coor_to_rebar_list.pop(index)
    return


def sort_arrow_line(coor_to_arrow_dic: dict,
                    coor_to_rebar_list: list,
                    coor_to_dim_list: list[tuple[tuple[float, float], float, tuple[tuple[float, float], tuple[float, float]]]],
                    coor_to_bend_rebar_list: list[tuple[tuple[float, float], tuple[float, float], float]],
                    **kwargs):
    start = time.time()
    # method 2
    new_coor_to_arrow_dic = {}
    no_arrow_line_list = []

    min_diff = 1
    dim_spacing = kwargs.get('dim_spacing', 200)

    coor_to_rebar_list.sort()
    # coor_to_rebar_list = [coor_to_rebar_list[2055]]
    for i, rebar in enumerate(coor_to_rebar_list):
        if i % 100 == 0:
            progress(f"整理直線鋼筋與標示 進度:{i}/{len(coor_to_rebar_list)}")
        head_coor = rebar[0]
        tail_coor = rebar[1]
        mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
        length = rebar[2]
        arrow_dict = {k: v for k, v in coor_to_arrow_dic.items() if (
            head_coor[0] - k[0]) * (tail_coor[0] - k[0]) <= 0}

        if not arrow_dict:
            continue
        value_pair = min(arrow_dict.items(),
                         key=lambda x: abs(mid_coor[1] - x[0][1]))
        if (abs(value_pair[0][1] - mid_coor[1]) > min_diff):
            no_arrow_line_list.append(rebar)
            continue
        for key, value in {k: v for k, v in arrow_dict.items() if
                           abs(k[1] - value_pair[0][1]) < min_diff}.items():
            with_dim = False
            mid_coor = (
                round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])

            rebar_coor1 = key
            rebar_coor2 = value
            # 下層筋
            if key[1] > value[1]:
                dim_match_line = True
                dim_list = [dim for dim in coor_to_dim_list if (
                    dim[2][0][0] == head_coor[0]) and (dim[2][1][0] == tail_coor[0])]
                if not dim_list:
                    # 處理只有單邊有標註線之梁配筋
                    if one_side:
                        dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < key[1]) and (
                            dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
                    else:
                        dim_match_line = False
                        dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < key[1]) and (
                            dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]

            # 上層筋
            else:
                dim_match_line = False
                dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] > key[1]) and (
                    dim[2][0][0] - key[0]) * (dim[2][1][0] - key[0]) <= 0]
            if dim_list:
                dim = min(dim_list, key=lambda dim: abs(
                    dim[0][1] - key[1]))
                text_postion, text_value, \
                    (line_1_point, line_2_point) = dim
                if abs(text_postion[1] - key[1]) < dim_spacing or \
                        (dim_match_line and abs(text_postion[1] - key[1]) < dim_spacing * 2.5):
                    rebar_coor1 = ((line_1_point[0] + line_2_point[0])/2,
                                   head_coor[1])
                    rebar_coor2 = (value[0], value[1])
                    length = text_value
                    with_dim = True
                    mid_coor = (
                        (line_1_point[0] + line_2_point[0])/2, head_coor[1])
            new_coor_to_arrow_dic.update(
                {rebar_coor1: (rebar_coor2, length, mid_coor, with_dim,
                               (mid_coor, rebar[2])
                               )})
    for i, bend_rebar in enumerate(coor_to_bend_rebar_list):
        if i % 100 == 0:
            progress(f"整理彎鉤鋼筋與標示 進度:{i}/{len(coor_to_rebar_list)}")
        arrow_dict = {}
        bend_coor = bend_rebar[0]
        line_coor = bend_rebar[1]
        mid_coor = (round((bend_coor[0] + line_coor[0]) / 2, 2), line_coor[1])
        length = bend_rebar[2]
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
            length = bend_rebar[2]
            for key, value in {k: v for k, v in arrow_dict.items() if abs(k[1] - value_pair[0][1]) < min_diff}.items():
                with_dim = False
                mid_coor = (
                    round((bend_coor[0] + line_coor[0]) / 2, 2), line_coor[1])
                length = bend_rebar[2]
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
                    {rebar_coor1: (rebar_coor2, length, mid_coor, with_dim, (mid_coor, bend_rebar[2]))})
    progress(f'sort arrow to line:{time.time() - start}')
    return new_coor_to_arrow_dic, no_arrow_line_list


# 整理箭頭與鋼筋文字對應
def sort_arrow_to_word(coor_to_arrow_dic: dict,
                       coor_to_data_list: list,
                       middle_tie_pattern: dict,
                       rebar_pattern: dict,
                       **kwargs):

    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        # 避免抓到其他層的主筋資料
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1]) * 3

    start = time.time()
    min_diff = kwargs.get('arrow_to_word_min_diff', 100)

    new_coor_to_arrow_dic = {}
    head_to_data_dic = {}  # 座標 -> (number, size)
    tail_to_data_dic = {}

    # tie_text_coor_list = [(text,coor) for text,coor in coor_to_data_list if re.match(tie_pattern['pattern'])]
    rebar_text_coor_list = [(text, coor) for text, coor in coor_to_data_list if
                            re.search(rebar_pattern['pattern'], text) or
                            re.search(middle_tie_pattern['pattern'], text)]

    # rebar_text_coor_list = [(text, coor) for text, coor in coor_to_data_list if
    #                         re.search(middle_tie_pattern['pattern'], text)]
    # method 2
    # 腰筋會抓到上層筋
    for i, data in enumerate(coor_to_arrow_dic.items()):
        if i % 100 == 0:
            progress(f"整理箭頭與鋼筋文字對應 進度:{i}/{len(coor_to_arrow_dic.items())}")
        arrow_head, arrow_data = data
        arrow_tail, rebar_length, line_coor, with_dim, origin_rebar = arrow_data

        rebar_data_temp = []
        if arrow_head[1] > arrow_tail[1]:
            # 箭頭朝上
            rebar_data_temp = [
                (text, coor) for text, coor in rebar_text_coor_list if coor[1] <= arrow_head[1]]
        if arrow_head[1] < arrow_tail[1]:
            # 箭頭朝下
            rebar_data_temp = [
                (text, coor) for text, coor in rebar_text_coor_list if coor[1] >= arrow_head[1]]

        if not rebar_data_temp:
            progress(
                f'arrow head:{arrow_head} has no rebar data')

            rebar_data_temp = rebar_text_coor_list

        if not rebar_data_temp:
            raise NoRebarDataError

        text, coor = min(rebar_data_temp,
                         key=lambda rebar_text: _get_distance(arrow_data[0], rebar_text[1]))

        if (abs(arrow_tail[1] - coor[1]) > min_diff):
            progress(
                f'{arrow_head} / {arrow_data} cant find pair arrow')
            continue

        # Decide the rebar size and number
        number = size = ''

        match_middle_tie = re.search(middle_tie_pattern['pattern'], text)
        match_rebar = re.search(rebar_pattern['pattern'], text)

        if match_middle_tie:
            # progress(f'{text} match middle tie rule {middle_tie_pattern}')
            number = match_middle_tie.group(middle_tie_pattern['num'])
            size = match_middle_tie.group(middle_tie_pattern['size']) + 'E.F.'
            text = f'{number}-{size}E.F.'
            # 腰筋不受標註線影響
            if with_dim:
                line_coor = origin_rebar[0]
                rebar_length = origin_rebar[1]
                with_dim = False
        elif match_rebar:
            number = match_rebar.group(rebar_pattern['num'])
            size = match_rebar.group(rebar_pattern['size'])

        rebar_data = [arrow_tail, rebar_length, line_coor, with_dim]

        # number = text.split('-')[0]
        # size = text.split('-')[1]
        if not isRebarSize(size):
            progress(f'{size} not satisfied rebar rule')
            continue
        if not number.isdigit():
            progress(f'{text} not satisfied rebar rule')
            continue

        rebar_data.extend([number, size, coor])
        new_coor_to_arrow_dic.update({arrow_head: (*rebar_data,)})
        head_to_data_dic.update({(line_coor[0] - rebar_length/2, line_coor[1]): {
                                'number': number, 'size': size, 'dim': with_dim}})
        tail_to_data_dic.update({(line_coor[0] + rebar_length/2, line_coor[1]): {
                                'number': number, 'size': size, 'dim': with_dim}})

    progress(f'sort arrow to word:{time.time() - start}')
    return new_coor_to_arrow_dic, head_to_data_dic, tail_to_data_dic


def sort_line_to_word(coor_to_rebar_list: list,
                      coor_to_data_list: list,
                      coor_to_dim_list: list,
                      rebar_pattern: str = r'(\d+)-(#\d)',
                      middle_tie_pattern: str = r'(\d+)-(#\d+)E.F',
                      tie_pattern: str = r'(#\d+)@(\d+)',
                      measure_type: str = 'cm'):
    '''
    For the case with no arrow
    coor_to_rebar_list = [(head , tail , length)]
    '''
    min_diff = 50
    new_coor_to_arrow_dic = {}
    no_arrow_line_list = []
    factor = 1
    if measure_type == 'mm':
        factor = 5

    coor_to_data_list = [(text, coor) for text, coor in coor_to_data_list if
                         (re.match(middle_tie_pattern, text) or
                          re.match(rebar_pattern, text)) and
                         not re.search(tie_pattern, text)]
    coor_to_rebar_list = [
        rebar for rebar in coor_to_rebar_list if rebar[0] == (7821.03, 7124.74)]
    for i, rebar in enumerate(coor_to_rebar_list):
        if i % 100 == 0:
            progress(f"整理直線鋼筋與標示 進度:{i}/{len(coor_to_rebar_list)}")
        head_coor = rebar[0]
        tail_coor = rebar[1]
        mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
        length = rebar[2]
        text_list = [(text, coor) for text, coor in coor_to_data_list if (
            head_coor[0] - coor[0]) * (tail_coor[0] - coor[0]) <= 0]

        if not text_list:
            continue

        closet_text = min(text_list,
                          key=lambda x: abs(mid_coor[1] - x[1][1]))

        closet_distance = abs(closet_text[1][1] - mid_coor[1])

        if closet_distance > min_diff * factor:
            no_arrow_line_list.append(rebar)
            continue

        for text, coor in [(text, coor) for text, coor in text_list if
                           abs(coor[1] - closet_text[1][1]) <= (min_diff * factor)]:
            with_dim = False
            rebar_coor1 = mid_coor
            rebar_coor2 = coor
            # 下層筋
            if mid_coor[1] > coor[1]:
                dim_match_line = True
                dim_list = [dim for dim in coor_to_dim_list if (
                    dim[2][0][0] == head_coor[0]) and (dim[2][1][0] == tail_coor[0])]
                if not dim_list:
                    # 處理只有單邊有標註線之梁配筋
                    if one_side:
                        dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < coor[1]) and (
                            dim[2][0][0] - coor[0]) * (dim[2][1][0] - coor[0]) <= 0]
                    else:
                        dim_match_line = False
                        dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] < coor[1]) and (
                            dim[2][0][0] - coor[0]) * (dim[2][1][0] - coor[0]) <= 0]
            # 上層筋
            else:
                dim_list = [dim for dim in coor_to_dim_list if (dim[0][1] > coor[1]) and (
                    dim[2][0][0] - coor[0]) * (dim[2][1][0] - coor[0]) <= 0]
            if dim_list:
                dim = min(dim_list, key=lambda dim: abs(
                    dim[0][1] - coor[1]))
                text_postion, text_value, (line_1_point,
                                           line_2_point) = dim
                if abs(text_postion[1] - coor[1]) < 200 * factor or \
                        (dim_match_line and abs(text_postion[1] - coor[1]) < 500 * factor):
                    rebar_coor1 = (
                        (line_1_point[0] + line_2_point[0])/2, head_coor[1])
                    rebar_coor2 = coor
                    length = text_value
                    with_dim = True
                    mid_coor = (
                        (line_1_point[0] + line_2_point[0])/2, head_coor[1])
            new_coor_to_arrow_dic.update(
                {rebar_coor1: (rebar_coor2, length, mid_coor, with_dim, (mid_coor, rebar[2]))})

    return new_coor_to_arrow_dic, no_arrow_line_list


def sort_noconcat_line(no_concat_line_list,
                       head_to_data_dic: dict,
                       tail_to_data_dic: dict):
    # start = time.time()
    coor_to_rebar_list_straight = []  # (頭座標，尾座標，長度，number，size)

    def _overlap(l1, l2):
        if l1[1] == l2[0][1]:
            return round(l2[0][0] - l1[0], 2)*round(l2[1][0] - l1[0], 2) <= 0
        return False

    def _cal_length(pt1, pt2):
        return sqrt((pt1[0]-pt2[0])**2 + (pt1[1]-pt2[1])**2)

    def _concat_line(line_list: list):
        for i, line in enumerate(line_list[:]):
            if i % 100 == 0:
                progress(f"搭接直線鋼筋 進度:{i}/{len(line_list)}")
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
                                                            tail_rebar['number'],
                                                            tail_rebar['size'], False))
                    if tail_rebar['dim']:
                        coor_to_rebar_list_straight.append((head_coor,
                                                            tail_coor,
                                                            _cal_length(
                                                                head_coor, tail_coor),
                                                            head_rebar['number'],
                                                            head_rebar['size'], False))
                    if not head_rebar['dim'] and not tail_rebar['dim']:
                        coor_to_rebar_list_straight.append((head_coor, tail_coor,
                                                            _cal_length(
                                                                head_coor, tail_coor),
                                                            tail_rebar['number'],
                                                            tail_rebar['size'], False))
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
            if coor_to_rebar_list_straight:
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
    coor_to_bend_rebar_list = []
    for i, bend in enumerate(no_concat_bend_list):
        if i % 100 == 0:
            progress(f"搭接彎鉤鋼筋 進度:{i}/{len(no_concat_bend_list)}")
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

            continue
        overlap_line = {k: v for k, v in tail_to_data_dic.items() if (
            k[0] >= horz_coor[0]) and (k[1] == horz_coor[1]) and (k[0] <= vert_coor[0])}
        if len(overlap_line.keys()) > 0:
            value_key, value_items = max(
                overlap_line.items(), key=lambda x: abs(x[0][0]-horz_coor[0]))
            coor_to_bend_rebar_list.append((vert_coor, value_key, line_length - abs(
                value_key[0] - horz_coor[0]), value_items['number'], value_items['size']))

            continue

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
    for i, bend in enumerate(rebar_bend_list[:]):
        if i % 100 == 0:
            progress(f"整理彎折與直線鋼筋 進度:{i}/{len(rebar_bend_list)}")
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


def sort_line_to_arrow(coor_to_arrow_line: list,
                       coor_to_data_list: list,
                       rebar_pattern: dict = None,
                       middle_tie_pattern: dict = None,
                       tie_pattern: dict = None,
                       **kwargs):
    '''
    For the dwg that the leader is line object
    coor_to_arrow_line = [(head , tail)]
    one arrow to one text
    There will be a problem if there are multi vertical line in a leader
    '''
    rebar_pattern_str: str = rebar_pattern['pattern']
    middle_tie_pattern_str: str = middle_tie_pattern['pattern']
    tie_pattern_str: str = tie_pattern['pattern']

    leader_spacing = kwargs.get('leader_spacing', 6)
    min_diff = kwargs.get('sort_line_to_arrow_mid_diff', 5)

    coor_to_arrow_dict = {}
    coor_to_data_list = [(text, coor) for text, coor in coor_to_data_list if
                         (re.match(middle_tie_pattern_str, text) or
                         re.match(rebar_pattern_str, text)) and
                         not re.search(tie_pattern_str, text)]

    horz_coor_to_arrow_line = [(head, tail) for (head, tail) in coor_to_arrow_line if
                               abs(head[1] - tail[1]) < 1]
    vert_coor_to_arrow_line = [(head, tail) for (head, tail) in coor_to_arrow_line if
                               abs(head[0] - tail[0]) < 1]
    for text, coor in coor_to_data_list:
        # First find horz line
        nearby_horz_lines = [(head, tail) for head, tail in horz_coor_to_arrow_line if
                             (abs(head[0] - coor[0]) < min_diff and abs(head[1] - coor[1]) < min_diff) or
                             (abs(tail[0] - coor[0]) < min_diff and abs(tail[1] - coor[1]) < min_diff)]
        if not nearby_horz_lines:
            continue

        for line in nearby_horz_lines:
            # define where is text pos , other is arrow line
            head, tail = line
            text_pos = arrow_pos = line_pos = None
            nearest_arrow_line = None
            if abs(head[0] - coor[0]) + abs(head[1] - coor[1]) < \
                    abs(tail[0] - coor[0]) + abs(tail[1] - coor[1]):
                text_pos = head
                arrow_pos = tail
            else:
                text_pos = tail
                arrow_pos = head
            # Then find the line to arrow pos
            possible_arrow_lines = [(head, tail) for (head, tail) in vert_coor_to_arrow_line if
                                    head == arrow_pos or tail == arrow_pos]

            if len(possible_arrow_lines) >= 1:
                nearest_arrow_line = possible_arrow_lines[0]
                # Find which is line pos
                line_pos = nearest_arrow_line[0] if nearest_arrow_line[1] == arrow_pos else nearest_arrow_line[1]

            if nearest_arrow_line is None:
                possible_across_arrow_lines = [(head, tail) for (head, tail) in vert_coor_to_arrow_line if
                                               ((head[1] - arrow_pos[1]) * (tail[1] - arrow_pos[1]) < 0) and
                                               head[0] == arrow_pos[0]]

                if len(possible_across_arrow_lines) >= 1:
                    nearest_arrow_line = possible_across_arrow_lines[0]

            if nearest_arrow_line is None:
                continue

            # For the line that across two horz , must modify the length
            vert_diff = 0
            vert_head, vert_tail = nearest_arrow_line

            possible_across_horz_lines = [(head, tail) for (head, tail) in horz_coor_to_arrow_line if
                                          ((head[1] - vert_head[1]) * (tail[1] - vert_tail[1]) < 0) and
                                          (head[0] == vert_head[0] or tail[0] == vert_tail[0])]

            if line_pos is not None:
                # if there is a line_pos , means it can define the point that may contact the line
                # Define its a upper or lower leader
                if min(vert_head[1], vert_tail[1]) < arrow_pos[1]:  # upper

                    possible_across_horz_lines = [
                        (head, tail) for (head, tail) in possible_across_horz_lines if
                        (arrow_pos[1] > head[1])
                    ]

                    if possible_across_horz_lines:
                        vert_diff = leader_spacing * \
                            len(possible_across_horz_lines)

                    coor_to_arrow_dict.update(
                        {(line_pos[0], round(line_pos[1] + vert_diff, 2)): text_pos})
                else:  # lower

                    possible_across_horz_lines = [
                        (head, tail) for (head, tail) in possible_across_horz_lines if
                        (arrow_pos[1] < head[1])
                    ]

                    if possible_across_horz_lines:
                        vert_diff = leader_spacing * \
                            len(possible_across_horz_lines)

                    coor_to_arrow_dict.update(
                        {(line_pos[0], round(line_pos[1] - vert_diff, 2)): text_pos})
            else:
                # for the line that cant decide the line pos , so add both leader
                if nearest_arrow_line[0][1] >= nearest_arrow_line[1][1]:
                    top_pt = nearest_arrow_line[0]
                    bot_pt = nearest_arrow_line[1]
                else:
                    top_pt = nearest_arrow_line[1]
                    bot_pt = nearest_arrow_line[0]

                # upper leader
                possible_across_horz_lines = [
                    (head, tail) for (head, tail) in possible_across_horz_lines if
                    (arrow_pos[1] > head[1])
                ]

                vert_diff = leader_spacing * len(possible_across_horz_lines)

                coor_to_arrow_dict.update(
                    {(top_pt[0], round(bot_pt[1] + vert_diff, 2)): text_pos})

                # lower leader
                possible_across_horz_lines = [
                    (head, tail) for (head, tail) in possible_across_horz_lines if
                    (arrow_pos[1] < head[1])
                ]

                vert_diff = leader_spacing * len(possible_across_horz_lines)

                coor_to_arrow_dict.update(
                    {(bot_pt[0], round(top_pt[1] - vert_diff, 2)): text_pos})

    return coor_to_arrow_dict


def count_tie(coor_to_tie_text_list: list,
              coor_to_tie_list: list,
              tie_pattern: dict = None,
              **kwargs):
    '''
    Sort Tie to format
    - '15-2#4@15',  # With a number before and after the hyphen
    - '2#3@20',     # With a number before the rebar
    - '13-#3@15',   # With a number prefix and hyphen
    - '#3@10'       # Without prefix 
    '''
    def extract_tie(tie: str):
        tie = tie.replace(' ', '')
        match = re.search(tie_pattern['pattern'], tie)
        if match is None:
            return None

        tie_string = tie_num = size = spacing = multi = None
        tie_num = match.group(tie_pattern['num'])
        size = match.group(tie_pattern['size'])
        spacing = match.group(tie_pattern['spacing'])
        multi = match.group(tie_pattern['multi'])

        if multi and tie_num:
            tie_string = f"{tie_num}-{multi}{size}@{spacing}"
        elif tie_num:
            tie_string = f"{tie_num}-{size}@{spacing}"
        else:
            tie_string = f"{size}@{spacing}"

        return {
            'tie_string': tie_string,
            'tie_num': tie_num,
            'size': size,
            'spacing': spacing,
            'multi': multi
        }

    if tie_pattern is None:
        error('No tie pattern in kwargs at count_tie')
        return []

    coor_sorted_tie_list = []
    for tie_text, coor in coor_to_tie_text_list:  # (字串，座標)

        tie = extract_tie(tie=tie_text)

        if tie is None:
            error(f'{tie_text} wrong tie format')
            continue

        tie_num = tie['tie_num']
        tie_string = tie['tie_string']
        multi = 1 if tie['multi'] == '' else int(tie['multi'])

        if tie_num == '':
            spacing = int(tie['spacing'])

            if spacing <= 0:
                error(f'{tie_text} wrong tie spacing')
                continue

            tie_left_list = [(bottom, top, length) for bottom, top, length in coor_to_tie_list if (
                bottom[0] < coor[0]) and (min(bottom[1], top[1]) < coor[1]) and (max(bottom[1], top[1]) > coor[1])]
            tie_right_list = [(bottom, top, length) for bottom, top, length in coor_to_tie_list if (
                bottom[0] > coor[0]) and (min(bottom[1], top[1]) < coor[1]) and (max(bottom[1], top[1]) > coor[1])]
            left_tie = min(tie_left_list, key=lambda t: abs(t[0][0] - coor[0]))
            right_tie = min(
                tie_right_list, key=lambda t: abs(t[0][0] - coor[0]))

            if not (tie_left_list and tie_right_list):
                error(f'{tie_text} {coor} no line bounded')
                continue

            count = int(
                abs(left_tie[0][0] - right_tie[0][0]) / spacing) * multi
        else:

            count = int(tie_num) * multi

        size = tie['size']

        coor_sorted_tie_list.append((tie_string, coor, tie_num, count, size))

    return coor_sorted_tie_list

# 組合手動框選與梁文字


def combine_beam_boundingbox(coor_to_block_list: list[tuple[tuple[tuple, tuple], dict, dict]],
                             coor_to_bounding_block_list: list,
                             class_beam_list: list[Beam],
                             coor_to_rc_block_list: list,
                             **kwargs):
    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])

    def _get_distance_block(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])*3

    same_block_rc_block_list = []
    # rank the border line (horztion > vertical)
    temp = []

    polyline_prior = kwargs.get('rc_block_polyline_prior', True)
    tol = kwargs.get('rc_block_tol', 100)

    for data in coor_to_rc_block_list:
        # determine priority horz >> vertical
        if abs(data[0][0][0] - data[0][1][0]) > tol // 2 and polyline_prior:
            rank = 0
        else:
            rank = 1
        temp.append((*data, rank))

    coor_to_rc_block_list = temp

    for i, beam in enumerate(class_beam_list):
        if i % 100 == 0:
            progress(f"整理梁與梁邊框 進度:{i}/{len(class_beam_list)}")
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
            else:
                same_block_rc_block_list = coor_to_rc_block_list

            left_bounding_box_list = [block for block in same_block_rc_block_list
                                      if block[0][1][0] <= beam.get_coor()[0] and block[0][1][1] > beam.get_coor()[1] - accept_y_diff]
            right_bounding_box_list = [block for block in same_block_rc_block_list
                                       if block[0][0][0] >= beam.get_coor()[0] and block[0][0][1] > beam.get_coor()[1] - accept_y_diff]
            if len(left_bounding_box_list) and len(right_bounding_box_list):
                left_bounding_box = min(
                    left_bounding_box_list, key=lambda b: (b[2], _get_distance_block(b[0][0], beam.get_coor())))
                right_bounding_box = min(
                    right_bounding_box_list, key=lambda b: (b[2], _get_distance_block(b[0][0], beam.get_coor())))
                top_bounding = max(left_bounding_box[0][0][1], left_bounding_box[0]
                                   [1][1], right_bounding_box[0][0][1], right_bounding_box[0][1][1])
                bot_bounding = min(left_bounding_box[0][0][1], left_bounding_box[0]
                                   [1][1], right_bounding_box[0][0][1], right_bounding_box[0][1][1])
                # 邊框距離編號過近
                count = 1
                while abs(beam.get_coor()[1] - max(top_bounding, bot_bounding)) < tol and count <= 3:
                    left_bounding_box_list = [
                        new_block for new_block in left_bounding_box_list if new_block[0][1][1] > top_bounding + 15]
                    right_bounding_box_list = [
                        new_block for new_block in right_bounding_box_list if new_block[0][1][1] > top_bounding + 15]
                    new_top_bound = []
                    if not len(left_bounding_box_list) and not len(right_bounding_box_list):
                        print(
                            f'{beam.floor}{beam.serial} no bounding, use outer block')
                        if not outer_block:
                            break
                        top_bounding = outer_block[0][0][1][1]
                        beam.set_bounding_box(left_bounding_box[0][1][0], beam.get_coor()[
                                              1], right_bounding_box[0][0][0], top_bounding)
                        break
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
                    count += 1
                beam.set_bounding_box(left_bounding_box[0][1][0], beam.get_coor()[
                                      1], right_bounding_box[0][0][0], top_bounding)
        else:
            nearest_block = min(
                bounding_box, key=lambda b: _get_distance(b[0][0], beam.get_coor()))
            beam.set_bounding_box(
                nearest_block[0][0][0], nearest_block[0][0][1], nearest_block[0][1][0], nearest_block[0][1][1])

# 組合箍筋與梁文字


def combine_beam_tie(coor_sorted_tie_list: list,
                     class_beam_list: list[Beam],
                     tie_pattern: dict = None,
                     **kwargs):
    # ((左下，右上),beam_name, list of tie, tie_count_dic, list of rebar,rebar_length_dic)
    def _get_distance(pt1, pt2):
        # return sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2)
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])

    for i, data in enumerate(coor_sorted_tie_list):
        if i % 100 == 0:
            progress(f"整理梁與箍筋 進度:{i}/{len(coor_sorted_tie_list)}")
        tie, coor, tie_num, count, size = data
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
        match_obj = re.search(tie_pattern['pattern'], tie)
        if match_obj:
            nearest_beam.add_tie(tie, coor, tie_num, count, size)
        else:
            progress(f'{tie} does not match tie type')

# 截斷主筋


def break_down_rebar(coor_to_arrow_dic: dict,
                     class_beam_list: list[Beam]):
    def _get_distance(pt1, pt2):
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])
    add_list = []

    for i, arrow_head in enumerate(list(coor_to_arrow_dic.keys())):
        if i % 100 == 0:
            progress(f"截斷直線鋼筋 進度:{i}/{len(coor_to_arrow_dic.keys())}")
        nearest_beam = None

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


def combine_beam_rebar(coor_to_arrow_dic: dict,
                       coor_to_rebar_list_straight: list,
                       coor_to_bend_rebar_list: list,
                       class_beam_list: list[Beam],
                       **kwargs):
    # 以箭頭的頭為搜尋中心
    def _get_distance(pt1, pt2):
        return abs(pt1[0]-pt2[0]) + abs(pt1[1]-pt2[1])

    for i, data in enumerate(coor_to_arrow_dic.items()):

        if i % 100 == 0:
            progress(f"結合梁與直線鋼筋 進度:{i}/{len(coor_to_arrow_dic.keys())}")
        arrow_head, arrow_item = data
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
        # if nearest_beam.serial == 'B2-3':
        #     print(f'B6-4:{arrow_head} {arrow_item}')
        nearest_beam.add_rebar(start_pt=line_mid_coor, end_pt=line_mid_coor,
                               length=length, number=number,
                               size=size, text=f'{number}-{size}',
                               arrow_coor=(arrow_head, text_coor),
                               with_dim=with_dim)

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


def count_rebar_in_block(coor_to_arrow_dic: dict,
                         coor_to_block_list: list,
                         coor_to_rebar_list_straight,
                         coor_to_bend_rebar_list):
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


def inblock(block: tuple, pt: tuple):
    pt_x = pt[0]
    pt_y = pt[1]
    if len(block) == 0:
        return False
    if (pt_x - block[0][0])*(pt_x - block[1][0]) < 0 and (pt_y - block[0][1])*(pt_y - block[1][1]) < 0:
        return True
    return False


def cal_beam_rebar(data: dict = None,
                   rebar_parameter_excel='',
                   **kwargs):
    # output_txt = f'{output_folder}{project_name}'
    progress('================start cal_beam_rebar================')
    if data is None:
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
    # (object.TextPosition, object.Measurement, (coor1, coor2))
    coor_to_dim_list = data['coor_to_dim_list']

    coor_to_dim_text = []

    if 'coor_to_dim_text' in data:
        coor_to_dim_text = data['coor_to_dim_text']

    coor_to_dim_line = []
    if 'coor_to_dim_line' in data:
        coor_to_dim_line = data['coor_to_dim_line']

    coor_to_arrow_line = []
    if 'coor_to_arrow_line' in data:
        coor_to_arrow_line = data['coor_to_arrow_line']

    coor_to_break_point = []
    if 'coor_to_break_point' in data:
        coor_to_break_point = data['coor_to_break_point']

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

    readRebarExcel(file_path=rebar_parameter_excel)

    # 2023-0505 add floor xlsx to found floor
    parameter_df: pd.DataFrame = read_parameter_df(
        rebar_parameter_excel, '梁參數表')
    floor_list = parameter_df['樓層'].tolist()

    def fix_floor_list(floor):
        if floor[-1] != 'F':
            floor += 'F'
        return floor
    floor_list = list(map(fix_floor_list, floor_list))

    # Extra Step: For dim is not dimension (If there coor_to_dim_line is not empty)
    if coor_to_dim_line:
        coor_to_dim_list = sort_dim_line_and_text(
            coor_to_dim_text, coor_to_dim_line)

    # Extra Step: leader is not leader (If there coor_to_arrow_line is not empty)
    if coor_to_arrow_line:
        coor_to_arrow_dic = sort_line_to_arrow(coor_to_arrow_line=coor_to_arrow_line,
                                               coor_to_data_list=coor_to_data_list,
                                               **kwargs)
    # Step 8. 對應箭頭跟鋼筋
    start = time.time()
    test_data = {}
    temp = {
        'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
        'coor_to_rebar_list': copy.deepcopy(coor_to_rebar_list),
        'coor_to_dim_list': copy.deepcopy(coor_to_dim_list),
        'coor_to_bend_rebar_list': copy.deepcopy(coor_to_bend_rebar_list),
    }

    break_down_line(coor_to_rebar_list=coor_to_rebar_list,
                    coor_to_break_point=coor_to_break_point)

    no_arrow_line_list = []
    if len(coor_to_arrow_dic) == 0:
        pass
        # coor_to_arrow_dic, no_arrow_line_list = sort_line_to_word(coor_to_rebar_list,
        #                                                           coor_to_data_list=coor_to_data_list,
        #                                                           coor_to_dim_list=coor_to_dim_list,
        #                                                           rebar_pattern=kwargs['rebar_pattern']['pattern'],
        #                                                           tie_pattern=kwargs['tie_pattern']['pattern'],
        #                                                           middle_tie_pattern=middle_tie_pattern['pattern'],
        #                                                           measure_type=measure_type
        #                                                           )
    else:
        coor_to_arrow_dic, \
            no_arrow_line_list = sort_arrow_line(coor_to_arrow_dic,
                                                 coor_to_rebar_list,
                                                 coor_to_dim_list=coor_to_dim_list,
                                                 coor_to_bend_rebar_list=coor_to_bend_rebar_list,
                                                 **kwargs)
    test_data.update({'sort_arrow_line_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
            'no_arrow_line_list': copy.deepcopy(no_arrow_line_list),
        }
    }}
    )
    progress(f'整理線段與鋼筋標示箭頭:{time.time() - start}s')

    # Step 9. 對應箭頭跟文字，並完成head_to_data_dic, tail_to_data_dic
    # start = time.time()
    # temp = {
    #     'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
    #     'coor_to_data_list': copy.deepcopy(coor_to_data_list),
    #     'middle_tie_pattern': middle_tie_pattern,
    #     'measure_type': measure_type
    # }
    try:
        coor_to_arrow_dic, \
            head_to_data_dic, \
            tail_to_data_dic = sort_arrow_to_word(coor_to_arrow_dic=coor_to_arrow_dic,
                                                  coor_to_data_list=coor_to_data_list,
                                                  **kwargs)
    except NoRebarDataError:
        print('NoRebarData')

    test_data.update({'sort_arrow_to_word_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
            'head_to_data_dic': copy.deepcopy(head_to_data_dic),
            'tail_to_data_dic': copy.deepcopy(tail_to_data_dic)
        }
    }})
    progress(f'整理鋼筋文字與鋼筋標示箭頭:{time.time() - start}s')

    # Step 10. 統計目前的type跟size

    # progress('梁配筋圖讀取進度 10/15', progress_file)
    # coor_to_rebar_list_straight_left,coor_to_rebar_list_straight_right, coor_to_bend_rebar_list,no_concat_line_list,no_concat_bend_list=concat_no_arrow_line(no_arrow_line_list=no_arrow_line_list,
    #                                                                                                                 head_to_data_dic=head_to_data_dic,
    #                                                                                                                 tail_to_data_dic=tail_to_data_dic,
    #                                                                                                                 coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 12. 拿彎的去找跟誰接在一起
    start = time.time()
    temp = {
        'no_concat_line_list': copy.deepcopy(no_arrow_line_list),
        'head_to_data_dic': copy.deepcopy(head_to_data_dic),
        'tail_to_data_dic': copy.deepcopy(tail_to_data_dic)
    }
    coor_to_rebar_list_straight = sort_noconcat_line(no_concat_line_list=no_arrow_line_list,
                                                     head_to_data_dic=head_to_data_dic,
                                                     tail_to_data_dic=tail_to_data_dic)
    test_data.update({'sort_noconcat_line_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_rebar_list_straight': copy.deepcopy(coor_to_rebar_list_straight),
        }
    }})

    temp = {
        'no_concat_bend_list': copy.deepcopy(coor_to_bend_rebar_list),
        'head_to_data_dic': copy.deepcopy(head_to_data_dic),
        'tail_to_data_dic': copy.deepcopy(tail_to_data_dic)
    }
    coor_to_bend_rebar_list = sort_noconcat_bend(no_concat_bend_list=coor_to_bend_rebar_list,
                                                 head_to_data_dic=head_to_data_dic,
                                                 tail_to_data_dic=tail_to_data_dic)
    test_data.update({'sort_noconcat_bend_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_bend_rebar_list': copy.deepcopy(coor_to_bend_rebar_list),
        }
    }})

    temp = {
        'rebar_bend_list': copy.deepcopy(coor_to_bend_rebar_list),
        'rebar_line_list': copy.deepcopy(coor_to_rebar_list_straight),
    }
    sort_rebar_bend_line(rebar_bend_list=coor_to_bend_rebar_list,
                         rebar_line_list=coor_to_rebar_list_straight)

    test_data.update({'sort_rebar_bend_line_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_bend_rebar_list': copy.deepcopy(coor_to_bend_rebar_list),
            'coor_to_rebar_list_straight': copy.deepcopy(coor_to_rebar_list_straight)
        }
    }})
    # 截斷處重複計算
    progress(f'整理彎折鋼筋與直線鋼筋:{time.time() - start}s')

    # Step 14-15 和 16 為箍筋部分，14-15在算框框內的數量，16在算每個梁的總長度，兩者獨立
    # count_rebar_in_block(coor_to_arrow_dic,coor_to_block_list,coor_to_rebar_list_straight=coor_to_rebar_list_straight,coor_to_bend_rebar_list=coor_to_bend_rebar_list)
    # Step 14. 算箍筋
    start = time.time()
    temp = {
        'coor_to_tie_text_list': copy.deepcopy(coor_to_tie_text_list),
        'coor_to_tie_list': copy.deepcopy(coor_to_tie_list),
    }
    coor_sorted_tie_list = count_tie(coor_to_tie_text_list=coor_to_tie_text_list,
                                     coor_to_tie_list=coor_to_tie_list,
                                     **kwargs)
    test_data.update({'count_tie_data': {
        'inputs': temp,
        'outputs': {
            'coor_sorted_tie_list': copy.deepcopy(coor_sorted_tie_list),
        }
    }})
    temp = {
        'coor_to_beam_list': copy.deepcopy(coor_to_beam_list),
        'floor_list': copy.deepcopy(floor_list),
        'class_beam_list': []
    }
    class_beam_list = []
    add_beam_to_list(coor_to_beam_list=coor_to_beam_list,
                     class_beam_list=class_beam_list,
                     floor_list=floor_list,
                     **kwargs)
    test_data.update({'add_beam_to_list_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_beam_list': copy.deepcopy(coor_to_beam_list),
            'class_beam_list': copy.deepcopy(class_beam_list),
            'floor_list': copy.deepcopy(floor_list),
        }
    }})
    progress(f'新增梁:{time.time() - start}s')
    start = time.time()
    temp = {
        'coor_to_block_list': copy.deepcopy(coor_to_block_list),
        'coor_to_bounding_block_list': copy.deepcopy(coor_to_bounding_block_list),
        'class_beam_list': copy.deepcopy(class_beam_list),
        'coor_to_rc_block_list': copy.deepcopy(coor_to_rc_block_list),
    }
    combine_beam_boundingbox(coor_to_block_list=coor_to_block_list,
                             coor_to_bounding_block_list=coor_to_bounding_block_list,
                             class_beam_list=class_beam_list,
                             coor_to_rc_block_list=coor_to_rc_block_list,
                             **kwargs)
    test_data.update({'combine_beam_boundingbox_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_block_list': copy.deepcopy(coor_to_block_list),
            'coor_to_bounding_block_list': copy.deepcopy(coor_to_bounding_block_list),
            'class_beam_list': copy.deepcopy(class_beam_list),
            'coor_to_rc_block_list': copy.deepcopy(coor_to_rc_block_list),
        }
    }})
    progress(f'整理梁邊界與邊框:{time.time() - start}s')
    start = time.time()
    temp = {
        'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
        'class_beam_list': copy.deepcopy(class_beam_list),
    }
    break_down_rebar(coor_to_arrow_dic=coor_to_arrow_dic,
                     class_beam_list=class_beam_list)
    test_data.update({'break_down_rebar_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
            'class_beam_list': copy.deepcopy(class_beam_list),
        }
    }})
    progress(f'截斷直線鋼筋:{time.time() - start}s')
    start = time.time()
    temp = {
        'coor_sorted_tie_list': copy.deepcopy(coor_sorted_tie_list),
        'class_beam_list': copy.deepcopy(class_beam_list),
        'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
        'coor_to_rebar_list_straight': copy.deepcopy(coor_to_rebar_list_straight),
        'coor_to_bend_rebar_list': copy.deepcopy(coor_to_bend_rebar_list)
    }
    combine_beam_tie(coor_sorted_tie_list=coor_sorted_tie_list,
                     class_beam_list=class_beam_list,
                     **kwargs)
    combine_beam_rebar(coor_to_arrow_dic=coor_to_arrow_dic,
                       coor_to_rebar_list_straight=coor_to_rebar_list_straight,
                       coor_to_bend_rebar_list=coor_to_bend_rebar_list,
                       class_beam_list=class_beam_list,
                       ** kwargs)
    test_data.update({'combine_beam_data': {
        'inputs': temp,
        'outputs': {
            'coor_sorted_tie_list': copy.deepcopy(coor_sorted_tie_list),
            'class_beam_list': copy.deepcopy(class_beam_list),
            'coor_to_arrow_dic': copy.deepcopy(coor_to_arrow_dic),
            'coor_to_rebar_list_straight': copy.deepcopy(coor_to_rebar_list_straight),
            'coor_to_bend_rebar_list': copy.deepcopy(coor_to_bend_rebar_list)
        }
    }})
    progress(f'配對梁與主筋、箍筋:{time.time() - start}s')
    start = time.time()
    temp = {
        'coor_to_dim_list': copy.deepcopy(coor_to_dim_list),
        'class_beam_list': copy.deepcopy(class_beam_list),
        'coor_to_block_list': copy.deepcopy(coor_to_block_list)
    }
    compare_line_with_dim(class_beam_list=class_beam_list,
                          coor_to_dim_list=coor_to_dim_list,
                          coor_to_block_list=coor_to_block_list)
    test_data.update({'compare_line_with_dim_data': {
        'inputs': temp,
        'outputs': {
            'coor_to_dim_list': copy.deepcopy(coor_to_dim_list),
            'class_beam_list': copy.deepcopy(class_beam_list),
            'coor_to_block_list': copy.deepcopy(coor_to_block_list)
        }
    }})
    progress(f'配對梁主筋與標註線:{time.time() - start}s')
    start = time.time()
    temp = {
        'class_beam_list': copy.deepcopy(class_beam_list),
    }
    sort_beam(class_beam_list=class_beam_list,
              middle_tie_type=kwargs['middle_tie_pattern']['type'])
    test_data.update({'sort_beam_data': {
        'inputs': temp,
        'outputs': {
            'class_beam_list': copy.deepcopy(class_beam_list),
        }
    }})
    progress(f'整理梁配筋:{time.time() - start}s')
    save_temp_file.save_pkl(
        test_data, r'tests\data\test-data-4.pkl'
    )

    assign_floor_prop(class_beam_list,
                      parameter_df=parameter_df)

    return class_beam_list, cad_data


def assign_floor_prop(beam_list: list[Beam],
                      parameter_df: pd.DataFrame):
    '''
    #### assign excel floor parameter to beam
    '''
    floor_list: list[Floor] = []
    parameter_df.set_index(['樓層'], inplace=True)

    for floor_name in parameter_df.index:
        floor_name: str
        temp_floor = Floor(str(floor_name))
        floor_list.append(temp_floor)
        temp_floor.set_beam_prop(parameter_df.loc[floor_name])

        current_floor_beam = [b for b in beam_list
                              if b.floor == floor_name or
                              b.floor.replace('F', '') == floor_name.replace('F', '')]
        temp_floor.add_beam(current_floor_beam)
        # for beam in current_floor_beam:

        #     beam.set_prop(temp_floor)

    return floor_list


def cal_beam_in_plan(beam_list: list[Beam],
                     plan_filename: str,
                     plan_layer_config: dict,
                     plan_pkl: str = ''):
    plan_floor_count = sort_plan_count(plan_filename=plan_filename,
                                       layer_config=plan_layer_config,
                                       plan_pkl=plan_pkl)
    for beam in beam_list:
        if beam.floor in plan_floor_count:
            if beam.serial in plan_floor_count[beam.floor]:
                beam.plan_count = plan_floor_count[beam.floor][beam.serial]
                continue
        beam.plan_count = 1


def create_report(class_beam_list: list[Beam],
                  output_folder: str,
                  project_name: str,
                  floor_parameter_xlsx: str,
                  cad_data: Counter,
                  plan_filename: str = '',
                  plan_layer_config: dict = None,
                  plan_pkl: str = '',
                  output_beam_type: list[Literal['GB', 'SB', 'FB']] = None):

    if output_beam_type is None:
        output_beam_type = []
    progress('產生報表')
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
    dxf_file_name = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'redraw.dxf'
    )
    output_file_list = []
    pdf_GB_file = ''
    pdf_FB_file = ''
    pdf_SB_file = ''

    # remove duplicate beam in dwg
    remove_duplicates_beams(class_beam_list)

    if plan_filename or plan_pkl:
        cal_beam_in_plan(beam_list=class_beam_list,
                         plan_filename=plan_filename,
                         plan_layer_config=plan_layer_config,
                         plan_pkl=plan_pkl)

    cad_df = pd.DataFrame.from_dict(
        data=cad_data, orient='index', columns=['數量'])

    parameter_df = read_parameter_df(floor_parameter_xlsx, '梁參數表')
    floor_list = assign_floor_prop(class_beam_list, parameter_df)

    # floor_list = floor_parameter(
    #     beam_list=class_beam_list, floor_parameter_xlsx=floor_parameter_xlsx)
    beam_df = output_beam(class_beam_list=class_beam_list)
    fbeam_list, sbeam_list, beam_list = seperate_beam(
        class_beam_list=class_beam_list)

    bs_list = create_beam_scan()
    sb_bs_list = create_sbeam_scan()
    fb_bs_list = create_fbeam_scan()
    if beam_list:
        pdf_GB_file = create_scan_report(floor_list=floor_list,
                                         beam_list=beam_list,
                                         bs_list=bs_list,
                                         excel_filename=excel_filename,
                                         beam_type=BeamType.Grider,
                                         project_name=project_name,
                                         output_folder=output_folder)
        output_file_list.extend(pdf_GB_file)

        if 'GB' in output_beam_type:
            return

    if fbeam_list:
        pdf_FB_file = create_scan_report(floor_list=floor_list,
                                         beam_list=fbeam_list,
                                         bs_list=fb_bs_list,
                                         excel_filename=excel_filename,
                                         beam_type=BeamType.FB,
                                         project_name=project_name,
                                         output_folder=output_folder)
        output_file_list.extend(pdf_FB_file)
        if 'FB' in output_beam_type:
            return
    if sbeam_list:
        pdf_SB_file = create_scan_report(floor_list=floor_list,
                                         beam_list=sbeam_list,
                                         bs_list=sb_bs_list,
                                         excel_filename=excel_filename,
                                         beam_type=BeamType.SB,
                                         project_name=project_name,
                                         output_folder=output_folder)
        output_file_list.extend(pdf_SB_file)
        if 'SB' in output_beam_type:
            return
    rebar_df, concrete_df, coupler_df, formwork_df, _ = summary_floor_rebar(
        floor_list=floor_list, item_type='beam')

    # rcad_df = output_rcad_beam(class_beam_list=class_beam_list)

    OutputExcel(df_list=[beam_df], file_path=excel_filename, sheet_name='梁統整表')

    OutputExcel(df_list=[rebar_df],
                file_path=excel_filename, sheet_name='鋼筋統計表')
    OutputExcel(df_list=[concrete_df],
                file_path=excel_filename, sheet_name='混凝土統計表')
    OutputExcel(df_list=[formwork_df],
                file_path=excel_filename, sheet_name='模板統計表')

    OutputExcel(df_list=[cad_df], file_path=excel_filename,
                sheet_name='CAD統計表')
    # OutputExcel(df_list=[rcad_df],
    #             file_path=excel_filename_rcad, sheet_name='RCAD撿料')
    output_file_list.append(excel_filename)
    output_file_list.append(excel_filename_rcad)

    dxf_file_name = draw_beam_rebar_dxf(output_folder=output_folder,
                                        beam_list=class_beam_list,
                                        dxf_file_name=dxf_file_name)

    output_file_list.append(dxf_file_name)

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

    now = time.strftime("%Y%m%d_%H%M%S", time.localtime())

    pdf_report = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{now}_'
        f'{item_name}_report.pdf'
    )

    pdf_report_appendix = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{now}_'
        f'{item_name}_appendix.pdf'
    )

    enoc_df, code_df = beam_check(beam_list=beam_list, beam_scan_list=bs_list)
    header_list, ratio_dict, ratio_df = summary_floor_rebar_ratio(floor_list=floor_list,
                                                                  beam_type=beam_type)
    rebar_df, concrete_df, coupler_df, formwork_df, detail_report = summary_floor_rebar(floor_list=floor_list,
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
                    item_name=item_name,
                    detail_report=detail_report,
                    appendix=pdf_report_appendix)
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
    return [pdf_report, pdf_report_appendix]


def seperate_beam(class_beam_list: list[Beam]):
    '''
    Seperate Beam into GB,SB,FB
    '''
    return [b for b in class_beam_list if b.beam_type == BeamType.FB], \
        [b for b in class_beam_list if b.beam_type == BeamType.SB], \
        [b for b in class_beam_list if b.beam_type == BeamType.Grider]


def remove_duplicates_beams(class_beam_list: list[Beam]):
    '''
    Remove Duplicates
    '''
    exists_beam = set()
    for beam in class_beam_list[:]:
        if (beam.floor, beam.serial, beam.width, beam.depth) in exists_beam:
            class_beam_list.remove(beam)
            progress(
                f'Remove {(beam.floor , beam.serial , beam.width , beam.depth)}')
        else:
            exists_beam.add((beam.floor, beam.serial, beam.width, beam.depth))


def add_beam_to_list(coor_to_beam_list: list,
                     class_beam_list: list,
                     floor_list: list,
                     measure_type: str,
                     name_pattern: dict = None,
                     size_pattern: dict = None,
                     **kwargs):
    '''
    Convert Beam Text to a Beam Object Based on
    - 0.first sub the floor text with "General" to prevent special prefix
    - 1.floor match floor pattern
    - 2.serial match name pattern (Grider >> FB >> SB )
    '''
    def regex_beam_string(pattern, beam_string):
        if pattern is None:
            return beam_string
        return re.sub(pattern, '', beam_string)

    if name_pattern is None:
        error('No name pattern in kwargs at add_beam_to_list')
        return

    if size_pattern is None:
        error('No size pattern in kwargs at add_beam_to_list')
        return

    floor_pattern = r'(\d+F|R\d+|PR|BS|B\d+|MF|RF|PF|FS)F*'

    for beam in coor_to_beam_list:
        try:
            serial = regex_beam_string(
                name_pattern.get('General', None), beam[0])
            b = Beam(serial, beam[1][0], beam[1][1])

            b: Beam = b.get_beam_info(floor_list=floor_list,
                                      measure_type=measure_type,
                                      name_pattern=name_pattern,
                                      floor_pattern=floor_pattern,
                                      size_pattern=size_pattern)
        except BeamFloorNameError:
            error(f'{beam} beam serial error at add_beam_to_list')
            continue
        if b is not None:
            class_beam_list.append(b)
        else:
            error(f'{beam} floor name error at add_beam_to_list')


def draw_rebar_line(class_beam_list: list[Beam],
                    msp_beam: object,
                    doc_beam: object,
                    output_folder: str,
                    project_name: str):
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
                        error(f'{beam.floor}{beam.serial} write rebar error')
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
                    error(f'{beam.floor}{beam.serial} write tie error')
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
                    error(f'{beam.floor}{beam.serial} write middle tie error')
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
            error(f'{beam.floor}{beam.serial} write block error')
            error_count += 1
            time.sleep(5)
            # coor_list2 = [beam.coor.x, beam.coor.y, 0, rebar.end_pt.x, rebar.end_pt.y, 0]
        # coor_list2 = [min_right_coor[0], min_right_coor[1], 0, x[1][0], x[1][1], 0]
    try:
        doc_beam.SaveAs(output_dwg)
        doc_beam.Close(SaveChanges=True)
    except:
        error(f'{project_name} cant save file')
    return output_dwg


def sort_beam(class_beam_list: list[Beam],
              **kwargs):
    for beam in class_beam_list:
        beam.sort_beam_rebar()
        beam.sort_beam_tie()
        if kwargs.get('middle_tie_type', '') == 'multi':
            beam.sort_middle_tie()

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
            if i > 0:
                beam.at[row + 1, pos] = f'{beam.at[row + 1,pos]}+{text}'
            else:
                beam.at[row - 1, pos] = f'{beam.at[row - 1,pos]}+{text}'
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
        row += 4

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


def sort_dim_line_and_text(coor_to_dim_text: list,
                           coor_to_dim_line: list,
                           **kwargs):
    '''
    coor_to_dim_text = [[object.TextString, midpoint]]
    coor_to_dim_line = [[(coor1,coor2), '', text_coor]]

    output = (object.TextPosition, object.Measurement, (coor1, coor2))
    '''
    output = []
    min_diff = kwargs.get('sort_dim_line_and_text_min_diff', 300)

    for dim_text in coor_to_dim_text:
        text, point = dim_text
        try:
            _ = int(text)
        except ValueError:  # text must be a "int-type"
            continue

        # find closet dim line (bewtween line and above)
        dim_line = [line for line in coor_to_dim_line if (line[0][0][0] - point[0]) * (line[0][1][0] - point[0]) < 0 and
                    0 < (point[1] - line[0][0][1]) < min_diff]

        if not dim_line:
            continue

        min_dim_line = min(dim_line, key=lambda l: (
            abs(point[1] - l[0][0][1]), abs((l[0][0][0] + l[0][1][0]) / 2 - point[0])))

        while min_dim_line[1] != '':
            pre_text = min_dim_line[1]
            pre_coor = min_dim_line[2]
            if abs(min_dim_line[0][0][1] - point[1]) < abs(min_dim_line[0][0][1] - pre_coor[1]):
                min_dim_line[1] = text
                min_dim_line[2] = point

                text = pre_text
                point = pre_coor
                dim_line = [line for line in coor_to_dim_line if (line[0][0][0] - point[0]) * (line[0][1][0] - point[0]) < 0 and
                            0 < (point[1] - line[0][1]) < min_diff]

                if not dim_line:
                    break

                min_dim_line = min(dim_line, key=lambda l: (
                    abs(point[1] - l[0][0][1]), abs((l[0][0][0] + l[0][1][0]) / 2 - point[0])))
            else:
                break

        if min_dim_line and min_dim_line[1] == '':
            min_dim_line[1] = text
            min_dim_line[2] = point

    for dim_line in coor_to_dim_line:
        if dim_line[1] != '':
            output.append((dim_line[2], float(dim_line[1]), dim_line[0]))

    return output


def count_beam_multiprocessing(beam_filenames: list,
                               layer_config: dict,
                               temp_file='temp_1221_1F.pkl',
                               output_folder='',
                               project_name='',
                               template_name: Literal["ELEMENTS",
                                                      "DRAWRC", "RCAD", "OTHER"] = '',
                               floor_parameter_xlsx='',
                               progress_file='',
                               plan_filename='',
                               plan_layer_config='',
                               client_id="temp",
                               measure_type='cm',
                               redraw=False):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    if progress_file == '':
        progress_file = './result/tmp'
    cad_counter = Counter()
    output_dwg_list = []
    output_dwg = ''

    def read_beam_multi(beam_filename, temp_file):
        msp_beam, doc_beam = read_beam_cad(
            beam_filename=beam_filename)
        sort_beam_cad(msp_beam=msp_beam,
                      doc_beam=doc_beam,
                      layer_config=layer_config,
                      entity_config=get_template(template_name),
                      temp_file=temp_file,
                      redraw=redraw)
        output_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(temp_file),
                                                    rebar_parameter_excel=floor_parameter_xlsx,
                                                    measure_type=measure_type)

        output_dwg = ''

        if redraw:
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
                                     progress_file=progress_file,
                                     plan_filename=plan_filename,
                                     plan_layer_config=plan_layer_config)
    end = time.time()
    progress(f'執行時間：{end - start}s')
    progress("執行時間：%f 秒" % (end - start))

    output_file_list = [os.path.basename(file) for file in output_file_list]
    return output_file_list, output_dwg_list, f'{temp_new}-beam_list.pkl'


def get_template(name: Literal["ELEMENTS", "DRAWRC", "RCAD", "OTHER"]):
    name = name.upper()
    if name == 'ELEMENTS':
        return {
            'rebar_layer': ['AcDbPolyline'],
            'rebar_data_layer': ['AcDbMText'],
            'rebar_data_leader_layer': ['AcDbLeader'],
            'tie_text_layer': ['AcDbText']
        }
    if name == 'DRAWRC':
        return {
            'rebar_layer': ['AcDbLine'],
            'rebar_data_layer': ['AcDbText', 'AcDbMText'],
            'rebar_data_leader_layer': ['AcDbPolyline'],
            'tie_text_layer': ['AcDbMText']
        }
    if name == 'OTHER':
        return {
            'rebar_layer': ['AcDbPolyline'],
            'rebar_data_layer': ['AcDbText'],
            'rebar_data_leader_layer': ['AcDbPolyline'],
            'tie_text_layer': ['AcDbText']
        }
    if name == 'RCAD':
        return {
            'rebar_layer': ['AcDbPolyline'],
            'rebar_data_layer': ['AcDbMText', 'AcDbText'],
            'rebar_data_leader_layer': ['AcDbLeader', 'AcDbPolyline'],
            'tie_text_layer': ['AcDbText']
        }


def floor_parameter(beam_list: list[Beam],
                    floor_parameter_xlsx: str):
    parameter_df: pd.DataFrame
    # floor_list: list[Floor]
    # floor_list = []

    parameter_df = read_parameter_df(floor_parameter_xlsx, '梁參數表')
    return assign_floor_prop(beam_list, parameter_df)
    # parameter_df.set_index(['樓層'], inplace=True)

    # assign_floor_prop

    # for floor_name in parameter_df.index:
    #     temp_floor = Floor(str(floor_name))
    #     floor_list.append(temp_floor)
    #     temp_floor.set_beam_prop(parameter_df.loc[floor_name])
    #     temp_floor.add_beam([b for b in beam_list if
    #                          b.floor == temp_floor.floor_name or
    #                          b.floor.replace('F', '') ==
    #                          temp_floor.floor_name.replace('F', '')])
    # return floor_list

# combine dim with text arrow


def compare_line_with_dim(coor_to_dim_list: list[tuple[tuple[float, float], float, tuple[tuple[float, float], tuple[float, float]]]],
                          class_beam_list: list[Beam],
                          coor_to_block_list: list[tuple, dict[str, str], dict[str, str]]):
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
            for i, beam in enumerate(beam_list):
                if i % 100 == 0:
                    progress(f"比較鋼筋與標註線 進度:{i}/{len(beam_list)}")
                for rebar in beam.rebar_list:
                    try:
                        temp_dim = [dim for dim in dim_list if round(rebar.start_pt.x, 2) ==
                                    round(dim[2][0][0], 2) and round(rebar.end_pt.x, 2) == round(dim[2][1][0], 2)]
                        if temp_dim:
                            continue  # 如果dim match line的話則不修正

                        rebar_dim = [dim for dim in dim_list if ((rebar.start_pt.x - dim[0][0]) * (rebar.end_pt.x - dim[0][0]) <= 0
                                                                 and abs(dim[0][1] - rebar.arrow_coor[0][1]) < min_diff)]
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
                                f'{rebar.start_pt} :{rebar.length} <> {min_dim[0]}:{min_dim[1]}')
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
                    except Exception:
                        progress(
                            f'{beam.floor}:{beam.serial} compare_line_with_dim error')
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


def read_parameter_json(template_name):

    filename = f'file\\parameter\\{template_name}.json'
    if os.path.exists(filename):
        with open(filename, "r", encoding='utf-8') as f:
            parameter = json.load(f)
        return parameter
    return {}


def count_beam_multifiles(project_name: str,
                          beam_filenames: list,
                          floor_parameter_xlsx: str,
                          pkl_file_folder: str,
                          output_folder: str,
                          layer_config: dict,
                          entity_type: dict,
                          name_pattern: str,
                          measure_type: Literal["cm", "mm"],
                          client_id="temp",
                          plan_filename: str = "",
                          plan_pkl: str = "",
                          plan_layer_config: dict = {},
                          **kwargs):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)

    cad_counter = Counter()
    output_dwg_list = []
    output_dwg = ''

    all_beam_list = []

    output_file_list = []

    now_time = time.strftime("%Y%m%d_%H%M%S")

    if plan_pkl:
        plan = {'plan_pkl': plan_pkl}
    elif plan_filename and plan_layer_config:
        plan = {'plan_filename': plan_filename,
                'plan_layer_config': plan_layer_config}
    else:
        plan = {}

    if kwargs.get('beam_pkl', "") != "" and kwargs.get('cad_data_pkl', "") != "" and not kwargs.get('pkl', []):

        all_beam_list = save_temp_file.read_temp(kwargs.get('beam_pkl', ""))

        cad_data = save_temp_file.read_temp(kwargs.get('cad_data_pkl', ""))

        output_file_list = create_report(class_beam_list=all_beam_list,
                                         floor_parameter_xlsx=floor_parameter_xlsx,
                                         output_folder=output_folder,
                                         project_name=project_name,
                                         cad_data=cad_counter,
                                         output_beam_type=kwargs.get(
                                             'beam_type', []),
                                         **plan)
        return output_file_list

    if kwargs.get('pkl', []):

        result_pkl = f'{pkl_file_folder}/{project_name}-{now_time}-beam-object.pkl'

        for i, filename in enumerate(kwargs.get('pkl', [])):

            base_filename_without_ext = os.path.splitext(
                os.path.basename(filename))[0]

            class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(filename),
                                                       rebar_parameter_excel=floor_parameter_xlsx,
                                                       name_pattern=name_pattern,
                                                       measure_type=measure_type,
                                                       **kwargs)
            all_beam_list.extend(class_beam_list)

            cad_counter.update(cad_data)

        save_temp_file.save_pkl(
            all_beam_list, tmp_file=result_pkl)
        save_temp_file.save_pkl(
            cad_data, tmp_file=f'{pkl_file_folder}/{project_name}-{now_time}-cad-data.pkl')

        if kwargs.get('beam_pkl', "") != "":
            pkl_beam_list = save_temp_file.read_temp(
                kwargs.get('beam_pkl', ""))
            all_beam_list.extend(pkl_beam_list)
        if all_beam_list:
            output_file_list = create_report(class_beam_list=all_beam_list,
                                             floor_parameter_xlsx=floor_parameter_xlsx,
                                             output_folder=output_folder,
                                             project_name=project_name,
                                             cad_data=cad_counter,
                                             **plan)
        return

    for i, filename in enumerate(beam_filenames):

        base_filename_without_ext = os.path.splitext(
            os.path.basename(filename))[0]
        try:
            tmp_file = f'{pkl_file_folder}/{project_name}-{now_time}-{base_filename_without_ext}-beam-data-{i}.pkl'
            msp_beam = None
            doc_beam = None

            if not os.path.exists(tmp_file):
                msp_beam, doc_beam = read_beam_cad(
                    beam_filename=filename)

                sort_beam_cad(msp_beam=msp_beam,
                              doc_beam=doc_beam,
                              layer_config=layer_config,
                              entity_config=entity_type,
                              temp_file=tmp_file,
                              redraw=True)

            class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(tmp_file),
                                                       rebar_parameter_excel=floor_parameter_xlsx,
                                                       name_pattern=name_pattern,
                                                       measure_type=measure_type,
                                                       **kwargs)
            if msp_beam and doc_beam:

                output_dwg = draw_rebar_line(class_beam_list=class_beam_list,
                                             msp_beam=msp_beam,
                                             doc_beam=doc_beam,
                                             output_folder=output_folder,
                                             project_name=f'{project_name}_{i}_{base_filename_without_ext}')

                output_dwg_list.append(os.path.basename(output_dwg))

            all_beam_list.extend(class_beam_list)

            cad_counter.update(cad_data)

        except Exception as ex:
            print(f'{filename} error {ex}')
            save_temp_file.save_pkl(
                all_beam_list, tmp_file=f'{pkl_file_folder}/{project_name}-{now_time}-beam-object-{i}.pkl')

    result_pkl = f'{pkl_file_folder}/{project_name}-{now_time}-beam-object-all.pkl'

    save_temp_file.save_pkl(
        all_beam_list, tmp_file=result_pkl)
    save_temp_file.save_pkl(
        cad_data, tmp_file=f'{pkl_file_folder}/{project_name}-{now_time}-cad-data.pkl')

    if all_beam_list:

        output_file_list = create_report(class_beam_list=all_beam_list,
                                         floor_parameter_xlsx=floor_parameter_xlsx,
                                         output_folder=output_folder,
                                         project_name=project_name,
                                         cad_data=cad_counter,
                                         **plan)
        save_temp_file.save_pkl(
            all_beam_list, tmp_file=result_pkl)
    output_file_list = [os.path.basename(file) for file in output_file_list]
    return output_file_list, output_dwg_list, result_pkl


if __name__ == '__main__':
    from main import GetAllFiles
    import json
    # beam_pkl_files = GetAllFiles(
    #     r'D:\Desktop\BeamQC\TEST\2024-1021\梁pkl', ext="*.pkl")
    # parameter = read_parameter_json('廍子')['beam']
    parameter = read_parameter_json('Elements')['beam']
    parameter['measure_type'] = "cm"
    # count_beam_multifiles(
    #     project_name='廍子社宅',
    #     beam_filenames=[],
    #     # beam_filenames=[
    #     #     r'D:\Desktop\BeamQC\TEST\2024-1021\梁\S2-B28_地下層大梁配筋圖.dwg'],
    #     floor_parameter_xlsx=r'TEST\2024-1021\floor.xlsx',
    #     pkl_file_folder=r'D:\Desktop\BeamQC\TEST\2024-1021',
    #     output_folder=r'D:\Desktop\BeamQC\TEST\2024-1021',
    #     # pkl=beam_pkl_files,
    #     # pkl=[r'TEST\2024-1021\梁pkl\廍子社宅-20241021_165328-S2-B28_地下層大梁配筋圖-beam-data-0.pkl'],
    #     #      r'TEST\2024-1011\SCAN\沙崙社宅-20241018_171452-2024-1018 沙崙社宅 下構大梁粗略配筋-beam-data-1.pkl'],
    #     # plan_pkl=r'TEST\2024-1021\2024-1021-2024-10-21-14-06-temp_plan_count_set.pkl',
    #     beam_pkl=r'TEST\2024-1021\廍子社宅-20241022_145715-beam-object.pkl',
    #     cad_data_pkl=r'TEST\2024-1011\2024-1011-20241011_155100-cad-data.pkl',
    #     # beam_type=['GB'],
    #     **parameter
    # )

    count_beam_multifiles(
        project_name='沙崙社宅',
        beam_filenames=[
            r'D:\Desktop\BeamQC\TEST\2024-1011\SCAN\2024-1018 沙崙社宅 上構大梁粗略配筋.dwg'],
        floor_parameter_xlsx=r'TEST\2024-1011\SCAN\floor.xlsx',
        pkl_file_folder=r'TEST\2024-1011\SCAN',
        output_folder=r'D:\Desktop\BeamQC\TEST\2024-1011\SCAN',
        beam_pkl=r'TEST\2024-1011\SCAN\沙崙社宅-20241022_155028-beam-object-all.pkl',
        cad_data_pkl=r'TEST\2024-1011\SCAN\沙崙社宅-20241022_155028-cad-data.pkl',
        ** parameter
    )
    # from multiprocessing import Process, Pool
    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    # beam_filename = r"D:\Desktop\BeamQC\TEST\INPUT\2022-11-18-17-16temp-XS-BEAM.dwg"#sys.argv[1] # XS-BEAM的路徑
    # beam_filename = r"D:\Desktop\BeamQC\TEST\2024-0830\梁\2024-0905\11002_S2302_一層大梁配筋圖.dwg"
    # beam_filenames = [r"D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S2A結構FB_1120829.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S2B結構B0_1120821.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S2D結構SB_1120821.dwg",
    #                   #   r"D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S3結構C0_1120829.dwg",
    #                   r"D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S2C結構B0_1120829.dwg"]

    # progress_file = './result/tmp'  # sys.argv[14]
    # output_folder = r'D:\Desktop\BeamQC\TEST\2024-0923'
    # # floor_parameter_xlsx = r'D:\Desktop\BeamQC\TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-floor.xlsx'
    # floor_parameter_xlsx = r'D:\Desktop\BeamQC\TEST\2024-0923\P2022-04A 國安社宅二期暨三期22FB4-2024-09-23-11-32-floor_1.xlsx'
    # project_name = '2024-0923'
    # plan_filename = r''
    # plan_layer_config = {
    #     'block_layer': ['AREA'],
    #     'name_text_layer': ['BTXT', 'CTXT', 'BTXT_S_'],
    #     'floor_text_layer': ['TEXT1']
    # }
    # plan_layer_config = {
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],
    #     'name_text_layer': ['S-TEXTG', 'S-TEXTB', 'S-TEXTC'],
    #     'floor_text_layer': ['S-TITLE']
    # }
    # 在beam裡面自訂圖層
    # layer_config = {
    #     'rebar_data_layer': ['S-LEADER'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['S-REINF'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['S-TEXT'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],  # 框框的圖層
    #     'beam_text_layer': ['S-RC'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['S-RC'],  # 支承端圖層
    #     's_dim_layer': ['S-DIM'],  # 標註線圖層
    #     'burst_layer_list': []
    # }
    # 在beam裡面自訂圖層
    # layer_config = {
    #     'rebar_data_layer': ['BARNOTE'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['BARA', 'BARS', 'BART'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['BARNOTE'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],  # 框框的圖層
    #     'beam_text_layer': ['TITLE'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['S-RC', 'OLINE'],  # 支承端圖層
    #     's_dim_layer': ['DIMS'],  # 標註線圖層
    #     'burst_layer_list': ['XREF']
    # }

    # layer_config = {
    #     'rebar_data_layer': ['P1'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['P2'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['P1'],  # 箍筋文字圖層
    #     'block_layer': ['0'],  # 框框的圖層
    #     'beam_text_layer': ['P2'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['P1'],  # 支承端圖層
    #     's_dim_layer': ['P1']  # 標註線圖層
    # }
    # layer_config = {
    #     'rebar_data_layer': ['主筋文字', '主筋文字引線', '扭力筋文字', '扭力筋文字引線', '扭力筋文字引線 no use'],
    #     'rebar_layer': ['主筋線', '剪力筋線', '扭力筋線'],
    #     'tie_text_layer': ['剪力筋文字'],
    #     'block_layer': ['_noprint'],
    #     'beam_text_layer': ['梁跨名稱'],
    #     'bounding_block_layer': [''],
    #     'rc_block_layer': ['邊界線', '梁柱截斷記號', '邊界線-梁支撐內線', '邊界線-梁支撐外線'],
    #     's_dim_layer': ['尺寸標註-字串']
    # }

    # layer_config = {
    #     'rebar_data_layer': ['NBAR'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['RBAR'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['NBAR'],  # 箍筋文字圖層
    #     'block_layer': ['DEFPOINTS'],  # 框框的圖層
    #     'beam_text_layer': ['TITLE'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['OLINE'],
    #     's_dim_layer': ['DIMS']
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

    # 2024-0903
    # layer_config = {
    #     'rebar_data_layer': ['TEXT-2', 'CT-1'],  # 箭頭和鋼筋文字的塗層
    #     'rebar_layer': ['CT-3', 'CT-2'],  # 鋼筋和箍筋的線的塗層
    #     'tie_text_layer': ['TEXT-2'],  # 箍筋文字圖層
    #     'block_layer': ['0', 'DwFm', 'DEFPOINTS'],  # 框框的圖層
    #     'beam_text_layer': ['TEXT-3'],  # 梁的字串圖層
    #     'bounding_block_layer': ['S-ARCH'],
    #     'rc_block_layer': ['CT-1', 'OLINE'],  # 支承端圖層
    #     's_dim_layer': ['TEXT-2', 'CT-1'],  # 標註線圖層
    #     'burst_layer_list': ['XREF']
    # }

    # entity_type ={
    #     'rebar_layer':['AcDbPolyline'],
    #     'rebar_data_layer':['AcDbText'],
    #     'rebar_data_leader_layer':['AcDbPolyline'],
    #     'tie_text_layer':['AcDbText']
    # }

    # Elements
    # entity_type = {
    #     'rebar_layer': ['AcDbPolyline'],
    #     'rebar_data_layer': ['AcDbMText', 'AcDbText'],
    #     'rebar_data_leader_layer': ['AcDbLeader', 'AcDbPolyline'],
    #     'tie_text_layer': ['AcDbText']
    # }

    # entity_type ={
    #     'rebar_layer':['AcDbLine'],
    #     'rebar_data_layer':['AcDbText'],
    #     'rebar_data_leader_layer':['AcDbLine'],
    #     'tie_text_layer':['AcDbText']
    # }

    # entity_type = {
    #     'rebar_layer': ['AcDbLine'],
    #     'rebar_data_layer': ['AcDbText', 'AcDbMText'],
    #     'rebar_data_leader_layer': ['AcDbPolyline', 'AcDbLeader'],
    #     'tie_text_layer': ['AcDbText', 'AcDbMText', 'AcDbRotatedDimension', 'AcDbAlignedDimension'],
    #     'rc_block_layer': ['AcDbPolyline', 'AcDbLine']
    # }

    # name_pattern = {
    #     'Grider': [r'(.*) ([G|B|E].*)'],
    #     'FB': [r'(.*) (F[G|b|B].*)'],
    #     'SB': [r'(.*) ([b|g].*)']
    # }

    # break_rebar_point = {
    #     'layer': ['CT-3'],
    #     'entity': ['AcDbArc']
    # }
    # middle_tie_pattern = r'(\d+)x(\d+)-(#\d+)'
    # middle_tie_pattern = r''

    # start = time.time()
    # main_logger = setup_custom_logger(__name__, client_id=project_name)
    # msp_beam, doc_beam = read_beam_cad(
    #     beam_filename=beam_filename)

    # temp_file = f'{output_folder}/P2022-04A 國安社宅二期暨三期22FB4-2024-09-23-11-32-temp-0.pkl'

    # with open(r'file\parameter\Elements.json', "r") as f:
    #     parameter = json.load(f)

    # layer_config = parameter['layer_config']
    # entity_type = parameter['entity_type']
    # break_rebar_point = parameter['break_rebar_point']
    # name_pattern = parameter['name_pattern']
    # middle_tie_pattern = parameter['middle_tie_pattern']
    # plan_layer_config = parameter['plan_layer_config']

    # sort_beam_cad(msp_beam=msp_beam,
    #               doc_beam=doc_beam,
    #               layer_config=layer_config,
    #               entity_config=entity_type,
    #               temp_file=temp_file,
    #               break_rebar_point=break_rebar_point,
    #               redraw=True)

    # class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(temp_file),
    #                                            rebar_parameter_excel=floor_parameter_xlsx,
    #                                            measure_type='cm',
    #                                            name_pattern=name_pattern,
    #                                            middle_tie_pattern=middle_tie_pattern)
    # draw_rebar_line(class_beam_list=class_beam_list,
    #                 msp_beam=msp_beam,
    #                 doc_beam=doc_beam,
    #                 output_folder=output_folder,
    #                 project_name=project_name)

    # from os import listdir
    # from os.path import isfile, join
    # mypath = r'D:\Desktop\BeamQC\TEST\2024-0830\梁'
    # onlyfiles = [os.path.join(mypath, f) for f in listdir(mypath) if isfile(
    #     join(mypath, f)) and os.path.splitext(f)[1] == ".dwg"]
    # all_beam_list = []

    # for i, filename in enumerate(onlyfiles):
    #     try:
    #         print(filename)
    #         tmp_file = f'{output_folder}/0903-data-{i}.pkl'
    #         msp_beam = None
    #         doc_beam = None

    #         if not os.path.exists(tmp_file):
    #             msp_beam, doc_beam = read_beam_cad(
    #                 beam_filename=filename, progress_file=progress_file)

    #             sort_beam_cad(msp_beam=msp_beam,
    #                           doc_beam=doc_beam,
    #                           layer_config=layer_config,
    #                           entity_config=entity_type,
    #                           temp_file=tmp_file,
    #                           redraw=True)

    #         class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(tmp_file),
    #                                                    rebar_parameter_excel=floor_parameter_xlsx,
    #                                                    name_pattern=name_pattern,
    #                                                    measure_type='mm')
    #         if msp_beam and doc_beam:
    #             draw_rebar_line(class_beam_list=class_beam_list,
    #                             msp_beam=msp_beam,
    #                             doc_beam=doc_beam,
    #                             output_folder=output_folder,
    #                             project_name=f'{project_name}_{i}')
    #         all_beam_list.extend(class_beam_list)
    #     except Exception as ex:
    #         print(f'{filename} error {ex}')
    #         save_temp_file.save_pkl(
    #             all_beam_list, tmp_file=f'{output_folder}/0904-beam-temp-2-{i}.pkl')

    # save_temp_file.save_pkl(
    #     class_beam_list, tmp_file=f'{output_folder}/beam-all.pkl')
    # save_temp_file.save_pkl(
    #     cad_data, tmp_file=f'{output_folder}/cad_list.pkl')

    # all_beam_list = save_temp_file.read_temp(
    #     f'{output_folder}/0904-beam-all-3.pkl')

    # cad_data = save_temp_file.read_temp(
    #     f'{output_folder}/cad_list.pkl')

    # tmp_file = f'{output_folder}/0830-data-0-2.pkl'
    # msp_beam, doc_beam = read_beam_cad(
    #     beam_filename=beam_filename, progress_file=progress_file)

    # sort_beam_cad(msp_beam=msp_beam,
    #               layer_config=layer_config,
    #               entity_config=entity_type,
    #               progress_file=progress_file,
    #               temp_file=tmp_file)

    # class_beam_list, cad_data = cal_beam_rebar(data=save_temp_file.read_temp(tmp_file),
    #                                            progress_file=progress_file,
    #                                            rebar_parameter_excel=floor_parameter_xlsx,
    #                                            measure_type='mm')
    # create_report(class_beam_list=class_beam_list,
    #               output_folder=output_folder,
    #               project_name=project_name,
    #               floor_parameter_xlsx=floor_parameter_xlsx,
    #               cad_data=cad_data,
    #               progress_file=progress_file,
    #               plan_pkl=r'TEST\2024-0923\P2022-04A 國安社宅二期暨三期22FB4-2024-09-23-11-32-XS-PLAN_plan_count_set.pkl',
    #               plan_layer_config=plan_layer_config)

    # class_beam_list = save_temp_file.read_temp(
    #     r'TEST\2024-0819\beam_list-2.pkl')

    # draw_rebar_line(class_beam_list=class_beam_list,
    #                 msp_beam=msp_beam,
    #                 doc_beam=doc_beam,
    #                 output_folder=output_folder,
    #                 project_name=project_name)
    # print(f'Total Time:{time.time() - start}')
