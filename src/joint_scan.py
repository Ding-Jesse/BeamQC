import pandas as pd
import time
import os
import time
import pythoncom
import win32com.client
import src.save_temp_file as save_temp_file
import re
import numpy as np
from main import OutputExcel
from utils.demand import calculate_column_beam_joint_shear
from item.beam import Beam
from item.column import Column
from item.floor import Floor
from item.floor import read_parameter_df
from utils.create_calculate_sheet import create_calculate_sheet
from column_count import floor_parameter as column_floor_parameter
from column_count import sort_floor_column
from beam_count import floor_parameter as beam_floor_parameter
from beam_count import read_beam_cad, sort_beam_cad, get_template, cal_beam_rebar
from plan_count import in_block, check_is_floor, sort_floor_text
from multiprocessing.pool import ThreadPool as Pool
from collections import Counter
from typing import Literal
from joint_draw import create_joint_plan_view
from enum import Enum
from src.logger import setup_custom_logger
from utils.algorithm import match_points
tol = 50
global main_logger


class UserDefineWarning(Enum):
    ColumnConnectError = "column_connect"
    ColumnJointShearFail = "shear"


class ColumnBlock:
    start: tuple
    end: tuple
    mid: tuple
    column_serial: str
    column_data: Column
    warning: list[UserWarning]

    def __init__(self, start, end) -> None:
        self.start = start
        self.end = end
        self.column_serial = None
        self.mid = ((start[0] + end[0]) / 2, (start[1] + end[1]) / 2, 0)
        self.column_data = None
        self.warning = []

    def get_corner(self) -> list:
        points = [
            *self.start,
            self.start[0], self.end[1], 0,
            *self.end,
            self.end[0], self.start[1], 0,
            *self.start,
        ]
        return points


class MlineObject:
    start: tuple
    mid: tuple
    end: tuple
    scale: float
    xy_direction: str
    beam_serial: tuple
    left_column: ColumnBlock
    right_column: ColumnBlock
    left_offset = float
    right_offset = float
    beam_data: Beam

    def __init__(self, start, mid, end, scale, xy_direction) -> None:
        self.start = start
        self.mid = mid
        self.end = end
        self.scale = scale
        self.xy_direction = xy_direction
        self.beam_serial = None
        self.left_column = None
        self.right_column = None
        self.beam_data = None

    def check(self) -> None:
        self.beam_serial = (self.beam_serial[0],
                            re.sub(r"\(.*\)", "", self.beam_serial[1]).strip(), self.beam_serial[2])

        if self.xy_direction == 'x':
            column_pos = 0
            beam_pos = 1
        elif self.xy_direction == 'y':
            column_pos = 1
            beam_pos = 0
        if self.left_column:
            if abs(self.left_column.end[column_pos] - self.start[column_pos]) > tol:
                self.left_column = None
            elif not self.left_column.column_serial:
                self.left_column = None
            else:
                if abs(self.left_column.end[column_pos] - self.start[column_pos]) > 1:
                    self.left_column.warning.append(
                        UserDefineWarning.ColumnConnectError)
                if abs(self.mid[beam_pos] - (self.left_column.start[beam_pos] + self.left_column.end[beam_pos]) / 2) < \
                        abs((self.left_column.start[beam_pos] - self.left_column.end[beam_pos])):
                    self.left_offset = self.mid[beam_pos] - \
                        (self.left_column.start[beam_pos] +
                            self.left_column.end[beam_pos]) / 2
                else:
                    self.left_column = None
        if self.right_column:
            if abs(self.right_column.start[column_pos] - self.end[column_pos]) > tol:
                self.right_column = None
            elif not self.right_column.column_serial:
                self.right_column = None
            else:
                if abs(self.right_column.start[column_pos] - self.end[column_pos]) > 1:
                    self.right_column.warning.append(
                        UserDefineWarning.ColumnConnectError)
                if abs(self.mid[beam_pos] - (self.right_column.start[beam_pos] + self.right_column.end[beam_pos]) / 2) < \
                        abs((self.right_column.start[beam_pos] - self.right_column.end[beam_pos])):
                    self.right_offset = self.mid[beam_pos] - (
                        self.right_column.start[beam_pos] + self.right_column.end[beam_pos]) / 2
                else:
                    self.right_column = None


def check_column_joint(excel_filename: str,
                       docx_filename: str,
                       column_beam_df: pd.DataFrame,
                       column_list: list[Column],
                       beam_list: list[Beam],
                       floor_list: list[Floor],
                       output_floor: list,
                       output_serial: list):
    result, no_rebar_data, column_beam_df, beams_df = calculate_column_beam_joint_shear(
        column_list=column_list, beam_list=beam_list, column_beam_df=column_beam_df)
    joint_df = pd.DataFrame(result)
    column_beam_df = sort_dataframe_by_floor(
        column_beam_df, [f.floor_name for f in floor_list], 'Floor', [c.serial for c in column_list], 'Column')
    no_rebar_data = sort_dataframe_by_floor(
        no_rebar_data, [f.floor_name for f in floor_list], '樓層')
    summary_df = summary_column_joint_df(joint_df, 'DCR')
    fine_summary_df = summary_column_joint_df(joint_df, 'Fine DCR')
    lee_summary_df = summary_column_joint_df(joint_df, 'lee_DCR')
    summary_df = sort_dataframe_by_floor(
        summary_df, [f.floor_name for f in floor_list], 'story')
    fine_summary_df = sort_dataframe_by_floor(
        fine_summary_df, [f.floor_name for f in floor_list], 'story')
    lee_summary_df = sort_dataframe_by_floor(
        lee_summary_df, [f.floor_name for f in floor_list], 'story')

    OutputExcel(df_list=[joint_df],
                file_path=excel_filename, sheet_name='梁柱接頭剪力表')
    OutputExcel(df_list=[no_rebar_data],
                file_path=excel_filename, sheet_name='錯誤表')
    OutputExcel(df_list=[column_beam_df],
                file_path=excel_filename, sheet_name='梁柱接頭統整表')
    OutputExcel(df_list=[summary_df],
                file_path=excel_filename, sheet_name='統整表')
    OutputExcel(df_list=[fine_summary_df],
                file_path=excel_filename, sheet_name='考量圍束統整表')
    OutputExcel(df_list=[lee_summary_df],
                file_path=excel_filename, sheet_name='未來式統整表')
    create_calculate_sheet(doc_filename=docx_filename, column_list=column_list,
                           output_serial=output_serial, output_floor=output_floor)
    return joint_df, beams_df


def match_rebar_data_with_object(data: dict, beams_df: pd.DataFrame, column_list: list[Column]):
    columns_df = pd.DataFrame(
        [{'floor': column.floor, 'serial': column.serial, 'column': column} for column in column_list])
    columns_df.set_index(['floor', 'serial'], inplace=True)
    for floor, items in data.items():
        mline_list: list[MlineObject] = items['mline_list']
        for mline in mline_list:
            if mline.beam_serial:
                try:
                    beam_data = beams_df.loc[(
                        floor, mline.beam_serial[1]), 'beam']
                    mline.beam_data = beam_data
                except KeyError:
                    continue
        column_block_list: list[ColumnBlock] = items['column_block_list']
        for column_block in column_block_list:
            if column_block.column_serial:
                try:
                    column_data = columns_df.loc[(
                        floor, column_block.column_serial[1]), 'column']
                    column_block.column_data = column_data
                except KeyError:
                    continue


def column_joint_main(output_folder: str,
                      project_name: str,
                      beam_pkl: str,
                      column_pkl: str,
                      column_beam_df: pd.DataFrame,
                      column_beam_joint_xlsx: str,
                      output_floor: list,
                      output_serial: list):
    user_define = False

    beam_list = save_temp_file.read_temp(beam_pkl)
    column_list = save_temp_file.read_temp(column_pkl)
    # column_list = [column for column in column_list if column.serial == 'C68']
    floor_list = column_floor_parameter(column_list=column_list,
                                        floor_parameter_xlsx=column_beam_joint_xlsx)  # sort column floor
    # sort up bottom column
    column_list = [
        column for column in column_list if column.floor in ['1F', '2F', '3F', '4F']]

    sort_floor_column(floor_list=floor_list, column_list=column_list)
    beam_floor_parameter(beam_list=beam_list,
                         floor_parameter_xlsx=column_beam_joint_xlsx)  # sort beam floor

    if user_define:
        column_beam_df = read_parameter_df(
            read_file=column_beam_joint_xlsx, sheet_name="梁柱接頭表")

    excel_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Joint_Result.xlsx'
    )
    docx_filename = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Joint_Result.docx'
    )

    joint_df, beams_df = check_column_joint(excel_filename=excel_filename,
                                            docx_filename=docx_filename,
                                            column_beam_df=column_beam_df,
                                            column_list=column_list,
                                            beam_list=beam_list,
                                            floor_list=floor_list,
                                            output_floor=output_floor,
                                            output_serial=output_serial)

    return joint_df, beams_df, column_list, excel_filename, docx_filename


def read_column_beam_plan(plan_filename, layer_config: dict):
    '''
    match column amd beam relation in plan
    '''
    read_column_beam_plan = main_logger.getChild('draw_rebar_data')

    error_count = 0
    pythoncom.CoInitialize()
    # Step 1. 打開應用程式
    wincad_plan = None
    while not wincad_plan and error_count <= 10:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")

        except Exception as e:
            read_column_beam_plan.info(e)
            error_count += 1
            time.sleep(5)

    # Step 2. 匯入檔案
    doc_plan = None
    while not doc_plan and error_count <= 10:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)

        except Exception as e:
            read_column_beam_plan.info(e)
            error_count += 1
            time.sleep(5)

    # Step 3. 匯入modelspace
    total = 0
    msp_plan = None
    while not msp_plan and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
            total = msp_plan.Count
        except Exception as e:
            read_column_beam_plan.info(e)
            error_count += 1
            time.sleep(5)

    used_layer_list = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name

    block_layer = layer_config['block_layer']
    beam_name_text_layer = layer_config['beam_name_text_layer']
    beam_mline_layer = layer_config['beam_mline_layer']
    column_name_text_layer = layer_config['column_name_text_layer']
    column_block_layer = layer_config['column_block_layer']
    floor_text_layer = layer_config['floor_text_layer']

    block_object_type = ['AcDbBlockReference', "AcDbPolyline"]
    text_object_type = ['AcDbText', 'AcDbMText']
    floor_object_type = ['AcDbText']
    mline_object_type = ['AcDbMline', 'AcDbLine']
    column_block_object_type = ['AcDbBlockReference', "AcDbPolyline"]

    block_entity = []
    beam_name_text_entity = []
    column_name_text_entity = []
    floor_text_entity = []
    mline_entity = []
    column_block_entity = []

    count = 0

    for msp_object in msp_plan:
        object_list = []
        count += 1
        if count % 1000 == 0:
            read_column_beam_plan.info(f'平面圖已讀取{count}/{total}個物件')
        while error_count <= 3 and not object_list:
            try:
                # print(msp_object.Layer)
                if msp_object.Layer not in used_layer_list:
                    break

                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference":
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer

                if object_layer in used_layer_list:
                    if object_layer in block_layer and object.EntityName in block_object_type:
                        block_entity.append(
                            (object.GetBoundingBox()[0], object.GetBoundingBox()[1]))
                    if object_layer in beam_name_text_layer and object.EntityName in text_object_type:
                        start = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        end = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        mid = ((start[0] + end[0]) / 2,
                               (start[1] + end[1]) / 2)
                        beam_name_text_entity.append(
                            (mid, object.TextString, object.rotation))
                    if object_layer in column_name_text_layer and object.EntityName in text_object_type:
                        column_name_text_entity.append(
                            (object.GetBoundingBox()[0], object.TextString))
                    if object_layer in floor_text_layer and object.EntityName in floor_object_type:
                        floor_text_entity.append(
                            (object.GetBoundingBox()[0], object.TextString))
                    if object_layer in beam_mline_layer and object.EntityName in mline_object_type:
                        start = (round(object.GetBoundingBox()[0][0], 2), round(
                            object.GetBoundingBox()[0][1], 2))
                        end = (round(object.GetBoundingBox()[1][0], 2), round(
                            object.GetBoundingBox()[1][1], 2))
                        x_diff = abs(start[0] - end[0])
                        y_diff = abs(start[1] - end[1])
                        mid = ((start[0] + end[0]) / 2,
                               (start[1] + end[1]) / 2)
                        xy_direction = ""
                        mline_scale = round(
                            abs(object.MLineScale)) if 'MLineScale' in dir(object) else 0
                        if x_diff + y_diff > 100:  # 防過短複線
                            if x_diff < y_diff:  # 直的
                                xy_direction = "y"
                            else:  # 橫的
                                xy_direction = "x"
                            # mline_object = MlineObject(
                            #     start, mid, end, mline_scale, xy_direction)
                            mline_entity.append(
                                (start, mid, end, mline_scale, xy_direction))
                    if object_layer in column_block_layer and object.EntityName in column_block_object_type:
                        # column_block = ColumnBlock(object.GetBoundingBox()[
                        #                            0], object.GetBoundingBox()[1])
                        column_block_entity.append((object.GetBoundingBox()[0],
                                                   object.GetBoundingBox()[1]))
            except Exception as ex:
                error_count += 1
                object_list.append(object)
                time.sleep(5)

    # Finish Reading doc, close it
    try:
        doc_plan.Close(SaveChanges=False)
    except Exception as ex:
        pass
    return {
        'block_entity': block_entity,
        'beam_name_text_entity': beam_name_text_entity,
        'column_name_text_entity': column_name_text_entity,
        'floor_text_entity': floor_text_entity,
        'mline_entity': mline_entity,
        'column_block_entity': column_block_entity
    }


def sort_entity_to_floor(data: dict):
    block_entity = data['block_entity']
    floor_text_entity = data['floor_text_entity']
    beam_name_text_entity = data['beam_name_text_entity']
    column_name_text_entity = data['column_name_text_entity']
    column_block_tuple_entity = data['column_block_entity']
    mline_tuple_entity: list[tuple] = data['mline_entity']
    mline_entity = []
    column_block_entity = []
    for entity in mline_tuple_entity:
        mline_entity.append(MlineObject(*entity))
    for entity in column_block_tuple_entity:
        column_block_entity.append(ColumnBlock(*entity))

    sort_floor_block: dict[str, dict] = {}
    text = ""
    for block in block_entity:
        floor_text = [text for coor,
                      text in floor_text_entity if in_block(coor, block)]
        if len(floor_text) > 1:
            for text in floor_text:
                if check_is_floor(text):
                    break
        elif floor_text:
            text = floor_text[0]
        else:
            continue
        if not check_is_floor(text):
            continue

        searchObj = re.search(r'\(B?P?R?\d*F?\W?P?R?\d*F?\)', text)
        text = searchObj.group()

        if text not in sort_floor_block:
            sort_floor_block.update({text: {}})
        else:
            continue

        beam_name_text_list = [
            entity for entity in beam_name_text_entity if in_block(entity[0], block)]
        column_name_text_list = [
            entity for entity in column_name_text_entity if in_block(entity[0], block)]
        column_block_list = [
            entity for entity in column_block_entity if in_block(entity.start, block)]
        mline_list = [
            entity for entity in mline_entity if in_block(entity.mid, block)]

        sort_floor_block[text].update({
            'block': [block, text],
            'beam_name_text_list': beam_name_text_list,
            'column_name_text_list': column_name_text_list,
            'column_block_list': column_block_list,
            'mline_list': mline_list
        })
    new_sort_floor = sort_floor_text(sort_floor_block)
    return new_sort_floor


def get_distance(coor1, coor2, direction: str = ''):
    from math import sqrt
    x_ratio = 1
    y_ratio = 1
    if direction == 'x':
        y_ratio = 1.5
    if direction == 'y':
        x_ratio = 1.5
    if isinstance(coor1, tuple) and isinstance(coor2, tuple):
        try:
            return sqrt(x_ratio * (coor1[0][0]-coor2[0][0]) ** 2 + y_ratio * (coor1[0][1]-coor2[0][1]) ** 2)
            return abs(coor1[0][0]-coor2[0][0]) + abs(coor1[0][1]-coor2[0][1])
        except TypeError:
            return sqrt(x_ratio * (coor1[0]-coor2[0]) ** 2 + y_ratio * (coor1[1]-coor2[1]) ** 2)
            return abs(coor1[0]-coor2[0]) + abs(coor1[1]-coor2[1])
    return 10000


def match_beam_mline(data: dict):
    string_pattern = r"[WB|B|G][\d|\w]'*-*\d*"
    line_rotation = ''
    for floor, items in data.items():
        mline_list: list[MlineObject] = items['mline_list']
        beam_name_text_list = items['beam_name_text_list']
        beam_name_text_list = [entity for entity in beam_name_text_list if re.search(
            string_pattern, entity[1])]
        closet_mline = None
        for coor, beam_name, rotation in beam_name_text_list:
            beam_name: str
            beam_name = beam_name.strip().replace('\P', '')

            if rotation == 0:
                line_rotation = 'x'
            else:
                line_rotation = 'y'

            closet_mline = min(
                [mline for mline in mline_list if mline.xy_direction == line_rotation], key=lambda mline: get_distance(coor, mline.mid, line_rotation))

            if closet_mline.beam_serial is not None:
                if get_distance(closet_mline.beam_serial[0], closet_mline.mid) > get_distance(coor, closet_mline.mid) or closet_mline.beam_serial[1] == 'WB1':
                    closet_mline.beam_serial = (coor, beam_name, rotation)
                else:
                    # 避免距離受到編號位置影響，如最接近複線已有編號且該編號之距離較短，則找尋找目前無編號之複線，同時透過最近複線之距離避免抓取到過遠之複線
                    none_mline_list = [
                        mline for mline in mline_list if mline.xy_direction == line_rotation and mline.beam_serial is None]

                    if not none_mline_list:
                        continue

                    closet_none_mline: MlineObject = min(
                        none_mline_list, key=lambda mline: get_distance(coor, mline.mid, line_rotation))
                    print(
                        f'{beam_name}:{5 * get_distance(closet_mline.beam_serial[0], closet_mline.mid)} , {get_distance(coor, closet_none_mline.mid)}')
                    if 5 * get_distance(closet_mline.beam_serial[0], closet_mline.mid) > get_distance(coor, closet_none_mline.mid):
                        closet_none_mline.beam_serial = (
                            coor, beam_name, rotation)
            else:
                closet_mline.beam_serial = (coor, beam_name, rotation)

        # for mline in mline_list:
        #     if mline.xy_direction == "x":
        #         mline.beam_serial = min([(coor, beam_name) for coor, beam_name, rotation in beam_name_text_list if rotation == 0],
        #                                 key=lambda beam: get_distance(beam[0], mline.mid))
        #     else:
        #         mline.beam_serial = min([(coor, beam_name) for coor, beam_name, rotation in beam_name_text_list if rotation != 0],
        #                                 key=lambda beam: get_distance(beam[0], mline.mid))


def match_column_block(data: dict):
    string_pattern = r"C\d"
    for floor, items in data.items():
        column_block_list: list[ColumnBlock] = items['column_block_list']
        column_name_text_list = items['column_name_text_list']
        column_name_text_list = [
            entity for entity in column_name_text_list if re.search(string_pattern, entity[1])]
        if not column_block_list:
            continue
        start = time.time()
        match_result, _ = match_points(points1=[coor for coor, column_name_text in column_name_text_list],
                                       points2=[block.mid for block in column_block_list])
        total_distance = 0
        for i, j in match_result:
            coor, column_name_text = column_name_text_list[i]
            closest_block = column_block_list[j]
            closest_block.column_serial = (coor, column_name_text)
            total_distance += get_distance(coor, closest_block.mid)
        print(
            f"Cost time {time.time() - start}s , min distance:{_} , total distance:{total_distance}")
        # for coor, column_name_text in column_name_text_list:
        #     column_name_text: str
        #     column_name_text = column_name_text.strip()
        #     try:
        #         assert column_name_text != "C124"
        #     except:
        #         pass
        #     closest_block = min(column_block_list, key=lambda column_block: get_distance(
        #         coor, column_block.start) + get_distance(coor, column_block.end))

        #     if closest_block.column_serial is not None:
        #         if get_distance(closest_block.column_serial[0], closest_block.start) + get_distance(closest_block.column_serial[0], closest_block.end) > \
        #                 get_distance(coor, closest_block.start) + get_distance(coor, closest_block.end):
        #             origin_block = closest_block.column_serial
        #             closest_block.column_serial = (coor, column_name_text)

        #     else:
        #         closest_block.column_serial = (coor, column_name_text)
        # for column_block in column_block_list:
        #     column_block.column_serial = min(column_name_text_list, key=lambda entity: get_distance(
        #         entity[0], column_block.start) + get_distance(entity[0], column_block.end))


def match_beam_column(data: dict):

    for floor, items in data.items():
        column_block_list: list[ColumnBlock] = items['column_block_list']
        if not column_block_list:
            continue
        mline_list: list[MlineObject] = items['mline_list']
        for mline in mline_list:
            mline.left_column = min(column_block_list, key=lambda columnBlock: get_distance(
                mline.start, columnBlock.start) + get_distance(mline.start, columnBlock.end))
            mline.right_column = min(column_block_list, key=lambda columnBlock: get_distance(
                mline.end, columnBlock.start) + get_distance(mline.end, columnBlock.end))


def output_match_result(data: dict):
    df = pd.DataFrame([], columns=['樓層', '梁編號', '左柱',
                      '左側偏心', '右柱', '右側偏心', '方向'])
    row = 0
    for floor, items in data.items():
        mline_list: list[MlineObject] = items['mline_list']
        mline_list = [mline for mline in mline_list if mline.beam_serial]
        for mline in mline_list:
            mline.check()
            df.loc[row] = [floor,
                           mline.beam_serial[1],
                           mline.left_column.column_serial[1] if mline.left_column else "",
                           round(mline.left_offset) if mline.left_column else "",
                           mline.right_column.column_serial[1] if mline.right_column else "",
                           round(mline.right_offset) if mline.right_column else "",
                           mline.xy_direction.upper()]
            row += 1
    # OutputExcel([df], "scipy_test.xlsx", "梁柱接頭表test")
    return df


def match_column_beam_plan(plan_filename, layer_config, pkl=""):

    if pkl == "":
        cad_result = read_column_beam_plan(plan_filename, layer_config)
        save_temp_file.save_pkl(
            data=cad_result, tmp_file=f'{os.path.splitext(plan_filename)[0]}_plan_set.pkl')
    else:
        cad_result = save_temp_file.read_temp(
            tmp_file=pkl)

    # seperate entity to floor block
    sort_floor_block = sort_entity_to_floor(cad_result)
    # match beam and mline
    match_beam_mline(sort_floor_block)
    # match column name and block
    match_column_block(sort_floor_block)
    # match column and beam
    match_beam_column(sort_floor_block)
    # output column beam match
    column_beam_df = output_match_result(sort_floor_block)

    main_logger.info('完成梁柱平面關係配對')
    return sort_floor_block, column_beam_df


def joint_scan_main(plan_filename,
                    layer_config,
                    output_folder,
                    project_name,
                    beam_pkl,
                    column_pkl,
                    column_beam_joint_xlsx,
                    client_id="temp",
                    pkl="",
                    output_floor: list = [],
                    output_serial: list = []):
    def concat_list(input_list):
        result = []
        for l in input_list:
            result += l
        return result
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id=client_id)
    sort_floor_block, column_beam_df = match_column_beam_plan(plan_filename=plan_filename,
                                                              layer_config=layer_config,
                                                              pkl=pkl)

    joint_df, beams_df, column_list, excel_filename, docx_filename = column_joint_main(output_folder=output_folder,
                                                                                       project_name=project_name,
                                                                                       beam_pkl=beam_pkl,
                                                                                       column_pkl=column_pkl,
                                                                                       column_beam_df=column_beam_df,
                                                                                       column_beam_joint_xlsx=column_beam_joint_xlsx,
                                                                                       output_floor=output_floor,
                                                                                       output_serial=output_serial)

    # output plan view
    match_rebar_data_with_object(data=sort_floor_block,
                                 beams_df=beams_df,
                                 column_list=column_list)
    new_plan_view = (
        f'{output_folder}/'
        f'{project_name}_'
        f'{time.strftime("%Y%m%d_%H%M%S", time.localtime())}_'
        f'Joint_Plan.dwg'
    )
    # new_plan_view = f'{output_folder}/{os.path.splitext(os.path.basename(plan_filename))[0]}_joint_plan.dwg'
    main_logger.info("正在輸出結果")
    create_joint_plan_view(plan_filename=new_plan_view,
                           mline_list=concat_list(
                               [items["mline_list"] for floor, items in sort_floor_block.items()]),
                           column_block_list=concat_list(
                               [items["column_block_list"] for floor, items in sort_floor_block.items()]),
                           block_list=[(floor, items["block"])
                                       for floor, items in sort_floor_block.items()],
                           layer_config={
                               "Block": {
                                   "ColorIndex": 2,
                                   "Linetype": "HIDDEN",
                                   "Lineweight": 0.5
                               },
                               "Warning": {
                                   "ColorIndex": 1,
                                   "Linetype": "HIDDEN",
                                   "Lineweight": 0.5
                               },
                               "Beam": {
                                   "ColorIndex": 2,
                                   "Linetype": "HIDDEN",
                                   "Lineweight": 0.5
                               },
                               "Column": {
                                   "ColorIndex": 4,
                                   "Linetype": "Continuous",
                                   "Lineweight": 0.5
                               },
                               "BeamText": {
                                   "ColorIndex": 2,
                                   "Linetype": "Continuous",
                                   "Lineweight": 0.5
                               },
                               "ColumnText": {
                                   "ColorIndex": 2,
                                   "Linetype": "Continuous",
                                   "Lineweight": 0.5
                               },
                               "RebarText": {
                                   "ColorIndex": 2,
                                   "Linetype": "Continuous",
                                   "Lineweight": 0.5
                               }
                           },
                           client_id=client_id)
    main_logger.info("EOF")
    return os.path.basename(new_plan_view), os.path.basename(excel_filename), os.path.basename(docx_filename)


def summary_column_joint_df(joint_df: pd.DataFrame, column_dcr: str):
    def merge_lists_of_dicts(series):
        result_list = []
        for lst in series:
            result_list.extend(lst)
        return result_list
    # Assume df is your DataFrame loaded with the appropriate data
    # Create 'X_DCR' and 'Y_DCR' columns before grouping
    joint_df = joint_df.copy()
    joint_df['X_DCR'] = joint_df.apply(
        lambda row: row[column_dcr] if row['pos'] == 'X' else None, axis=1)
    joint_df['Y_DCR'] = joint_df.apply(
        lambda row: row[column_dcr] if row['pos'] == 'Y' else None, axis=1)

    # Group by 'story' and 'column'
    joint_df = joint_df.groupby(['story', 'column']).agg({
        'beams_rebar': merge_lists_of_dicts,  # Use your previously defined function
        'X_DCR': 'max',  # Since only one value per group should be non-null, max works
        'Y_DCR': 'max',  # Same logic as for 'X_DCR'
    }).reset_index()

    joint_df['type'] = np.where(
        joint_df['beams_rebar'].apply(len) < 8, '外柱', '內柱')
    group_df = joint_df.groupby(['story', 'type'])
    # Define aggregation functions for counts
    agg_funcs = {
        'X_DCR': [
            ('DCR >= 1.5', lambda x: (x >= 1.5).sum()),
            ('1.5 > DCR >= 1.25', lambda x: ((x >= 1.25) & (x < 1.5)).sum()),
            ('1.25 > DCR >= 1', lambda x: ((x >= 1) & (x < 1.25)).sum()),
            ('DCR < 1', lambda x: (x < 1).sum()),
            ('Count', 'count')
        ],
        'Y_DCR': [
            ('DCR >= 1.5', lambda x: (x >= 1.5).sum()),
            ('1.5 > DCR >= 1.25', lambda x: ((x >= 1.25) & (x < 1.5)).sum()),
            ('1.25 > DCR >= 1', lambda x: ((x >= 1) & (x < 1.25)).sum()),
            ('DCR < 1', lambda x: (x < 1).sum()),
            ('Count', 'count')
        ]
    }

    # Perform the aggregation
    result = group_df.agg(agg_funcs)

    # Calculate the proportions
    result[('X_DCR (%)', 'DCR >= 1.5 (%)')] = (
        result[('X_DCR', 'DCR >= 1.5')] / result[('X_DCR', 'Count')] * 100).round(2)
    result[('X_DCR (%)', '1.5 > DCR >= 1.25 (%)')] = (
        result[('X_DCR', '1.5 > DCR >= 1.25')] / result[('X_DCR', 'Count')] * 100).round(2)
    result[('X_DCR (%)', '1.25 > DCR >= 1 (%)')] = (
        result[('X_DCR', '1.25 > DCR >= 1')] / result[('X_DCR', 'Count')] * 100).round(2)
    result[('X_DCR (%)', 'DCR < 1 (%)')] = (
        result[('X_DCR', 'DCR < 1')] / result[('X_DCR', 'Count')] * 100).round(2)

    result[('Y_DCR (%)', 'DCR >= 1.5 (%)')] = (
        result[('Y_DCR', 'DCR >= 1.5')] / result[('Y_DCR', 'Count')] * 100).round(2)
    result[('Y_DCR (%)', '1.5 > DCR >= 1.25 (%)')] = (
        result[('Y_DCR', '1.5 > DCR >= 1.25')] / result[('Y_DCR', 'Count')] * 100).round(2)
    result[('Y_DCR (%)', '1.25 > DCR >= 1 (%)')] = (
        result[('Y_DCR', '1.25 > DCR >= 1')] / result[('Y_DCR', 'Count')] * 100).round(2)
    result[('Y_DCR (%)', 'DCR < 1 (%)')] = (
        result[('Y_DCR', 'DCR < 1')] / result[('Y_DCR', 'Count')] * 100).round(2)

    return result


def sort_dataframe_by_floor(df: pd.DataFrame, floor_list: list[str], column1: str, serial_list: list[str] = [], column2: str = ""):
    # Create a mapping from story value to a custom sort key
    order_mapping = {key: i for i, key in enumerate(floor_list)}

    # Add a helper column based on the custom order
    df['sort_key1'] = df.index.get_level_values(column1).map(order_mapping)

    if column2 == "":
        # Sort the DataFrame by the helper column
        df_sorted = df.sort_values(['sort_key1'])

        # Optionally, remove the helper column if not needed
        df_sorted.drop(['sort_key1'], axis=1, inplace=True)

        return df_sorted

    # serial_list
    new_serial_list = [re.findall(r'(\D*)(\d*)', s)[0] for s in serial_list]
    new_serial_list.sort(key=lambda s: (
        s[0], float(s[1]) if s[1].isnumeric() else 0))
    serial_list = [s[0] + s[1] for s in new_serial_list]

    order_mapping = {key: i for i, key in enumerate(serial_list)}
    df['sort_key2'] = df.index.get_level_values(column2).map(order_mapping)

    # Sort the DataFrame by the helper column
    df_sorted = df.sort_values(['sort_key1', 'sort_key2'])

    # Optionally, remove the helper column if not needed
    df_sorted.drop(['sort_key1', 'sort_key2'], axis=1, inplace=True)

    return df_sorted


if __name__ == "__main__":
    output_folder = r"D:\Desktop\BeamQC\TEST\2024-0822"
    project_name = r"0822"
    beam_pkl = r"TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-temp-beam_list.pkl"
    column_pkl = r"TEST\2024-0822\column_list.pkl"
    column_beam_joint_xlsx = r"TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-floor.xlsx"

    plan_filename = r"D:\Desktop\BeamQC\TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-XS-PLAN.dwg"
    layer_config = {
        'block_layer': ['0', 'DwFm', 'DEFPOINTS'],
        'beam_name_text_layer': ['S-TEXTG'],
        'beam_mline_layer': ['S-RCBMG'],
        'column_name_text_layer': ['S-TEXTC'],
        'column_block_layer': ['S-COL'],
        'floor_text_layer': ['S-TITLE']
    }

    joint_scan_main(plan_filename=plan_filename,
                    layer_config=layer_config,
                    output_folder=output_folder,
                    project_name=project_name,
                    beam_pkl=beam_pkl,
                    column_pkl=column_pkl,
                    column_beam_joint_xlsx=column_beam_joint_xlsx,
                    pkl=r'TEST\2024-0822\P2022-04A 國安社宅二期暨三期22FB4-2024-08-22-10-00-XS-PLAN_plan_set.pkl',
                    output_floor=['2F', '3F'])
    # match_column_beam_plan(plan_filename=plan_filename,
    #                        layer_config=layer_config)

    # column_joint_main(output_folder=output_folder,
    #                   project_name=project_name,
    #                   beam_pkl=beam_pkl,
    #                   column_pkl=column_pkl,
    #                   column_beam_joint_xlsx=column_beam_joint_xlsx)
