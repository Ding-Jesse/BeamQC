from __future__ import annotations
from typing import Literal
import time
import pythoncom
import win32com.client
import save_temp_file
import re
from plan_to_beam import turn_floor_to_float, turn_floor_to_string, turn_floor_to_list, floor_exist, vtFloat, error, progress
from collections import Counter

layer_config: dict[Literal['block_layer',
                           'name_text_layer', 'floor_text_layer'], str]


def read_plan_cad(plan_filename, progress_file, layer_config: dict[Literal['block_layer', 'name_text_layer', 'floor_text_layer'], str]):
    error_count = 0
    pythoncom.CoInitialize()
    progress('開始讀取平面圖', progress_file)
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_plan = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 1: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 1/15', progress_file)

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_plan = wincad_plan.Documents.Open(plan_filename)
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 2: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 2/15', progress_file)

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_plan = doc_plan.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(
                f'read_beam error in step 3: {e}, error_count = {error_count}.')
    progress('平面圖讀取進度 3/15', progress_file)
    used_layer_list = []
    count = 0
    total = msp_plan.Count

    for key, layer_name in layer_config.items():
        used_layer_list += layer_name

    block_layer = layer_config['block_layer']
    name_text_layer = layer_config['name_text_layer']
    floor_text_layer = layer_config['floor_text_layer']

    block_object_type = ['AcDbBlockReference', "AcDbPolyline"]
    text_object_type = ['AcDbText']
    floor_object_type = ['AcDbText']

    block_entity = []
    name_text_entity = []
    floor_text_entity = []
    for key, layer_name in layer_config.items():
        used_layer_list += layer_name
    for msp_object in msp_plan:
        object_list = []
        error_count = 0
        count += 1
        if count % 1000 == 0 or count == total:
            progress(f'平面圖已讀取{count}/{total}個物件', progress_file)
        while error_count <= 3 and not object_list:
            try:
                if msp_object.Layer not in used_layer_list:
                    break
                # print(f'{msp_object.Layer}:{msp_object.EntityName}')
                object_list = [msp_object]
                if msp_object.EntityName == "AcDbBlockReference" and msp_object.Layer not in block_layer:
                    if msp_object.GetAttributes():
                        object_list = list(msp_object.GetAttributes())
                    else:
                        object_list = list(msp_object.Explode())
            except Exception as ex:
                error_count += 1
                time.sleep(2)
                error(
                    f'read_plan error in step 7-1: {ex}, error_count = {error_count}.')
        while error_count <= 3 and object_list:
            object = object_list.pop()
            try:
                if object.Layer == '0':
                    object_layer = msp_object.Layer
                else:
                    object_layer = object.Layer

                if object_layer in used_layer_list:
                    print(object_layer)
                    if object_layer in block_layer and object.EntityName in block_object_type:
                        block_entity.append(
                            (object.GetBoundingBox()[0], object.GetBoundingBox()[1]))
                    if object_layer in name_text_layer and object.EntityName in text_object_type:
                        name_text_entity.append(
                            (object.GetBoundingBox()[0], object.TextString))
                    if object_layer in floor_text_layer and object.EntityName in floor_object_type:
                        floor_text_entity.append(
                            (object.GetBoundingBox()[0], object.TextString))

            except Exception as ex:
                error_count += 1
                object_list.append(object)
                time.sleep(5)
                error(
                    f'read_plan error in step 7: {ex}, error_count = {error_count}.')
    return {
        'block_entity': block_entity,
        'name_text_entity': name_text_entity,
        'floor_text_entity': floor_text_entity
    }


def in_block(coor: tuple, block: tuple[tuple, tuple]):
    return ((coor[0] - block[0][0]) * (coor[0] - block[1][0]) < 0) and ((coor[1] - block[0][1]) * (coor[1] - block[1][1]) < 0)


def check_is_floor(text: str):
    if re.match(r'\(B?P?R?\d*F?\W?P?R?\d*F?\)', text):
        return True
    return False


def sort_floor_text(data: dict[str, Counter]):
    '''
    make 6-7F to ['6F','7F']
    '''
    sort_dict: dict[str, Counter] = {}
    floor_float_list = [turn_floor_to_float(re.sub(r'\(|\)', '', floor)) for floor in data.keys(
    ) if not re.findall(r'\W', re.sub(r'\(|\)', '', floor))]
    # not equal to -1000 (FB)
    Bmax = min([f for f in floor_float_list if f < 0 and f != -1000])
    Fmax = max([f for f in floor_float_list if 1000 > f > 0])
    # not equal to PRF (FB)
    Rmax = max([f for f in floor_float_list if f > 1000 and f != 2000])
    for floor in data.keys():
        if re.findall(r'\W', floor):
            for new_floor in turn_floor_to_list(floor=re.sub(r'\(|\)', '', floor), Bmax=Bmax, Fmax=Fmax, Rmax=Rmax):
                sort_dict.update({new_floor: data[floor]})
        else:
            sort_dict.update({re.sub(r'\(|\)', '', floor): data[floor]})
    return sort_dict


def sort_name_text(data):
    block_entity = data['block_entity']
    name_text_entity = data['name_text_entity']
    floor_text_entity = data['floor_text_entity']
    sort_result = {}
    for block in block_entity:
        name_text_list = [text for coor,
                          text in name_text_entity if in_block(coor, block)]
        floor_text = [text for coor,
                      text in floor_text_entity if in_block(coor, block)]
        if len(floor_text) > 1:
            for text in floor_text:
                if check_is_floor(text):
                    break
        if text not in sort_result:
            sort_result.update({text: Counter(name_text_list)})
    return sort_result


if __name__ == "__main__":
    plan_filename = r'D:\Desktop\BeamQC\TEST\2023-1013\華泰電子_S1結構平面圖_1120829.dwg'
    progress_file = './result/tmp'
    layer_config = {
        'block_layer': ['AREA'],
        'name_text_layer': ['BTXT', 'CTXT', 'BTXT_S_'],
        'floor_text_layer': ['TEXT1']
    }
    # cad_result = read_plan_cad(plan_filename, progress_file, layer_config)
    # save_temp_file.save_pkl(
    #     data=cad_result, tmp_file=r'D:\Desktop\BeamQC\TEST\2023-1013\1013-cal.pkl')
    data = save_temp_file.read_temp(
        tmp_file=r'D:\Desktop\BeamQC\TEST\2023-1013\1013-cal.pkl')
    result = sort_name_text(data)
    result = sort_floor_text(data=result)
    print(result)
