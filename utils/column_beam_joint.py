import pandas as pd
import math
from typing import Literal
from item.rebar import _rebar


def calculate_column_beam_joint(shaped_column: object,
                                columns: pd.DataFrame,
                                beams: pd.DataFrame,
                                storys: list[tuple]):
    columns.set_index(['Story', 'Column'], inplace=True)
    beams.set_index(['Story', 'Column'], inplace=True)
    (S1, H1), (S2, H2) = find_story_height(shaped_column.story, storys)
    row = columns.loc[(shaped_column.story, shaped_column.column)]
    result = []
    story = shaped_column.story
    point1_beam, point2_beam = row['Beam']
    joint_beams = {
        '1': None,
        '2': None,
        '3': None,
        '4': None,
    }
    for beam in point1_beam:
        row = beams.loc[(story, beam)]
        if 45 < row['Angle'] < 135:
            joint_beams['4'] = row
        if 225 < row['Angle'] < 315:
            joint_beams['3'] = row
        if 45 > row['Angle'] or row['Angle'] > 315:
            joint_beams['2'] = row
        if 225 > row['Angle'] > 135:
            joint_beams['1'] = row
    for beam in point2_beam:
        row = beams.loc[(story, beam)]
        if 45 < row['Angle'] < 135:
            joint_beams['3'] = row
        if 225 < row['Angle'] < 315:
            joint_beams['4'] = row
        if 45 > row['Angle'] or row['Angle'] > 315:
            joint_beams['1'] = row
        if 225 > row['Angle'] > 135:
            joint_beams['2'] = row
    # y dir joint shaer

    Mpr1 = 0
    Mpr2 = 0
    Ts1 = Ts2 = Cc1 = Cc2 = 0
    for pos in ['x', 'y']:
        bj = []
        if pos == 'x':
            left_pos = '1'
            right_pos = '2'
            hj = shaped_column.max_x - shaped_column.min_x
        if pos == 'y':
            left_pos = '3'
            right_pos = '4'
            hj = shaped_column.max_y - shaped_column.min_y
        if not joint_beams[left_pos] is None:
            beam = joint_beams[left_pos]
            fy, fc, fys = beam['Material']
            d = beam['H'] * 100
            b = beam['B'] * 100
            x1 = min(hj/4, 0)
            x2 = min(hj/4, 0)
            bj.append(x1 + x2 + b)
            Ts1 = cal_rebar_As(beam['Right Rebar'][0:2]) * 1.25 * fy / 10
            Mpr1 += Ts1 * (d - 0.5*(Ts1/(0.85*fc*b)))
            Ts2 = cal_rebar_As(beam['Right Rebar'][2:4]) * 1.25 * fy / 10
            Mpr2 += Ts2 * (d - 0.5*(Ts2/(0.85*fc*b)))
        if not joint_beams[right_pos] is None:
            beam = joint_beams[right_pos]
            fy, fc, fys = beam['Material']
            d = beam['H'] * 100
            b = beam['B'] * 100
            x1 = min(hj/4, 0)
            x2 = min(hj/4, 0)
            bj.append(x1 + x2 + b)
            Cc2 = cal_rebar_As(beam['Left Rebar'][2:4]) * 1.25 * fy / 10
            Mpr1 += Cc2 * (d - 0.5*(Cc2/(0.85*fc*b)))
            Cc1 = cal_rebar_As(beam['Left Rebar'][0:2]) * 1.25 * fy / 10
            Mpr2 += Cc1 * (d - 0.5*(Cc1/(0.85*fc*b)))
        bij = sum(bj) / len(bj)
        Aj = bij * hj
        design_code = determine_design_code(top_floor=S2 is None,
                                            joint_beams=joint_beams,
                                            beams_df=beams_df,
                                            hj=hj,
                                            dir=pos)
        Vn = 0.85 * design_code * math.sqrt(shaped_column.fc) * Aj
        Vu = max(Ts1 + Cc2 - Mpr1 / ((H1 * 100 + H2*100)/2),
                 Ts2 + Cc1 - Mpr2 / ((H1 * 100 + H2*100)/2))
        result.append({
            'story': story,
            'column': shaped_column.column,
            'pos': pos,
            'design_code': design_code,
            'hj': hj,
            'bj': bij,
            'Aj': Aj,
            'Mpr1': Mpr1,
            'Mpr2': Mpr2,
            'Vu': Vu,
            'Vn': Vn,
            'DCR': round(Vu/Vn, 2)
        })
    return result


def determine_design_code(top_floor: bool,
                          joint_beams: pd.DataFrame,
                          beams_df: pd.DataFrame,
                          floor: str,
                          hj: float,
                          dir: Literal['x', 'y'],
                          offset_beams: pd.DataFrame):
    '''
    401-112 15.2.6-15.2.8
    '''
    _code_15_2_6 = False
    _code_15_2_7 = False
    _code_15_2_8 = False
    design_code = 0
    dir = dir.lower()
    if dir == 'x':
        if joint_beams['X_Left'] and joint_beams['X_Right']:
            _code_15_2_7 = offset_beams['X_Left'] == offset_beams['X_Right']
        if joint_beams['Y_Left'] and joint_beams['Y_Right']:
            if beams_df.loc[floor, joint_beams['Y_Left']]['beam'].width > 0.75 * hj and beams_df.loc[floor, joint_beams['Y_Right']]['beam'].width > 0.75 * hj:
                _code_15_2_8 = True
    if dir == 'y':
        if joint_beams['Y_Left'] and joint_beams['Y_Right']:
            _code_15_2_7 = offset_beams['Y_Left'] == offset_beams['Y_Right']
        if joint_beams['X_Left'] and joint_beams['X_Right']:
            if beams_df.loc[floor, joint_beams['X_Left']]['beam'].width > 0.75 * hj and beams_df.loc[floor, joint_beams['X_Right']]['beam'].width > 0.75 * hj:
                _code_15_2_8 = True
    if not top_floor:
        _code_15_2_6 = True
    design_code = get_design_code_value(
        _code_15_2_6, _code_15_2_7, _code_15_2_8)
    return design_code, ["V" if code else "X" for code in (_code_15_2_6, _code_15_2_7, _code_15_2_8)]


def get_design_code_value(_code_15_2_6, _code_15_2_7, _code_15_2_8):
    design_code = 0
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (True, True, True):
        design_code = 5.3
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (True, True, False):
        design_code = 3.9
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (True, False, True):
        design_code = 3.9
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (True, False, False):
        design_code = 3.2
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (False, True, True):
        design_code = 3.9
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (False, True, False):
        design_code = 3.2
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (False, False, True):
        design_code = 3.2
    if (_code_15_2_6, _code_15_2_7, _code_15_2_8) == (False, False, False):
        design_code = 2.1
    return design_code


def cal_rebar_As(rebars: list[str]):
    As = 0
    for rebar in rebars:
        if '-#' in str(rebar):
            num, size = rebar.split('-')
            As += float(num) * _rebar[size]
    return As


def find_story_height(item_to_find, list_of_item: list):
    index_of_item = [i for i, value in enumerate(
        list_of_item) if value[0] == item_to_find]

    if len(index_of_item) > 0 and index_of_item[0] < len(list_of_item) - 1:
        return list_of_item[index_of_item[0]], list_of_item[index_of_item[0] - 1]
    return list_of_item[index_of_item[0]], (None, 0)
