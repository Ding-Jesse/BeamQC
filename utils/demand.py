import pandas as pd
import numpy as np
from typing import Literal
from item.beam import Beam
from item.beam import RebarType
from item.column import Column
from math import sqrt, pow
from utils.column_beam_joint import cal_rebar_As, determine_design_code, get_design_code_value


def calculate_beam_gravity_load(beam: Beam):
    '''
    return w , Vg , Mg
    '''
    rebar_list = beam.rebar_table[RebarType.Bottom.value][RebarType.Middle.value]
    AsFy = sum([r.As * r.fy for r in rebar_list])
    length = beam.length
    Mn = AsFy * (beam.depth - beam.protect_layer)
    Mu = 0.9 * Mn
    # Mu = 1/12wl2
    w = Mu * 12 / (length * length)
    Vg = w * length / 2
    return w, Vg, Mu


def calculate_beam_earthquake_load(w, beam: Beam, pos: Literal[RebarType.Top, RebarType.Bottom]):
    '''
    return left_Mu_eq, right_Mu_eq, Veq, Vp
    '''
    rebar_list = beam.rebar_table[pos.value][RebarType.Right.value]
    AsFy = sum([r.As * r.fy for r in rebar_list])
    # 1/24wl2
    length = beam.length
    Mn = AsFy * (beam.depth - beam.protect_layer)
    Mu = 0.9 * Mn
    Mu_dl = 1 / 24 * w * length * length
    right_Mu_eq = Mu - Mu_dl
    right_Mpr = 1.25 * Mn

    rebar_list = beam.rebar_table[pos.value][RebarType.Left.value]
    AsFy = sum([r.As * r.fy for r in rebar_list])
    # 1/24wl2
    Mn = AsFy * (beam.depth - beam.protect_layer)
    Mu = 0.9 * Mn
    Mu_dl = 1 / 24 * w * length * length
    left_Mu_eq = Mu - Mu_dl
    left_Mpr = 1.25 * Mn
    Veq = (left_Mu_eq + right_Mu_eq) / length
    Vp = (right_Mpr + left_Mpr) / length

    return left_Mu_eq, right_Mu_eq, Veq, Vp


def calculate_column_axial_force(column: Column, ratio: float):
    return ratio * column.x_size * column.y_size * column.fc


def calculate_column_earthquake_moment(column: Column, theta: float):
    '''
    M = EI/L * theta
    '''
    E = 12000 * sqrt(column.fc)
    I = column.x_size * pow(column.y_size, 3) / 12
    L = column.height

    return E * I / L * theta


def combine_column_beams(column_beam_df: pd.DataFrame):
    df = pd.concat([
        column_beam_df[['樓層', '梁編號', '左柱', '左側偏心', '方向']].rename(
            columns={'樓層': "Floor", '左柱': 'Column', '梁編號': 'Beam', '左側偏心': 'Offset'}).assign(Side='Right'),
        column_beam_df[['樓層', '梁編號', '右柱', '右側偏心', '方向']].rename(
            columns={'樓層': "Floor", '右柱': 'Column', '梁編號': 'Beam', '右側偏心': 'Offset'}).assign(Side='Left')
    ])
    df['Column_Name'] = df['方向'] + '_' + df['Side']

    # Creating the pivot table
    pivot_df = df.pivot_table(
        index=['Floor', 'Column'], columns='Column_Name', values=['Beam', 'Offset'], aggfunc='first').reset_index()
    pivot_df = pivot_df.replace(np.nan, None)
    pivot_df.set_index(['Floor', 'Column'], inplace=True)
    # Reindexing to remove the name of the index column for aesthetics
    pivot_df.columns.names = [None, None]

    return pivot_df


def calculate_column_beam_joint_shear(column_list: list[Column], beam_list: list[Beam], column_beam_df: pd.DataFrame):
    for beam in beam_list:
        try:
            beam.sort_rebar_table()
        except AttributeError as ex:
            print(f'{beam.floor} {beam.serial} : {ex}')

    beams_df = pd.DataFrame(
        [{'floor': beam.floor, 'serial': beam.serial, 'beam': beam} for beam in beam_list])
    beams_df.set_index(['floor', 'serial'], inplace=True)

    # remove no rebar data beam
    column_beam_df.set_index(['樓層', '梁編號'], inplace=True)
    no_rebar_data = column_beam_df[~column_beam_df.index.isin(
        beams_df.index)].copy()
    no_rebar_data.loc[:, '錯誤'] = "無鋼筋資料"
    column_beam_df = column_beam_df[column_beam_df.index.isin(
        beams_df.index)]
    column_beam_df.reset_index(inplace=True)

    # transform to column center
    column_beam_df = combine_column_beams(column_beam_df)
    beams_df = beams_df[~beams_df.index.duplicated(keep='first')]
    result = []

    for column in column_list:
        column.joint_result = {}
        try:
            connect_beams = column_beam_df.loc[column.floor, column.serial]
        except KeyError:
            print(f"{column.floor + column.serial}:no beam data")
            continue
        H1 = column.floor_object.height
        column.connect_beams = connect_beams
        e1 = float("inf")
        e2 = float("inf")
        if column.up_column:
            H2 = column.up_column.floor_object.height
        else:
            H2 = 0
        for pos, hj, column_width in [('X', column.x_size, column.y_size), ('Y', column.y_size, column.x_size)]:
            bj = []
            beams_rebar = []
            Mpr1_top = Mpr1_bot = Mpr2_top = Mpr2_bot = 0
            Ts1_top = Ts2_top = Ts1_bot = Ts2_bot = 0
            As1_top = As2_top = As1_bot = As2_bot = 0
            x11 = x12 = x21 = x22 = 0
            bj1 = bj2 = 0
            b1 = b2 = d1 = d2 = 0
            left_beam_serial = right_beam_serial = "X"

            if not connect_beams['Beam'][f'{pos}_Left'] is None:
                beam_serial = connect_beams['Beam'][f'{pos}_Left']
                beam: Beam = beams_df.loc[(column.floor, beam_serial), 'beam']
                left_beam_serial = beam_serial
                fc = beam.fc
                d1 = beam.depth
                b1 = beam.width
                e1 = -1 * \
                    connect_beams['Offset'][f'{pos}_Left'] if pos == "Y" else connect_beams['Offset'][f'{pos}_Left']
                x11 = (column_width - b1) / 2 - \
                    connect_beams['Offset'][f'{pos}_Left']
                x12 = (column_width - b1) / 2 + \
                    connect_beams['Offset'][f'{pos}_Left']

                beams_rebar.append(
                    {f"左梁{beam.serial}上層鋼筋": beam.rebar_table[RebarType.Top.value][RebarType.Right.value]})

                As1_top = beam.get_rebar_table(
                    rebar_type1=RebarType.Top, rebar_type2=RebarType.Right)
                Ts1_top = As1_top * 1.25 * \
                    beam.rebar_table[RebarType.Top.value][RebarType.Right.value][0].fy if As1_top != 0 else beam.fy

                As1_bot = beam.get_rebar_table(
                    rebar_type1=RebarType.Bottom, rebar_type2=RebarType.Right)
                Ts1_bot = As1_bot * 1.25 * \
                    beam.rebar_table[RebarType.Bottom.value][RebarType.Right.value][0].fy if As1_bot != 0 else beam.fy

                beams_rebar.append(
                    {f"左梁{beam.serial}下層鋼筋": beam.rebar_table[RebarType.Bottom.value][RebarType.Right.value]})

                Mpr1_top += Ts1_top * (d1 - 0.5*(Ts1_top/(0.85*fc*b1)))
                Mpr1_bot += Ts1_bot * (d1 - 0.5*(Ts1_bot/(0.85*fc*b1)))
                x11 = min(hj / 4, x11)
                x12 = min(hj / 4, x12)
                bj1 = min(hj / 4, x11) + min(hj / 4, x12) + b1
                bj.append(bj1)

            if not connect_beams['Beam'][f'{pos}_Right'] is None:
                beam_serial = connect_beams['Beam'][f'{pos}_Right']
                beam: Beam = beams_df.loc[(column.floor, beam_serial), 'beam']
                right_beam_serial = beam_serial

                fc = beam.fc
                d2 = beam.depth
                b2 = beam.width
                e2 = -1 * \
                    connect_beams['Offset'][f'{pos}_Right'] if pos == "Y" else connect_beams['Offset'][f'{pos}_Right']
                x21 = (column_width - b2) / 2 - \
                    connect_beams['Offset'][f'{pos}_Right']
                x22 = (column_width - b2) / 2 + \
                    connect_beams['Offset'][f'{pos}_Right']

                beams_rebar.append(
                    {f"右梁{beam.serial}下層鋼筋": beam.rebar_table[RebarType.Bottom.value][RebarType.Left.value]})
                As2_bot = beam.get_rebar_table(
                    rebar_type1=RebarType.Bottom, rebar_type2=RebarType.Left)
                Ts2_bot = As2_bot * \
                    beam.rebar_table[RebarType.Bottom.value][RebarType.Left.value][0].fy if As2_bot != 0 else beam.fy

                Mpr2_bot = Ts2_bot * (d2 - 0.5*(Ts2_bot/(0.85*fc*b2)))

                beams_rebar.append(
                    {f"右梁{beam.serial}上層鋼筋": beam.rebar_table[RebarType.Top.value][RebarType.Left.value]})
                As2_top = beam.get_rebar_table(
                    rebar_type1=RebarType.Top, rebar_type2=RebarType.Left)
                Ts2_top = As2_top * \
                    beam.rebar_table[RebarType.Top.value][RebarType.Left.value][0].fy if As2_top != 0 else beam.fy

                Mpr2_top = Ts2_top * (d2 - 0.5*(Ts2_top/(0.85*fc*b2)))
                x21 = min(hj / 4, x21)
                x22 = min(hj / 4, x22)
                bj2 = min(hj / 4, x21) + min(hj / 4, x22) + b2
                bj.append(bj2)

            if not bj:
                print(f"{column.floor + column.serial}:no {pos} beam data")
                continue

            bij = sum(bj) / len(bj)
            bij = min(bij, column_width)
            Aj = bij * hj
            design_code, detail_design_code = determine_design_code(top_floor=column.up_column is None,
                                                                    joint_beams=connect_beams['Beam'],
                                                                    beams_df=beams_df,
                                                                    floor=column.floor,
                                                                    hj=hj,
                                                                    dir=pos,
                                                                    offset_beams=connect_beams['Offset'])
            # Conside if confine is fine with all beam
            consider_confine_code = get_design_code_value(detail_design_code[0] == "V",
                                                          detail_design_code[0] == "V",
                                                          True)

            Vn = 0.85 * design_code * sqrt(column.fc) * Aj
            Vh1 = (Mpr1_bot + Mpr2_top) / ((H1 + H2)/2)
            Vh2 = (Mpr2_bot + Mpr1_top) / ((H1 + H2)/2)
            Vu = max(Ts1_top + Ts2_bot - Vh1,
                     Ts1_bot + Ts2_top - Vh2)
            new_Vn = 0.85 * consider_confine_code * sqrt(column.fc) * Aj

            # Lee's methods
            lee_result = {}
            if not connect_beams['Beam'][f'{pos}_Right'] is None and not connect_beams['Beam'][f'{pos}_Left'] is None:

                inner_design_code = get_design_code_value(detail_design_code[0] == "V",
                                                          True, detail_design_code[2] == "V")

                outer_design_code = get_design_code_value(detail_design_code[0] == "V",
                                                          False, detail_design_code[2] == "V")
                inner_bj = (b1 + b2) / 2 - \
                    abs(connect_beams['Offset'][f'{pos}_Right'] -
                        connect_beams['Offset'][f'{pos}_Left'])

                outer_x1 = min((-b1 / 2 + e1 - x11), (-b2 / 2 + e2 - x21))
                outer_x2 = max((b1 / 2 + e1 + x12), (b2 / 2 + e2 + x22))

                outer_bj = max(outer_x2 - outer_x1 - inner_bj, 0)
                inner_Vn = inner_design_code * sqrt(column.fc) * inner_bj * hj
                outer_Vn = outer_design_code * sqrt(column.fc) * outer_bj * hj
                lee_Vn = 0.85 * (inner_design_code * sqrt(column.fc) * inner_bj *
                                 hj + outer_design_code * sqrt(column.fc) * outer_bj * hj)
                lee_result = {
                    'inner_design_code': inner_design_code,
                    'outer_design_code': outer_design_code,
                    'outer_x1': outer_x1,
                    'outer_x2': outer_x2,
                    'inner_bj': inner_bj,
                    'outer_bj': outer_bj,
                    'inner_Vn': inner_Vn / 1000,
                    'outer_Vn': outer_Vn / 1000,
                    'lee_Vn': lee_Vn / 1000,
                    'lee_DCR': round(Vu/lee_Vn, 2)
                }

            cal_result = {
                'story': column.floor,
                'column': column.serial,
                'left_beam': left_beam_serial,
                'right_beam': right_beam_serial,
                'beams_rebar': beams_rebar,
                'pos': pos,
                'design_code': design_code,
                '_code_15_2_6': detail_design_code[0],
                '_code_15_2_7': detail_design_code[1],
                '_code_15_2_8': detail_design_code[2],
                'column_width': column_width,
                'x11': x11,
                'x12': x12,
                'x21': x21,
                'x22': x22,
                'bj1': bj1,
                'bj2': bj2,
                'As1_top': As1_top,
                'As1_bot': As1_bot,
                'As2_top': As2_top,
                'As2_bot': As2_bot,
                'Ts1_top': Ts1_top / 1000,
                'Ts1_bot': Ts1_bot / 1000,
                'Ts2_top': Ts2_top / 1000,
                'Ts2_bot': Ts2_bot / 1000,
                'H1': H1,
                'H2': H2,
                'hc': hj,
                'bj': bij,
                'Aj': Aj,
                'Mpr1+(tf-m)': Mpr1_bot / 1000 / 100,
                'Mpr1-(tf-m)': Mpr1_top / 1000 / 100,
                'Mpr2+(tf-m)': Mpr2_bot / 1000 / 100,
                'Mpr2-(tf-m)': Mpr2_top / 1000 / 100,
                'Vh1(tf)': Vh1 / 1000,
                'Vh2(tf)': Vh2 / 1000,
                'Vu(tf)': Vu / 1000,
                'Vn(tf)': Vn / 1000,
                'DCR': round(Vu/Vn, 2),
                'Fine DCR': round(Vu/new_Vn, 2),
                'lee_DCR': round(Vu/Vn, 2)
            }
            if lee_result:
                cal_result.update(lee_result)
            result.append(cal_result)
            column.joint_result.update({pos: cal_result})

    return result, no_rebar_data, column_beam_df, beams_df
