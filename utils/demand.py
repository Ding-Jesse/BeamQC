import pandas as pd
import numpy as np
from typing import Literal
from item.beam import Beam
from item.beam import RebarType
from item.column import Column
from math import sqrt, pow
from utils.column_beam_joint import cal_rebar_As, determine_design_code


def calculate_beam_gravity_load(beam: Beam):
    '''
    return w , Vg
    '''
    rebar_list = beam.rebar_table[RebarType.Bottom.value][RebarType.Middle.value]
    AsFy = sum([r.As * r.fy for r in rebar_list])
    length = beam.length
    Mn = AsFy * (beam.depth - beam.protect_layer)
    Mu = 0.9 * Mn
    # Mu = 1/12wl2
    w = Mu * 12 / (length * length)
    Vg = w * length / 2
    return w, Vg


def calculate_beam_earthquake_load(w, beam: Beam, pos: Literal[RebarType.Top, RebarType.Bottom]):
    '''
    return left_Mu_eq, right_Mu_eq, Veq, Vp
    '''
    rebar_list = beam.rebar_table[pos][RebarType.Right.value]
    AsFy = sum([r.As * r.fy for r in rebar_list])
    # 1/24wl2
    length = beam.length
    Mn = AsFy * (beam.depth - beam.protect_layer)
    Mu = 0.9 * Mn
    Mu_dl = 1 / 24 * w * length * length
    right_Mu_eq = Mu - Mu_dl
    right_Mpr = 1.25 * Mn

    rebar_list = beam.rebar_table[pos][RebarType.Left.value]
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

        if column.up_column:
            H2 = column.up_column.floor_object.height
        else:
            H2 = 0
        for pos, hj, column_width in [('X', column.x_size, column.y_size), ('Y', column.y_size, column.x_size)]:
            bj = []
            beams_rebar = []
            Mpr1 = 0
            Mpr2 = 0
            Ts1 = Ts2 = Cc1 = Cc2 = 0

            if not connect_beams['Beam'][f'{pos}_Left'] is None:
                beam_serial = connect_beams['Beam'][f'{pos}_Left']
                beam: Beam = beams_df.loc[(column.floor, beam_serial), 'beam']
                fy, fc = beam.fy, beam.fc
                d = beam.depth
                b = beam.width

                x1 = min(hj/4, (column_width - b) / 2 -
                         connect_beams['Offset'][f'{pos}_Left'])
                x2 = min(hj/4, (column_width - b) / 2 +
                         connect_beams['Offset'][f'{pos}_Left'])
                bj.append(x1 + x2 + b)
                beams_rebar.append(
                    {f"左梁{beam.serial}上層鋼筋": beam.rebar_table[RebarType.Top.value][RebarType.Right.value]})
                Ts1 = beam.get_rebar_table(
                    rebar_type1=RebarType.Top, rebar_type2=RebarType.Right) * 1.25 * fy
                Cc1 = Ts1
                Mpr1 += Ts1 * (d - 0.5*(Ts1/(0.85*fc*b)))
                beams_rebar.append(
                    {f"左梁{beam.serial}下層鋼筋": beam.rebar_table[RebarType.Bottom.value][RebarType.Right.value]})
                Ts2 = beam.get_rebar_table(
                    rebar_type1=RebarType.Bottom, rebar_type2=RebarType.Right) * 1.25 * fy
                Cc2 = Ts2
                Mpr2 += Ts2 * (d - 0.5*(Ts2/(0.85*fc*b)))

            Ts1 = Ts2 = 0
            if not connect_beams['Beam'][f'{pos}_Right'] is None:
                beam_serial = connect_beams['Beam'][f'{pos}_Right']
                beam: Beam = beams_df.loc[(column.floor, beam_serial), 'beam']

                fy, fc = beam.fy, beam.fc
                d = beam.depth
                b = beam.width
                x1 = min(hj/4, (column_width - b) / 2 -
                         connect_beams['Offset'][f'{pos}_Right'])
                x2 = min(hj/4, (column_width - b) / 2 +
                         connect_beams['Offset'][f'{pos}_Right'])
                bj.append(x1 + x2 + b)
                beams_rebar.append(
                    {f"右梁{beam.serial}下層鋼筋": beam.rebar_table[RebarType.Bottom.value][RebarType.Left.value]})
                Ts1 = beam.get_rebar_table(
                    rebar_type1=RebarType.Bottom, rebar_type2=RebarType.Left) * 1.25 * fy
                Mpr1 += Ts1 * (d - 0.5*(Ts1/(0.85*fc*b)))
                beams_rebar.append(
                    {f"右梁{beam.serial}上層鋼筋": beam.rebar_table[RebarType.Top.value][RebarType.Left.value]})
                Ts2 = beam.get_rebar_table(
                    rebar_type1=RebarType.Top, rebar_type2=RebarType.Left) * 1.25 * fy
                Mpr2 += Ts2 * (d - 0.5*(Ts2/(0.85*fc*b)))
            if not bj:
                print(f"{column.floor + column.serial}:no {pos} beam data")
                continue
            bij = sum(bj) / len(bj)
            Aj = bij * hj
            design_code, detail_design_code = determine_design_code(top_floor=column.up_column is None,
                                                                    joint_beams=connect_beams['Beam'],
                                                                    beams_df=beams_df,
                                                                    floor=column.floor,
                                                                    hj=hj,
                                                                    dir=pos)
            Vn = 0.85 * design_code * sqrt(column.fc) * Aj
            Vu = max(Ts1 + Cc2 - Mpr1 / ((H1 + H2)/2),
                     Ts2 + Cc1 - Mpr2 / ((H1 + H2)/2))
            cal_result = {
                'story': column.floor,
                'column': column.serial,
                'beams_rebar': beams_rebar,
                'pos': pos,
                'design_code': design_code,
                '_code_15_2_6': detail_design_code[0],
                '_code_15_2_7': detail_design_code[1],
                '_code_15_2_8': detail_design_code[2],
                'column_width': column_width,
                'Ts1': Ts1,
                'Cc1': Cc1,
                'Ts2': Ts2,
                'Cc2': Cc2,
                'H1': H1,
                'H2': H2,
                'hj': hj,
                'bj': bij,
                'Aj': Aj,
                'Mpr1(tf-m)': Mpr1 / 1000 / 100,
                'Mpr2(tf-m)': Mpr2 / 1000 / 100,
                'Vu(tf)': Vu / 1000,
                'Vn(tf)': Vn / 1000,
                'DCR': round(Vu/Vn, 2)
            }
            result.append(cal_result)
            column.joint_result.update({pos: cal_result})

    return result, no_rebar_data, column_beam_df, beams_df
