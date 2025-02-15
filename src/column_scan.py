from __future__ import annotations

import pandas as pd
import numpy as np
import pprint

from item.column import Column
from math import ceil
from item.floor import read_parameter_df
from item.rebar import RebarDiameter
# from column_count import OutputExcel


class ColumnScan:
    scan_index: int
    index_score: dict[str, int]
    check_item: dict[str, str]
    criteria: str
    pass_criteria_formula: str
    ng_message: str
    check_function: function

    def __init__(self, kwagrs, scan_index) -> None:
        self.scan_index = scan_index
        self.check_item = {key: item for key,
                           item in kwagrs.items() if key[0] == 'Item'}
        self.criteria = kwagrs[('Criteria', '')]
        self.pass_criteria_formula = kwagrs[('PassCriteriaFormula', '')]
        self.ng_message = kwagrs[('NG Message', '')]
        self.index_score = {key: item for key,
                            item in kwagrs.items() if 'Index' in key[0]}

    def check(self, column: Column):
        try:
            return self.check_function(column)
        except:
            return "Error"

    def set_check_function(self, func):
        self.check_function = func


def rename_unnamed(df: pd.DataFrame):
    """Rename unamed columns name for Pandas DataFrame

    See https://stackoverflow.com/questions/41221079/rename-multiindex-columns-in-pandas

    Parameters
    ----------
    df : pd.DataFrame object
        Input dataframe

    Returns
    -------
    pd.DataFrame
        Output dataframe

    """
    for i, columns in enumerate(df.columns.levels):
        columns_new = columns.tolist()
        for j, row in enumerate(columns_new):
            if "Unnamed: " in str(row):
                columns_new[j] = ""
        if pd.__version__ < "0.21.0":  # https://stackoverflow.com/a/48186976/716469
            df.columns.set_levels(columns_new, level=i, inplace=True)
        else:
            df = df.rename(columns=dict(zip(columns.tolist(), columns_new)),
                           level=i)
    return df


def create_column_scan():
    column_scan_list: list[ColumnScan]
    column_scan_list = []
    df = read_parameter_df(r'file\柱SCAN.xlsx', '柱', [0, 1])
    df.set_index([('#', '0400')], inplace=True)
    df.fillna('', inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if df.loc[index][('Type', '柱')] == 'X' or np.isnan(index):
            continue
        column_scan = ColumnScan(df.loc[index].to_dict(), scan_index=index)
        set_check_scan(column_scan=column_scan)
        column_scan_list.append(column_scan)
    return column_scan_list


def output_detail_scan_report(column_list: list[Column]):
    import src.beam_scan as beam_scan
    return beam_scan.output_detail_scan_report(column_list)
    # ng_df = pd.DataFrame(columns = ['樓層','編號','備註'],index=[])
    # for b in column_list:
    #     for ng_message in b.ng_message:
    #         temp_df = pd.DataFrame(data={'樓層':b.floor,'編號':b.serial,'備註':ng_message},index=[0])
    #         ng_df = pd.concat([ng_df,temp_df],verify_integrity=True,ignore_index=True)
    # return ng_df


def output_ng_ratio(df: pd.DataFrame):
    import src.beam_scan as beam_scan
    return beam_scan.output_ng_ratio(df)


def column_check(column_list: list[Column], column_scan_list: list[ColumnScan]):
    df: pd.DataFrame
    df = pd.DataFrame(columns=[str(c.floor)+":"+str(c.serial)
                      for c in column_list], index=[cs.ng_message for cs in column_scan_list])
    for c in column_list:
        for cs in column_scan_list:
            df.loc[cs.ng_message, str(c.floor)+":"+str(c.serial)] = cs.check(c)
    return df


def set_check_scan(column_scan: ColumnScan):
    pass_syntax = 'OK'
    fail_syntax = 'NG.'

    def index_0401(c: Column):
        if not c.up_column:
            return '無上層柱'
        if c.total_As == 0:
            return '無鋼筋資料'
        if c.up_column.total_As / c.total_As < 0.6:
            return fail_syntax
        if len(c.up_column.x_row) - len(c.x_row) > 0:
            return f'X向倒插{len(c.up_column.x_row) - len(c.x_row)}支'
        if len(c.up_column.y_row) - len(c.y_row) > 0:
            return f'Y向倒插{len(c.up_column.y_row) - len(c.y_row)}支'
        return pass_syntax

    def index_0402(c: Column):
        if not c.bot_column:
            return '無下層柱'
        if c.total_As == 0:
            return '無鋼筋資料'
        if c.bot_column.total_As / c.total_As < 0.7:
            return fail_syntax
        return pass_syntax

    def index_0403(c: Column):
        if not c.confine_tie:
            return '無端部鋼筋'
        if c.fc == 0 or c.fy == 0:
            return '無樓層資料'
        x_Ash = c.confine_tie.Ash * (c.x_tie + 2)
        Ag = c.x_size * c.y_size
        Ach = (c.x_size - 8) * (c.y_size - 8)
        code15_3_X = 0.3 * (c.y_size - 8) * \
            c.confine_tie.spacing * c.fc/c.fy * (Ag/Ach - 1)
        code15_4_X = 0.09 * (c.y_size - 8) * c.confine_tie.spacing * c.fc/c.fy
        if not c.floor_object.is_seismic:
            code15_4_X = 0
        if x_Ash < code15_3_X or x_Ash < code15_4_X:
            # if True:
            c.ng_message.append(
                f'0403:X: {c.floor}{c.serial} => Ash = {round(x_Ash,2)} / 15.3Code = 0.3 * ({c.y_size} - 8)* {c.confine_tie.spacing} * {c.fc}/{c.fy} * ({Ag}/{Ach} - 1)={round(code15_3_X,2)}')
            c.ng_message.append(
                f'0403:X: {c.floor}{c.serial} => Ash = {round(x_Ash,2)} / 15.4Code = 0.09 * ({c.y_size} - 8) * {c.confine_tie.spacing} * {c.fc}/{c.fy} ={round(code15_4_X,2)}')
            return fail_syntax
        return pass_syntax

    def index_0404(c: Column):
        if not c.confine_tie:
            return '無端部鋼筋'
        if c.fc == 0 or c.fy == 0:
            return '無樓層資料'
        y_Ash = c.confine_tie.Ash * (c.y_tie + 2)
        Ag = c.x_size * c.y_size
        Ach = (c.x_size - 8) * (c.y_size - 8)
        code15_3_Y = 0.3 * (c.x_size-8) * \
            c.confine_tie.spacing * c.fc/c.fy * (Ag/Ach - 1)
        code15_4_Y = 0.09 * (c.x_size-8) * c.confine_tie.spacing * c.fc/c.fy
        if not c.floor_object.is_seismic:
            code15_4_Y = 0
        if y_Ash < code15_3_Y or y_Ash < code15_4_Y:
            # if True:
            c.ng_message.append(
                f'0404:Y: {c.floor}{c.serial} => Ash = {round(y_Ash,2)} / 15.3Code = 0.3 * ({c.x_size}-8) * {c.confine_tie.spacing} * {c.fc}/{c.fy} * ({Ag}/{Ach} - 1) ={round(code15_3_Y,2)}')
            c.ng_message.append(
                f'0404:Y: {c.floor}{c.serial} => Ash = {round(y_Ash,2)} / 15.4Code = 0.09 * ({c.x_size}-8) * {c.confine_tie.spacing} * {c.fc}/{c.fy} = {round(code15_4_Y,2)}')
            return fail_syntax
        return pass_syntax

    def index_0405(c: Column):
        if not c.total_rebar:
            return '無鋼筋資料'
        if c.total_As / (c.x_size*c.y_size) > 0.012:
            return fail_syntax
        return pass_syntax

    def index_0406(c: Column):
        if c.x_tie < ceil((len(c.y_row)-1)/2)-1:
            return fail_syntax
        return pass_syntax

    def index_0407(c: Column):
        if c.y_tie < ceil((len(c.x_row)-1)/2)-1:
            return fail_syntax
        return pass_syntax

    def index_0408(c: Column):
        if c.up_column and c.up_column.up_column:
            first = max(c.rebar, key=lambda r: RebarDiameter(r.size)).size
            second = max(c.up_column.rebar,
                         key=lambda r: RebarDiameter(r.size)).size
            third = max(c.up_column.up_column.rebar,
                        key=lambda r: RebarDiameter(r.size)).size
            if first != second and second != third:
                return fail_syntax
        return pass_syntax

    def index_0409(c: Column):

        rebar_dia = max(RebarDiameter(r[1]) for r in c.x_row)
        spacing = (c.x_size - c.protect_layer * 2 - 1.27 * 2 -
                   len(c.x_row)*rebar_dia)/(len(c.x_row) - 1)
        if spacing < 1.5*rebar_dia:
            c.ng_message.append(
                f'0409:X向{len(c.x_row)} 支 {rebar_dia} => 淨間距為{round(spacing,2)} < 1.5db:{round(1.5*rebar_dia,2)}')
            return fail_syntax
        return pass_syntax

    def index_0410(c: Column):
        rebar_dia = max(RebarDiameter(r[1]) for r in c.y_row)
        spacing = (c.y_size - c.protect_layer * 2 - 1.27 * 2 -
                   len(c.y_row)*rebar_dia)/(len(c.y_row) - 1)
        if spacing < 1.5*rebar_dia:
            c.ng_message.append(
                f'0409:Y向{len(c.y_row)} 支 {rebar_dia} => 淨間距為{round(spacing,2)} < 1.5db:{round(1.5*rebar_dia,2)}')
            return fail_syntax
        return pass_syntax
    if column_scan.scan_index == 401:
        column_scan.set_check_function(index_0401)
    if column_scan.scan_index == 402:
        column_scan.set_check_function(index_0402)
    if column_scan.scan_index == 403:
        column_scan.set_check_function(index_0403)
    if column_scan.scan_index == 404:
        column_scan.set_check_function(index_0404)
    if column_scan.scan_index == 405:
        column_scan.set_check_function(index_0405)
    if column_scan.scan_index == 406:
        column_scan.set_check_function(index_0406)
    if column_scan.scan_index == 407:
        column_scan.set_check_function(index_0407)
    if column_scan.scan_index == 408:
        column_scan.set_check_function(index_0408)
    if column_scan.scan_index == 409:
        column_scan.set_check_function(index_0409)
    if column_scan.scan_index == 410:
        column_scan.set_check_function(index_0410)
# def read_scan_excel(read_file:str,sheet_name:str):
#     return pd.read_excel(
#         read_file, sheet_name=sheet_name,header=[0,1])


if __name__ == '__main__':
    pass
