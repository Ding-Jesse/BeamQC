from __future__ import annotations
from item.column import Column
import pandas as pd
import pprint
from math import ceil
from item.floor import read_parameter_df
# from column_count import OutputExcel
class ColumnScan:
    scan_index:int
    index_score:dict[str,int]
    check_item:dict[str,str]
    criteria:str
    pass_criteria_formula:str
    ng_message:str
    check_function:function
    def __init__(self,kwagrs,scan_index) -> None:
        self.scan_index =scan_index
        self.check_item = {key:item for key,item in kwagrs.items() if key[0] == 'Item'}
        self.criteria = kwagrs[('Criteria', '')]
        self.pass_criteria_formula = kwagrs[('PassCriteriaFormula', '')]
        self.ng_message = kwagrs[('NG Message', '')]
        self.index_score = {key:item for key,item in kwagrs.items() if 'Index' in key[0]}
    def check(self,column:Column):
        return self.check_function(column)
    def set_check_function(self,func):
        self.check_function = func

def rename_unnamed(df:pd.DataFrame):
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
    column_scan_list:list[ColumnScan]
    column_scan_list = []
    df = read_parameter_df(r'file\柱SCAN.xlsx','柱',[0,1])
    df.set_index([('#','0400')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        column_scan = ColumnScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(column_scan=column_scan)
        column_scan_list.append(column_scan)
    return column_scan_list
def column_check(column_list:list[Column],column_scan_list:list[ColumnScan]):
    df:pd.DataFrame
    df = pd.DataFrame(columns=[str(c.floor)+str(c.serial) for c in column_list],index=[cs.ng_message for cs in column_scan_list])
    for c in column_list:
        for cs in column_scan_list:
            df.loc[cs.ng_message,str(c.floor)+str(c.serial)] = cs.check(c)
    return df
def set_check_scan(column_scan:ColumnScan):
    pass_syntax = 'OK'
    fail_syntax = 'NG.'
    def index_0401(c:Column):
        if not c.up_column:return '無上層柱'
        if c.total_As == 0: return '無鋼筋資料'
        if c.up_column.total_As / c.total_As < 0.6:
            return fail_syntax
        return pass_syntax
    def index_0402(c:Column):
        if not c.bot_column:return '無下層柱'
        if c.total_As == 0: return '無鋼筋資料'
        if c.bot_column.total_As / c.total_As < 0.7:
            return fail_syntax
        return pass_syntax
    def index_0403(c:Column):
        if not c.confine_tie:return '無端部鋼筋'
        if c.fc == 0 or c.fy == 0 :return '無樓層資料'
        x_Ash = c.confine_tie.Ash * (c.x_tie + 2)
        Ag = c.x_size * c.y_size
        Ach = (c.x_size - 8) * (c.y_size - 8)
        code15_3_X = 0.3 * (c.y_size - 8)* c.confine_tie.spacing * c.fc/c.fy * (Ag/Ach - 1)
        code15_4_X = 0.09 * (c.y_size - 8) * c.confine_tie.spacing * c.fc/c.fy
        if x_Ash < code15_3_X or x_Ash <code15_4_X:
            print(f'X: {c.floor}{c.serial} => Ash = {x_Ash} / 15.3Code = {code15_3_X}')
            print(f'X: {c.floor}{c.serial} => Ash = {x_Ash} / 15.4Code = {code15_4_X}')
            return fail_syntax
        return pass_syntax
    def index_0404(c:Column):
        if not c.confine_tie:return '無端部鋼筋'
        if c.fc == 0 or c.fy == 0 :return '無樓層資料'
        y_Ash = c.confine_tie.Ash * (c.y_tie + 2)
        Ag = c.x_size * c.y_size
        Ach = (c.x_size - 8) * (c.y_size - 8)
        code15_3_Y = 0.3 * (c.x_size-8) * c.confine_tie.spacing * c.fc/c.fy * (Ag/Ach - 1)
        code15_4_Y = 0.09 * (c.x_size-8) * c.confine_tie.spacing * c.fc/c.fy
        if y_Ash < code15_3_Y or y_Ash <code15_4_Y:
            print(f'Y: {c.floor}{c.serial} => Ash = {y_Ash} / 15.3Code = {code15_3_Y}')
            print(f'Y: {c.floor}{c.serial} => Ash = {y_Ash} / 15.4Code = {code15_4_Y}')
            return fail_syntax
        return pass_syntax
    def index_0405(c:Column):
        if not c.total_rebar:return '無鋼筋資料'
        if c.total_As /(c.x_size*c.y_size) > 0.012:
            return fail_syntax
        return pass_syntax
    def index_0406(c:Column):
        if c.x_tie <  ceil((len(c.y_row)-1)/2)-1:
            return fail_syntax
        return pass_syntax
    def index_0407(c:Column):
        if c.y_tie <  ceil((len(c.x_row)-1)/2)-1:
            return fail_syntax
        return pass_syntax
    if column_scan.scan_index == 401:column_scan.set_check_function(index_0401)       
    if column_scan.scan_index == 402:column_scan.set_check_function(index_0402)
    if column_scan.scan_index == 403:column_scan.set_check_function(index_0403)
    if column_scan.scan_index == 404:column_scan.set_check_function(index_0404)
    if column_scan.scan_index == 405:column_scan.set_check_function(index_0405)
    if column_scan.scan_index == 406:column_scan.set_check_function(index_0406)
    if column_scan.scan_index == 407:column_scan.set_check_function(index_0407)
# def read_scan_excel(read_file:str,sheet_name:str):
#     return pd.read_excel(
#         read_file, sheet_name=sheet_name,header=[0,1])

if __name__ == '__main__':
    pass