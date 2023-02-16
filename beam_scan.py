from __future__ import annotations
from column_scan import ColumnScan,read_scan_excel,rename_unnamed
from item.beam import Beam,RebarType
from math import sqrt,floor,ceil
import pandas as pd
class BeamScan(ColumnScan):
    def __init__():
        pass
def beam_check(beam_list:list[Beam],beam_scan_list:list[BeamScan]):
    df:pd.DataFrame
    df = pd.DataFrame(columns=[str(c.floor)+str(c.serial) for c in beam_list],index=[cs.ng_message for cs in beam_scan_list])
    for c in beam_list:
        for cs in beam_scan_list:
            df.loc[cs.ng_message,str(c.floor)+str(c.serial)] = cs.check(c)
    return df

def set_check_scan(beam_scan:BeamScan):
    pass_syntax = 'OK'
    fail_syntax = 'NG.'
    def get_code_3_6(b:Beam):
        code3_3 = 0.8*sqrt(b.fc)/b.fy*b.width*(b.depth-7)
        code3_4 = 14/b.fy*b.width*(b.depth-7)
        return code3_3,code3_4
    def index201(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index202(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index203(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index204(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index205(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index206(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index207(b:Beam):
        if b.depth >= 90:
            if len(b.middle_tie) < ceil((b.depth - 15 - 10)/30 - 1):
                return fail_syntax
        return pass_syntax
    def index208(b:Beam):
        if b.depth*4 > b.length:
            return fail_syntax
        return pass_syntax
    def index209(b:Beam):
        for pos,tie in b.tie.items():
            Vs = tie.size*2*b.fy*(b.depth - 7)/tie.
            pass