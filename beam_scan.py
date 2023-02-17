from __future__ import annotations
from column_scan import ColumnScan,rename_unnamed
from item.beam import Beam,RebarType
from math import sqrt,ceil
from item.rebar import RebarDiameter
from item.floor import read_parameter_df
from column_scan import rename_unnamed
import pandas as pd
class BeamScan(ColumnScan):
    pass
def beam_check(beam_list:list[Beam],beam_scan_list:list[BeamScan]):
    df:pd.DataFrame
    df = pd.DataFrame(columns=[str(b.floor)+str(b.serial) for b in beam_list],index=[bs.ng_message for bs in beam_scan_list])
    for b in beam_list:
        for bs in beam_scan_list:
            df.loc[bs.ng_message,str(b.floor)+str(b.serial)] = bs.check(b)
    return df

def create_beam_scan():
    beam_scan_list:list[ColumnScan]
    beam_scan_list = []
    df = read_parameter_df(r'D:\Desktop\BeamQC\TEST\柱SCAN.xlsx','梁',[0,1])
    df.set_index([('#','0200')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        beam_scan = BeamScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(beam_scan=beam_scan)
        beam_scan_list.append(beam_scan)
    return beam_scan_list

def set_check_scan(beam_scan:BeamScan):
    pass_syntax = 'OK'
    fail_syntax = 'NG.'
    protect_layer = 7
    def get_code_3_6(b:Beam):
        code3_3 = 0.8*sqrt(b.fc)/b.fy*b.width*(b.depth-protect_layer)
        code3_4 = 14/b.fy*b.width*(b.depth-protect_layer)
        return code3_3,code3_4
    def get_code_15_4_2_1(b:Beam):
        code15_4_2 = (b.fc + 100)/(4*b.fy)
        code15_4_2_1 = 0.025
        return code15_4_2 ,code15_4_2_1
    def index_0201(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0202(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0203(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0204(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0205(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0206(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0207(b:Beam):
        if b.depth >= 90:
            if len(b.middle_tie) < ceil((b.depth - 15 - 10)/30 - 1):
                return fail_syntax
        return pass_syntax
    def index_0208(b:Beam):
        if b.depth*4 > b.length:
            return fail_syntax
        return pass_syntax
    def index_0209(b:Beam):
        for pos,tie in b.tie.items():
            Vs = tie.Ash*2*b.fy*(b.depth - protect_layer)/tie.spacing
            if Vs > 2.12*sqrt(b.fc)*b.width*(b.depth - protect_layer):
                return fail_syntax
        return pass_syntax
    def index_0210(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                spacing = (b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1)
                if spacing < RebarDiameter(rebar.size):
                    return fail_syntax
        return pass_syntax
    def index_0211(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                if rebar.number < 2:
                    return fail_syntax
        return pass_syntax
    def index_0212(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1 :
            return fail_syntax
        return pass_syntax
    def index_0213(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            return fail_syntax
        return pass_syntax
    def index_0214(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            return fail_syntax
        return pass_syntax
    def index_0215(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            return fail_syntax
        return pass_syntax
    def index_0216(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            return fail_syntax
        return pass_syntax
    def index_0217(b:Beam):
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            return fail_syntax
        return pass_syntax
    def index_0218(b:Beam):
        rebarAs = []
        for rebar_type1 in [RebarType.Top , RebarType.Bottom]:
            for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
                rebarAs.append(b.get_rebar_table(rebar_type1=rebar_type1,rebar_type2=rebar_type2))
        if any([r for r in rebarAs if r < 0.25*max(rebarAs)]):
            return fail_syntax
        for i in [0,2]:
            if not 0.5 <= rebarAs[i]/rebarAs[i+3] <= 2:
                return fail_syntax
        return pass_syntax
    def index_0219(b:Beam):
        for pos,tie in b.tie.items():
            if tie.spacing < 10:
                return fail_syntax
        return pass_syntax
    def index_0220(b:Beam):
        for pos,tie in b.tie.items():
            if tie.spacing > 30:
                return fail_syntax
        return pass_syntax
    def index_0221(b:Beam):
        middle_number = b.middle_tie[0].text if b.middle_tie else 0
        if middle_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle) > \
                max(b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right),
                    b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left))*0.7:
                return fail_syntax
        return pass_syntax
    def index_0222(b:Beam):
        if b.length < b.depth * 4:
            if 0.0015 * b.width > sum([r.As for r in b.middle_tie]):
                return fail_syntax
        return pass_syntax
    if beam_scan.scan_index == 201:beam_scan.set_check_function(index_0201)       
    if beam_scan.scan_index == 202:beam_scan.set_check_function(index_0202)       
    if beam_scan.scan_index == 203:beam_scan.set_check_function(index_0203)       
    if beam_scan.scan_index == 204:beam_scan.set_check_function(index_0204)   
    if beam_scan.scan_index == 205:beam_scan.set_check_function(index_0205)       
    if beam_scan.scan_index == 206:beam_scan.set_check_function(index_0206)       
    if beam_scan.scan_index == 207:beam_scan.set_check_function(index_0207)       
    if beam_scan.scan_index == 208:beam_scan.set_check_function(index_0208)  
    if beam_scan.scan_index == 209:beam_scan.set_check_function(index_0209)       
    if beam_scan.scan_index == 210:beam_scan.set_check_function(index_0210)       
    if beam_scan.scan_index == 211:beam_scan.set_check_function(index_0211)       
    if beam_scan.scan_index == 212:beam_scan.set_check_function(index_0212) 
    if beam_scan.scan_index == 213:beam_scan.set_check_function(index_0213)       
    if beam_scan.scan_index == 214:beam_scan.set_check_function(index_0214)       
    if beam_scan.scan_index == 215:beam_scan.set_check_function(index_0215)       
    if beam_scan.scan_index == 216:beam_scan.set_check_function(index_0216)   
    if beam_scan.scan_index == 217:beam_scan.set_check_function(index_0217)       
    if beam_scan.scan_index == 218:beam_scan.set_check_function(index_0218)       
    if beam_scan.scan_index == 219:beam_scan.set_check_function(index_0219)       
    if beam_scan.scan_index == 220:beam_scan.set_check_function(index_0220)  
    if beam_scan.scan_index == 221:beam_scan.set_check_function(index_0221)       
    if beam_scan.scan_index == 222:beam_scan.set_check_function(index_0222)       