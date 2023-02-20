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
    beam_scan_list:list[BeamScan]
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

def create_sbeam_scan():
    beam_scan_list:list[BeamScan]
    beam_scan_list = []
    df = read_parameter_df(r'D:\Desktop\BeamQC\TEST\柱SCAN.xlsx','小梁',[0,1])
    df.set_index([('#','0300')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if df.loc[index][('Type','小梁')] == 'X': continue
        beam_scan = BeamScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(beam_scan=beam_scan)
        beam_scan_list.append(beam_scan)
    return beam_scan_list

def create_fbeam_scan():
    beam_scan_list:list[BeamScan]
    beam_scan_list = []
    df = read_parameter_df(r'D:\Desktop\BeamQC\TEST\柱SCAN.xlsx','地梁',[0,1])
    df.set_index([('#','0100')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if df.loc[index][('Type','地梁')] == 'X': continue
        beam_scan = BeamScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(beam_scan=beam_scan)
        beam_scan_list.append(beam_scan)
    return beam_scan_list

def set_sb_check_scan(beam_scan:BeamScan):
    pass

def set_check_scan(beam_scan:BeamScan):
    pass_syntax = 'OK'
    fail_syntax = 'NG.'
    protect_layer = 7
    def index_0101(b:Beam):
        for pos,tie in b.tie.items():
            if 0.0025 * b.width > tie.Ash/tie.spacing:
                return fail_syntax
        return pass_syntax
    def index_0102(b:Beam):
        if b.length < b.depth * 4 and b.middle_tie:
            if 0.0015 * b.width * 1.5 < b.middle_tie[0].As:
                return fail_syntax
        return pass_syntax
    def index_0103(b:Beam):
        if b.length < b.depth * 4 and b.middle_tie:
            if 0.0015 * b.width * 1.5 < b.middle_tie[0].As:
                return fail_syntax
        return pass_syntax
    def index_0104(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Top,rebar_type2=rebar_type2)
            if rebar_As/(b.width * (b.depth - protect_layer)) < 0.003 :
                return fail_syntax
        return pass_syntax
    def index_0105(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Bottom,rebar_type2=rebar_type2)
            if rebar_As/(b.width * (b.depth - protect_layer)) < 0.003 :
                return fail_syntax
        return pass_syntax
    def index_0106(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                spacing = (b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1)
                if spacing < 25 :
                    return fail_syntax
        return pass_syntax
    def index_0107(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                if rebar.number < 2 :
                    return fail_syntax
        return pass_syntax

    def index_0108(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Top,rebar_type2=rebar_type2)
            if rebar_As/(b.width * (b.depth - protect_layer)) > 0.02 :
                return fail_syntax
        return pass_syntax
    def index_0109(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Bottom,rebar_type2=rebar_type2)
            if rebar_As/(b.width * (b.depth - protect_layer)) > 0.02 :
                return fail_syntax
        return pass_syntax 
    def index_0110(b:Beam):
        middle_number = b.rebar_table['bottom']['middle'][0].number
        if middle_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle) > \
                max(b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right),
                    b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left))*0.7:
                return fail_syntax
        return pass_syntax
    def index_0111(b:Beam):
        left_number = b.rebar_table['top']['left'][0].number
        if left_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left) > \
                b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)*0.7:
                return fail_syntax
        return pass_syntax
    def index_0112(b:Beam):
        right_number = b.rebar_table['top']['right'][0].number
        if right_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right) > \
                b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)*0.7:
                return fail_syntax
        return pass_syntax
    def index_0113(b:Beam):
        for pos,tie in b.tie.items():
            if tie.spacing < 10 :
                return fail_syntax
        return pass_syntax
    def index_0114(b:Beam):
        for pos,tie in b.tie.items():
            if tie.spacing > 30 :
                return fail_syntax
        return pass_syntax     
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
        if rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0202(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if rebar_As < code3_3 or rebar_As < code3_4:
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
        middle_number = b.rebar_table['bottom']['middle'][0].number
        if middle_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle) > \
                max(b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right),
                    b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left))*0.7:
                return fail_syntax
        return pass_syntax
    def index_0222(b:Beam):
        if b.length < b.depth * 4 and b.middle_tie:
            if 0.0015 * b.width * 1.5 < b.middle_tie[0].As:
                return fail_syntax
        return pass_syntax
    def index_0301(b:Beam):
        return index_0201(b=b)
    def index_0302(b:Beam):
        code3_3,code3_4 = get_code_3_6(b=b)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            return fail_syntax
        return pass_syntax
    def index_0303(b:Beam):
        return index_0203(b=b)
    def index_0304(b:Beam):
        return index_0204(b=b)
    def index_0305(b:Beam):
        return index_0205(b=b)
    def index_0306(b:Beam):
        return index_0206(b=b)
    def index_0307(b:Beam):
        return index_0207(b=b)
    def index_0308(b:Beam):
        return index_0210(b=b)
    def index_0309(b:Beam):
        return index_0211(b=b)
    def index_0310(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Top,rebar_type2=rebar_type2)
            if rebar_As / (b.width*(b.depth-protect_layer)) > 0.025:
                return fail_syntax
        return pass_syntax
    def index_0311(b:Beam):
        for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
            rebar_As = b.get_rebar_table(rebar_type1= RebarType.Bottom,rebar_type2=rebar_type2)
            if rebar_As / (b.width*(b.depth-protect_layer)) > 0.025:
                return fail_syntax
        return pass_syntax
    def index_0312(b:Beam):
        band_width = 300
        L = b.length
        wu = b.get_loading(band_width=band_width)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if 0.6*(1/8)*wu*L**2 >  rebar_As*b.fy*(b.depth - protect_layer):
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

    if beam_scan.scan_index == 101:beam_scan.set_check_function(index_0101)
    if beam_scan.scan_index == 102:beam_scan.set_check_function(index_0102)
    if beam_scan.scan_index == 103:beam_scan.set_check_function(index_0103)
    if beam_scan.scan_index == 104:beam_scan.set_check_function(index_0104)
    if beam_scan.scan_index == 105:beam_scan.set_check_function(index_0105)
    if beam_scan.scan_index == 106:beam_scan.set_check_function(index_0106)
    if beam_scan.scan_index == 107:beam_scan.set_check_function(index_0107)
    if beam_scan.scan_index == 108:beam_scan.set_check_function(index_0108)
    if beam_scan.scan_index == 109:beam_scan.set_check_function(index_0109)
    if beam_scan.scan_index == 110:beam_scan.set_check_function(index_0110)
    if beam_scan.scan_index == 111:beam_scan.set_check_function(index_0111)
    if beam_scan.scan_index == 112:beam_scan.set_check_function(index_0112)
    if beam_scan.scan_index == 113:beam_scan.set_check_function(index_0113)
    if beam_scan.scan_index == 114:beam_scan.set_check_function(index_0114)

    if beam_scan.scan_index == 301:beam_scan.set_check_function(index_0301)
    if beam_scan.scan_index == 302:beam_scan.set_check_function(index_0302)
    if beam_scan.scan_index == 303:beam_scan.set_check_function(index_0303)
    if beam_scan.scan_index == 304:beam_scan.set_check_function(index_0304)
    if beam_scan.scan_index == 305:beam_scan.set_check_function(index_0305)
    if beam_scan.scan_index == 306:beam_scan.set_check_function(index_0306)
    if beam_scan.scan_index == 307:beam_scan.set_check_function(index_0307)
    if beam_scan.scan_index == 308:beam_scan.set_check_function(index_0308)
    if beam_scan.scan_index == 309:beam_scan.set_check_function(index_0309)
    if beam_scan.scan_index == 310:beam_scan.set_check_function(index_0310)
    if beam_scan.scan_index == 311:beam_scan.set_check_function(index_0311)
    if beam_scan.scan_index == 312:beam_scan.set_check_function(index_0312)