from __future__ import annotations
from column_scan import ColumnScan,rename_unnamed
from item.beam import Beam,RebarType,Rebar,Tie
from math import sqrt,ceil
from item.rebar import RebarDiameter
from item.floor import read_parameter_df
from column_scan import rename_unnamed
import pandas as pd
import numpy as np
class BeamScan(ColumnScan):
    pass
def beam_check(beam_list:list[Beam],beam_scan_list:list[BeamScan]):
    df:pd.DataFrame
    enoc_list = [bs for bs in beam_scan_list if bs.index_score[('Index2','經濟性')]== 1]
    code_list = [bs for bs in beam_scan_list if bs not in enoc_list]
    enoc_df = pd.DataFrame(columns=[str(b.floor)+':'+str(b.serial) for b in beam_list],index=[bs.ng_message for bs in enoc_list])
    code_df = pd.DataFrame(columns=[str(b.floor)+':'+str(b.serial) for b in beam_list],index=[bs.ng_message for bs in code_list])
    for b in beam_list:
        for bs in enoc_list:
            enoc_df.loc[bs.ng_message,str(b.floor)+':'+str(b.serial)] = bs.check(b)
        for bs in code_list:
            code_df.loc[bs.ng_message,str(b.floor)+':'+str(b.serial)] = bs.check(b)
    return enoc_df,code_df
def output_detail_scan_report(beam_list:list[Beam]):
    ng_df = pd.DataFrame(columns = ['樓層','編號','檢核項目','備註'],index=[])
    for b in beam_list:
        for ng_message in b.ng_message:
            ng_serial = '0000'
            try:
                ng_serial = ng_message.split(':')[0]
            except:
                pass
            temp_df = pd.DataFrame(data={'樓層':b.floor,'編號':b.serial,'檢核項目':ng_serial,'備註':ng_message},index=[0])
            ng_df = pd.concat([ng_df,temp_df],verify_integrity=True,ignore_index=True)
    return ng_df
def output_ng_ratio(df:pd.DataFrame):
    ng_df = pd.DataFrame(columns = ['樓層','編號','NG項目','NG率(NG項目/總檢核項目)'],index=[])
    df = df.loc[:, ~df.columns.duplicated()]
    for column in df:
        ng_item = list(df.loc[df[column] == "NG."].index)
        serial = column.split(":")
        temp_df = pd.DataFrame(data={'樓層':serial[0],'編號':serial[1],'NG項目':'\n'.join(ng_item),'NG率(NG項目/總檢核項目)':f'{len(ng_item)}/{len(df.index)}'},index=[0])
        ng_df = pd.concat([ng_df,temp_df],verify_integrity=True,ignore_index=True)
    #     count_df = df[column].value_counts()
    #     if "NG." in count_df.index:
    #         df[column].loc['Total_NG'] = count_df.loc["NG."]
    #     else:
    #         df[column].loc['Total_NG'] = 0
    #     if "OK" in count_df.index:
    #         df[column].loc['Total_Ok'] = count_df.loc["OK"]
    #     else:
    #         df[column].loc['Total_Ok'] = 0
    # new_df = df.apply(lambda s: s.value_counts(),axis=1).fillna(0)
    # new_df = df.apply(pd.value_counts).fillna(0).astype(int)
    sum_df = df.apply(pd.Series.value_counts, axis=1).fillna(0).astype(int)
    # print(new_df)
    # print(ng_df)
    return ng_df,sum_df
def create_beam_scan():
    beam_scan_list:list[BeamScan]
    beam_scan_list = []
    df = read_parameter_df(r'file\柱SCAN.xlsx','梁',[0,1])
    df.set_index([('#','0200')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if np.isnan(index): continue
        beam_scan = BeamScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(beam_scan=beam_scan)
        beam_scan_list.append(beam_scan)
    return beam_scan_list

def create_sbeam_scan():
    beam_scan_list:list[BeamScan]
    beam_scan_list = []
    df = read_parameter_df(r'file\柱SCAN.xlsx','小梁',[0,1])
    df.set_index([('#','0300')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if df.loc[index][('Type','小梁')] == 'X' and np.isnan(index): continue
        beam_scan = BeamScan(df.loc[index].to_dict(),scan_index=index)
        set_check_scan(beam_scan=beam_scan)
        beam_scan_list.append(beam_scan)
    return beam_scan_list

def create_fbeam_scan():
    beam_scan_list:list[BeamScan]
    beam_scan_list = []
    df = read_parameter_df(r'file\柱SCAN.xlsx','地梁',[0,1])
    df.set_index([('#','0100')],inplace=True)
    df.fillna('',inplace=True)
    df = rename_unnamed(df=df)
    for index in df.index:
        if df.loc[index][('Type','地梁')] == 'X' or np.isnan(index): continue
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
        clear_depth = (b.depth - b.floor_object.slab_height['top'] - b.floor_object.slab_height['bot'])
        if b.length < clear_depth * 4 and b.middle_tie:
            for pos,tie in b.tie.items():
                if tie is None:continue
                if 0.0025 * b.width > tie.Ash/tie.spacing:
                    b.ng_message.append(f'0101:0.0025 * {b.width} > {tie.Ash}/{tie.spacing} => {round(0.0025 * b.width,2)} > {round(tie.Ash/tie.spacing,2)}')
                    return fail_syntax
        return pass_syntax
    def index_0102(b:Beam):
        clear_depth = (b.depth - b.floor_object.slab_height['top'] - b.floor_object.slab_height['bot'])
        if b.length < clear_depth * 4 and b.middle_tie:
            if 0.0015 * b.width * clear_depth > b.middle_tie[0].As*2:
                b.ng_message.append(f'0102:0.0015 * {b.width} * {clear_depth} > 2 * {round(b.middle_tie[0].As,2)} => {round(0.0015 * b.width * clear_depth,2)} > {round(b.middle_tie[0].As*2)}')
                return fail_syntax
        return pass_syntax
    def index_0103(b:Beam):
        clear_depth = (b.depth - b.floor_object.slab_height['top'] - b.floor_object.slab_height['bot'])
        if b.length < clear_depth * 4 and b.middle_tie:
            if 0.0015 * b.width * clear_depth * 1.5 < b.middle_tie[0].As:
                b.ng_message.append(f'0103:0.0015 *{b.width} * {clear_depth} < 1.5 * {round(b.middle_tie[0].As,2)} => {round(0.0015 * b.width * clear_depth,2)} < {round(1.5*b.middle_tie[0].As,2)}')
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
        for rebar_dict in [b.rebar_table[RebarType.Top.value],b.rebar_table[RebarType.Bottom.value]]:
            for pos2, rebar_list in rebar_dict.items():
                    if rebar_list:
                        rebar = rebar_list[0]
                        spacing = round((b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1),2)
                        if spacing > 25 :
                            b.ng_message.append(f'0106:({b.width}- 4*2 - 1.27*2 -{RebarDiameter(rebar.size)})/{rebar.number - 1} = {spacing}cm > 25 cm')
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
        if not b.rebar_table['bottom']['middle']:return f'無鋼筋資料'
        middle_number = b.rebar_table['bottom']['middle'][0].number
        if middle_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle) > \
                max(b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right),
                    b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left))*0.7:
                return fail_syntax
        return pass_syntax
    def index_0111(b:Beam):
        if not b.rebar_table['top']['left']:return f'無鋼筋資料'
        left_number = b.rebar_table['top']['left'][0].number
        if left_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left) > \
                b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)*0.7:
                return fail_syntax
        return pass_syntax
    def index_0112(b:Beam):
        if not b.rebar_table['top']['right']:return f'無鋼筋資料'
        right_number = b.rebar_table['top']['right'][0].number
        if right_number > 3:
            if b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right) > \
                b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)*0.7:
                return fail_syntax
        return pass_syntax
    def index_0113(b:Beam):
        if all([tie is None for pos,tie in b.tie.items()]): return '無箍筋資料'
        for pos,tie in b.tie.items():
            if tie is None:continue
            if tie.spacing < 10 :
                return fail_syntax
        return pass_syntax
    def index_0114(b:Beam):
        if all([tie is None for pos,tie in b.tie.items()]): return '無箍筋資料'
        for pos,tie in b.tie.items():
            if tie is None:continue
            if tie.spacing > 30 :
                return fail_syntax
        return pass_syntax     
    def get_code_3_6(b:Beam,rebar_type1:RebarType,rebar_type2:RebarType):
        fy = b.fy
        rebar_list = b.rebar_table[rebar_type1.value][rebar_type2.value]
        if rebar_list:
            fy = rebar_list[0].fy
        code3_3 = 0.8*sqrt(b.fc)/fy*b.width*(b.depth-protect_layer)
        code3_4 = 14/fy*b.width*(b.depth-protect_layer)
        return code3_3,code3_4
    def get_code_15_4_2_1(b:Beam,rebar_type1:RebarType,rebar_type2:RebarType):
        fy = b.fy
        rebar_list = b.rebar_table[rebar_type1.value][rebar_type2.value]
        if rebar_list:
            fy = rebar_list[0].fy
        code15_4_2 = (b.fc + 100)/(4*fy)
        code15_4_2_1 = 0.025
        return code15_4_2 ,code15_4_2_1
    def index_0201(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0201:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0202(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0202:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0203(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Right)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0203:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0204(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0204:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0205(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0205:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0206(b:Beam):
        code3_3 ,code3_4= get_code_3_6(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        if  rebar_As < code3_3 or rebar_As < code3_4:
            b.ng_message.append(f'0206:max(code3_3:{round(code3_3,2)}cm2 ,code3_4:{round(code3_4,2)}cm2) > 鋼筋總面積:{rebar_As}')
            return fail_syntax
        return pass_syntax
    def index_0207(b:Beam):
        if b.depth >= 90:
            if len(b.middle_tie) < ceil((b.depth - b.floor_object.slab_height['top'] - 10)/30 - 1):
                b.ng_message.append(f'0207:腰筋支數:{len(b.middle_tie)} < (梁深{b.depth} - 上版厚{b.floor_object.slab_height["top"]} - 鋼筋中心至邊緣距離{10})/30 - 1 = {round(ceil((b.depth - b.floor_object.slab_height["top"] - 10)/30 - 1),2)}')
                return fail_syntax
        return pass_syntax
    def index_0208(b:Beam):
        clear_depth = (b.depth - b.floor_object.slab_height['top'] - b.floor_object.slab_height['bot'])
        if clear_depth*4 > b.length:
            return fail_syntax
        return pass_syntax
    def index_0209(b:Beam):
        if all([tie is None for pos,tie in b.tie.items()]): return '無箍筋資料'
        for pos,tie in b.tie.items():
            if tie is None:continue
            Vs = round(tie.Ash * tie.fy*(b.depth - protect_layer)/tie.spacing,2)
            if Vs > 2.12*sqrt(b.fc)*b.width*(b.depth - protect_layer):
                b.ng_message.append(f'0209:Vs:{Vs}  > 4Vc:{round(2.12*sqrt(b.fc)*b.width*(b.depth - protect_layer),2)}')
                return fail_syntax
        return pass_syntax
    def index_0210(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                if rebar.number <= 1 : continue
                spacing = round((b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1),2)
                if spacing < RebarDiameter(rebar.size) or spacing < 2.5:
                    b.ng_message.append(f'0210:單排淨距={spacing} < ({RebarDiameter(rebar.size)},2.5)')
                    return fail_syntax
        return pass_syntax
    def index_0211(b:Beam):
        for pos,rebar_list in b.rebar.items():
            for rebar in rebar_list:
                if rebar.number < 2:
                    return fail_syntax
        return pass_syntax
    def index_0212(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1 :
            b.ng_message.append(f'0212:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0213(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            b.ng_message.append(f'0213:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0214(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Top,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            b.ng_message.append(f'0214:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0215(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Left)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            b.ng_message.append(f'0215:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0216(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Middle)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            b.ng_message.append(f'0216:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0217(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        code15_4_2 ,code15_4_2_1 = get_code_15_4_2_1(b=b,rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        rebar_As = b.get_rebar_table(rebar_type1=RebarType.Bottom,rebar_type2=RebarType.Right)
        if rebar_As/(b.width * (b.depth - protect_layer))> code15_4_2 or rebar_As/(b.width * (b.depth - protect_layer)) > code15_4_2_1:
            b.ng_message.append(f'0217:鋼筋As:{rebar_As}/梁面積:{(b.width * (b.depth - protect_layer))} = {rebar_As/(b.width * (b.depth - protect_layer))} < max(code15_4_2:{code15_4_2} , code15_4_2_1{code15_4_2_1})')
            return fail_syntax
        return pass_syntax
    def index_0218(b:Beam):
        if not b.floor_object.is_seismic:return "不檢核此項"
        rebarAs = []
        for rebar_type1 in [RebarType.Top , RebarType.Bottom]:
            for rebar_type2 in [RebarType.Left , RebarType.Middle, RebarType.Right]:
                rebarAs.append(b.get_rebar_table(rebar_type1=rebar_type1,rebar_type2=rebar_type2))
        if any([r for r in rebarAs if r < 0.25*max(rebarAs)]):
            temp = [r for r in rebarAs if r < 0.25*max(rebarAs)]
            b.ng_message.append(f'0218:{temp}cm2 < code15_4_2_2:{0.25*max(rebarAs)}cm2')
            return fail_syntax
        for i,pos in enumerate(['左','右']):
            if rebarAs[i+3] == 0:return "無鋼筋資料"
            if 0.5 > rebarAs[i]/rebarAs[i+3]:
                b.ng_message.append(f'0218:位置:{pos}端上層/{pos}端下層 = {round(rebarAs[i]/rebarAs[i+3],2)}')
                return fail_syntax
            if rebarAs[i]/rebarAs[i+3] > 2:
                b.ng_message.append(f'0218:位置:{pos}端下層/{pos}端上層 = {round(rebarAs[i+3]/rebarAs[i],2)}')
                return fail_syntax
        return pass_syntax
    def index_0219(b:Beam):
        if all([tie is None for pos,tie in b.tie.items()]): return '無箍筋資料'
        for pos,tie in b.tie.items():
            if tie.spacing < 10:
                return fail_syntax
        return pass_syntax
    def index_0220(b:Beam):
        if all([tie is None for pos,tie in b.tie.items()]): return '無箍筋資料'
        for pos,tie in b.tie.items():
            if tie is None:continue
            if tie.spacing > 30:
                return fail_syntax
        return pass_syntax
    def index_0221(b:Beam):
        if not b.rebar_table['bottom']['middle']:return f'無鋼筋資料'
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
    def index_0223(b:Beam):
        for rebar_type1 in [RebarType.Top.value,RebarType.Bottom.value]:
            for rebar_type2 in [RebarType.Right.value,RebarType.Middle.value,RebarType.Left.value]:
                if b.rebar_table[rebar_type1][rebar_type2]:
                    rebar = b.rebar_table[rebar_type1][rebar_type2][0]
                    spacing = (b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1)
                    if spacing < 1.5 * RebarDiameter(rebar.size):
                        b.ng_message.append(f'{rebar.text} : 間距為{spacing} cm < 1.5db :{1.5 * RebarDiameter(rebar.size)}')
                        return fail_syntax
        return pass_syntax

    def index_0301(b:Beam):
        return index_0201(b=b)
    def index_0302(b:Beam):
        code3_3,code3_4 = get_code_3_6(b=b,rebar_type1=RebarType.Top,rebar_type2=RebarType.Middle)
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
        if rebar_As == 0:return '無鋼筋資料'
        rebar_fy = b.rebar_table[RebarType.Bottom.value][RebarType.Middle.value][0].fy
        if 0.6*(1/8)*wu*L**2 >  rebar_As*rebar_fy*(b.depth - protect_layer):
            return fail_syntax
        return pass_syntax
    def index_0313(b:Beam):
        for rebar_type1 in [RebarType.Top.value,RebarType.Bottom.value]:
            for rebar_type2 in [RebarType.Right.value,RebarType.Middle.value,RebarType.Left.value]:
                if b.rebar_table[rebar_type1][rebar_type2]:
                    rebar = b.rebar_table[rebar_type1][rebar_type2][0]
                    spacing = (b.width - 4*2 - 1.27*2 - RebarDiameter(rebar.size))/(rebar.number - 1)
                    if spacing < 1.2 * RebarDiameter(rebar.size):
                        b.ng_message.append(f'{rebar.text} : 間距為{spacing} cm < 1.5db :{1.5 * RebarDiameter(rebar.size)}')
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
    if beam_scan.scan_index == 223:beam_scan.set_check_function(index_0223)

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
    if beam_scan.scan_index == 313:beam_scan.set_check_function(index_0313)