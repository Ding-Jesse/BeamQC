from __future__ import annotations
from item import column,beam
from collections import defaultdict
from collections import Counter
from numpy import arange,empty
import pandas as pd
class Floor:
    height:float
    material_list:dict[str,float]
    column_list:list[column.Column]
    beam_list:list[beam.Beam]
    overlap_option:dict[str,str]
    rebar_count:dict[str,float]
    beam_rebar_count:defaultdict[str,float]
    concrete_count:defaultdict[str,float]
    formwork_count:float
    coupler:dict[str,float]
    floor_name:str
    loading:dict[str,float]
    is_seismic:bool
    slab_height:dict[str,float]
    def __init__(self,floor_name):
        if floor_name[-1] != 'F':
            floor_name += 'F'
        self.floor_name = floor_name
        self.rebar_count ={}
        self.column_list = []
        self.beam_list = []
        self.material_list = {}
        self.overlap_option ={}
        self.slab_height = {}
        self.concrete_count =defaultdict(lambda:0)
        self.beam_rebar_count = defaultdict(lambda:0)
        self.loading = defaultdict(lambda:0)
        self.coupler = {}
        self.formwork_count = 0
        self.is_seismic = False
        pass

    def set_beam_prop(self,kwargs):
        self.material_list.update({'fc':kwargs["混凝土強度fc'(kgf/cm2)"]})
        self.material_list.update({'fy':kwargs["鋼筋強度fy(kgf/cm2)"]})
        self.loading.update({"SDL":kwargs["SDL(t/m2)"]})
        self.loading.update({"LL":kwargs["LL(t/m2)"]})
        self.is_seismic = kwargs["是否需耐震"] == "是"
        self.slab_height.update({'top':kwargs["上版厚(cm)"]})
        self.slab_height.update({'bot':kwargs["下版厚(cm)"]})
        try: 
            self.height = float(kwargs["樓高"])
        except:
            self.height = 0
    def set_column_prop(self,kwargs):
        self.material_list.update({'fc':kwargs["混凝土強度fc'(kgf/cm2)"]})
        self.material_list.update({'fy':kwargs["鋼筋強度fy(kgf/cm2)"]})
        self.overlap_option.update({"tight_tie":kwargs["全段緊密"],"coupler":kwargs["續接器"],"overlap":kwargs["續接方式"]})
        try: 
            self.height = float(kwargs["樓高"])
        except:
            self.height = 0
        self.is_seismic = kwargs["是否需耐震"] == "是"

    def add_column(self,c_list:list[column.Column]):
        if not c_list:return
        for c in c_list:
            c.set_prop(self)
            c.floor_object = self
        self.column_list.extend(c_list)
    def add_beam(self,beam_list:list[beam.Beam]):
        if not beam_list:return
        for b in beam_list:
            b.set_prop(self)
        self.beam_list.extend(beam_list)
    def summary_rebar(self):
        for c in self.column_list:
            for size,count in c.rebar_count.items():
                if not size in self.rebar_count : self.rebar_count[size] = 0
                self.rebar_count[size] += round(count/1000/1000,2)
            for size,coupler in c.coupler.items():
                if size == ('',''):continue
                if not size in self.coupler : self.coupler[size] = 0
                if self.overlap_option['overlap'] == "隔根隔層續接":
                    self.coupler[size] += coupler//2
                else:
                    self.coupler[size] += coupler
            # if not c.fc in self.concrete_count:self.concrete_count[c.fc] = 0
            self.concrete_count[c.fc] += c.concrete
            self.formwork_count += c.formwork
        self.rebar_count['total'] = sum(self.rebar_count.values())
    def summary_beam(self):
        for b in self.beam_list:
            for size,count in b.rebar_count.items():
                self.beam_rebar_count[size] += round(count/1000/1000,2)
            for size,count in b.tie_count.items():
                self.beam_rebar_count[size] += round(count/1000/1000,2)
            self.concrete_count[b.fc] +=   b.concrete
            self.formwork_count += b.formwork
        self.beam_rebar_count['total'] = sum(self.beam_rebar_count.values())
def read_parameter_df(read_file,sheet_name,header_list=[0]):
    return pd.read_excel(
        read_file, sheet_name=sheet_name,header=header_list)

def summary_floor_rebar(floor_list:list[Floor],item_type = ''):
    df = pd.DataFrame(columns=['#3','#4','#5','#6','#7','#8','#10','#11'],index=[])
    concrete_df = pd.DataFrame(columns=[],index=[])
    coupler_df = pd.DataFrame(columns=[],index=[])
    formwork_df = pd.DataFrame(columns=[],index=[])
    if item_type == 'column':
        for floor in floor_list:
            list(map(lambda c:c.calculate_rebar() ,floor.column_list))
            floor.summary_rebar()
            new_row = pd.DataFrame(floor.rebar_count,index=[floor.floor_name])
            new_row_concrete = pd.DataFrame(floor.concrete_count,index=[floor.floor_name])
            new_row_coupler = pd.DataFrame(floor.coupler,index=[floor.floor_name])
            new_row_formwork = pd.DataFrame({'模板':floor.formwork_count},index=[floor.floor_name])

            df = pd.concat([df, new_row], verify_integrity=True)
            concrete_df = pd.concat([concrete_df,new_row_concrete],verify_integrity=True)
            coupler_df = pd.concat([coupler_df,new_row_coupler],verify_integrity=True)
            formwork_df = pd.concat([formwork_df,new_row_formwork],verify_integrity=True)
        try:
            coupler_df.loc['Sum'] = coupler_df.sum()
        except:
            pass
    if item_type == 'beam':
        for floor in floor_list:
            floor.summary_beam()
            new_row = pd.DataFrame(floor.beam_rebar_count,index=[floor.floor_name])
            new_row_concrete = pd.DataFrame(floor.concrete_count,index=[floor.floor_name])
            new_row_formwork = pd.DataFrame({'模板':floor.formwork_count},index=[floor.floor_name])

            df = pd.concat([df, new_row], verify_integrity=True)
            concrete_df = pd.concat([concrete_df,new_row_concrete],verify_integrity=True)
            formwork_df = pd.concat([formwork_df,new_row_formwork],verify_integrity=True)
    df.fillna(value=0,inplace=True)
    df.loc['Sum'] = df.sum()
    try:
        concrete_df.loc['Sum'] = concrete_df.sum()
        formwork_df.loc['Sum'] = formwork_df.sum()
    except:
        pass
    return df,concrete_df,coupler_df,formwork_df
def summary_floor_rebar_ratio(floor_list:list[Floor]):
    # df = pd.DataFrame(columns=["0-0.5%","0.5%-1.0%",'1.0%-1.5%','1.5%-2.0%','2.0%-2.5%','2.5%-'],index=[])
    def def_value():
        return [
            [],[],[],
            [],[],[]
            ]
    def def_value_count():
        return defaultdict(
            lambda:[0,0,0,
                    0,0,0])
    # {
    #     'floor':[
    #         [],[],[],
    #         [],[],[]
    #         ]
    # }

    # {
    #     'ratio':{
    #         'floor':[
    #             0,0,0,
    #             0,0,0
    #         ]
    #     }
    # }
    pos = {
        0:'左',
        1:'中',
        2:'右'
    }

    ratio_interval_group = list(arange(0.005,0.03,0.005))
    temp_dict = defaultdict(def_value)
    floor_dict = defaultdict(def_value_count)
    header_list = list(map(lambda r:f'< {r*100}%',ratio_interval_group))
    header_list.append(f'>= {ratio_interval_group[-1]*100}%')
    
    for floor in floor_list:
        for beam in floor.beam_list:
            for i,ratio in enumerate(beam.get_rebar_ratio()):
                for j,ratio_interval in enumerate(ratio_interval_group):
                    if ratio >= ratio_interval_group[-1]:
                        floor_dict[floor.floor_name][header_list[j]][i] += 1
                        break
                    if ratio < ratio_interval:
                        floor_dict[floor.floor_name][header_list[j]][i] += 1
                        break
                temp_dict[floor.floor_name][i].append(ratio)
    row = 0
    df_header_list = []
    df_header_list.insert(0,('樓層',''))
    df_header_list.insert(1,('位置',''))
    for header in header_list:
        df_header_list.append((header,'左'))
        df_header_list.append((header,'中'))
        df_header_list.append((header,'右'))

    df_header_list = pd.MultiIndex.from_tuples(df_header_list)
    ratio_beam = pd.DataFrame(empty([len(floor_list)*2,len(df_header_list)],dtype='<U16'),columns=df_header_list)
    
    for floor,ratio_dict in floor_dict.items():
        ratio_beam.at[row,('樓層','')] = floor
        ratio_beam.at[row + 1,('樓層','')] = floor
        ratio_beam.at[row,('位置','')] = '上'
        ratio_beam.at[row + 1,('位置','')] = '下'
        for ratio,count_list in ratio_dict.items():
            for i,count in enumerate(count_list[:3]):
                ratio_beam.at[row,(ratio,pos[i])] = count
            for i,count in enumerate(count_list[3:]):
                ratio_beam.at[row + 1,(ratio,pos[i])] = count
        row += 2
    return ratio_beam