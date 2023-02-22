from __future__ import annotations
from item import column,beam
from collections import defaultdict
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
    def __init__(self,floor_name):
        if floor_name[-1] != 'F':
            floor_name += 'F'
        self.floor_name = floor_name
        self.rebar_count ={}
        self.column_list = []
        self.beam_list = []
        self.material_list = {}
        self.overlap_option ={}
        self.concrete_count =defaultdict(lambda:0)
        self.beam_rebar_count = defaultdict(lambda:0)
        self.loading = defaultdict(lambda:0)
        self.coupler = {}
        self.formwork_count = 0
        pass
    def set_prop(self,kwargs):
        self.material_list.update({'fc':kwargs["混凝土強度fc'(kgf/cm2)"]})
        self.material_list.update({'fy':kwargs["鋼筋強度fy(kgf/cm2)"]})
        self.overlap_option.update({"tight_tie":kwargs["全段緊密"],"coupler":kwargs["續接器"],"overlap":kwargs["續接方式"]})
        self.loading.update({"SDL":kwargs["SDL(t/m2)"]})
        self.loading.update({"LL":kwargs["LL(t/m2)"]})
        self.height = float(kwargs["樓高"])
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
                self.coupler[size] += coupler
            # if not c.fc in self.concrete_count:self.concrete_count[c.fc] = 0
            self.concrete_count[c.fc] += c.concrete
            self.formwork_count += c.formwork
        self.rebar_count['total'] = sum(self.rebar_count.values())
    def summary_beam(self):
        for b in self.beam_list:
            for size,count in b.rebar_count.items():
                self.beam_rebar_count[size] = round(count/1000/1000,2)
            for size,count in b.tie_count.items():
                self.beam_rebar_count[size] = round(count/1000/1000,2)
            self.concrete_count[b.fc] +=   b.concrete
            self.formwork_count += b.formwork
        self.beam_rebar_count['total'] = sum(self.rebar_count.values())
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
            new_row_formwork = pd.DataFrame(floor.formwork_count,index=[floor.floor_name])

            df = pd.concat([df, new_row], verify_integrity=True)
            concrete_df = pd.concat([concrete_df,new_row_concrete],verify_integrity=True)
            coupler_df = pd.concat([coupler_df,new_row_coupler],verify_integrity=True)
            formwork_df = pd.concat([formwork_df,new_row_formwork],verify_integrity=True)
        coupler_df.loc['Sum'] = coupler_df.sum()
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
    concrete_df.loc['Sum'] = concrete_df.sum()
    formwork_df.loc['Sum'] = formwork_df.sum()
    return df,concrete_df,coupler_df,formwork_df 