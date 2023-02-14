from __future__ import annotations
from typing import Tuple
import re
import pandas as pd
from rebar import RebarInfo,RebarArea
from beam import Point,Rebar
from math import pow,sqrt
import copy
class Column:
    height = 0
    size = ''
    serial = ''
    floor = ''
    # rebar_text = ''
    # rebar_text_coor =''
    rebar_coor:list[Point]
    rebar:list[Rebar]
    tie:list[Tie]
    confine_tie:Tie
    middle_tie:Tie
    grid_coor:dict[str,Point]
    tie_list:list[Tuple[Point,Point]]
    tie_text_list:list[Tuple[Point,str]]
    tie_dict:dict[str,Tuple[Point,str]]
    up_column:Column|None
    bot_column:Column|None
    rebar_count:dict[str,float] #以樓層作為區分
    multi_floor:list[str]
    coupler:dict[Tuple[str,str],int]
    floor_object:Floor
    total_rebar:list[Tuple[Rebar,str]]
    total_mass:float
    total_As:float
    fc:int
    fy:int
    concrete:float
    formwork:float
    x_dict:dict[str,float]
    y_dict:dict[str,float]
    def __init__(self):
        # self.rebar_text = ''
        self.fc = 0
        self.fy = 0
        self.x_row = set()
        self.y_row = set()
        self.x_tie = 0
        self.y_tie = 0
        self.concrete = 0
        self.formwork = 0
        self.total_As = 0
        self.total_rebar = []
        self.rebar = []
        self.tie = []
        self.coupler = {}
        self.grid_coor = {}
        self.x_dict = {}
        self.y_dict = {}
        self.rebar_coor = []
        self.tie_list = []
        self.multi_floor = []
        self.multi_column = []
        self.multi_rebar_text = []
        self.tie_text_list = []
        self.tie_dict = {}
        self.up_column = None
        self.bot_column = None
        self.confine_tie = None
        self.middle_tie = None
        self.rebar_count ={}
        self.floor_object = None
    def set_border(self,list1:list,list2:list):
        left_bot = Point((list1[0],list2[2]))
        left_top = Point((list1[0],list2[3]))
        right_top = Point((list1[1],list2[3]))
        right_bot = Point((list1[1],list2[2]))
        self.grid_coor.update({'left_bot': left_bot, 'left_top': left_top, 'right_top': right_top,'right_bot':right_bot})
        # self.grid_coor.extend([left_bot,left_top,right_top,right_bot])
        pass
    def in_grid(self,coor:tuple):
        pt_x = coor[0]
        pt_y = coor[1]
        if len(self.grid_coor) == 0:return False
        if (pt_x - self.grid_coor['left_bot'].x)*(pt_x - self.grid_coor['right_top'].x)<0 and (pt_y - self.grid_coor['right_top'].y)*(pt_y - self.grid_coor['left_bot'].y)<0:
            return True
        return False
    def set_size(self,size):
        self.size = size
        match_obj = re.search(r'([\d|.]+)[x|X]([\d|.]+)',size)
        if match_obj:
            self.x_size = float(match_obj.group(1))
            self.y_size = float(match_obj.group(2))
    def set_prop(self,floor:Floor):
        self.height = floor.height
        self.fc = floor.material_list['fc']
        self.fy = floor.material_list['fy']
    def add_rebar_coor(self,coor):
        self.rebar_coor.append(coor)
    def add_tie(self,coor):
        pt_1 = Point(coor[0])
        pt_2 = Point(coor[1])
        self.tie_list.append((pt_1,pt_2))
    def add_tie_text(self,coor,text):
        pt = Point(((coor[0][0] + coor[1][0])/2,(coor[0][1] + coor[1][1])/2))
        self.tie_text_list.append((pt,text))
    def sort_rebar(self):
        # self.total_rebar = Rebar(self.rebar_text)
        if not self.rebar_coor:return
        if not self.multi_rebar_text:return
        # if self.rebar_text_coor =='':return
        for rebar_coor,rebar_text in self.multi_rebar_text:
            target_rebar = min(self.rebar_coor,key=lambda r:abs(r[0][0]-rebar_coor[0])+abs(r[0][1]-rebar_coor[1]))
            self.rebar_coor.remove(target_rebar)
            self.total_rebar.append((Rebar(rebar_text),target_rebar[1]))
        self.total_As = sum([r[0].As for r in self.total_rebar])
        self.total_mass = sum([r[0].mass for r in self.total_rebar])

        self.x_row = set(map(lambda r:(r[0][0],r[1]),self.rebar_coor))
        self.y_row = set(map(lambda r:(r[0][1],r[1]),self.rebar_coor))
        for total_rebar in self.total_rebar:
            self.x_dict.update({total_rebar[0].size:len([x for x in self.x_row if x[1] == total_rebar[1]])})
            self.y_dict.update({total_rebar[0].size:len([y for y in self.y_row if y[1] == total_rebar[1]])})
            if len(self.total_rebar) > 1:
                min_y = min([r for r in self.rebar_coor if r[1] == total_rebar[1]],key=lambda x:x[1])[0][1]
                min_x = min([r for r in self.rebar_coor if r[1] == total_rebar[1]],key=lambda x:x[0])[0][0]
                self.x_dict[total_rebar[0].size] = len([r for r in self.rebar_coor if r[1] == total_rebar[1] and r[0][1] == min_y])
                self.y_dict[total_rebar[0].size] = len([r for r in self.rebar_coor if r[1] == total_rebar[1] and r[0][0] == min_x])
    def cal_length(self,coor1:Point,coor2:Point):
        return sqrt(pow((coor1.x - coor2.x),2) + pow((coor1.y - coor2.y),2))
    def sort_tie(self):
        if not self.tie_text_list:return
        temp_list = []
        for tie_text in self.tie_text_list:
            match_str = re.findall(r'(#\d+@\d+)',tie_text[1])
            if match_str:
                temp_list.append((tie_text[0],match_str[0]))
        for coor,tie_text in self.tie_text_list:
            if '接頭' in tie_text:
                joint_text = min(temp_list,key = lambda r:self.cal_length(coor,r[0]))
                self.tie_dict.update({'接頭':joint_text})
            if '端部' in tie_text or '圍束' in tie_text or 'COF' in tie_text:
                confine_text = min(temp_list,key = lambda r:self.cal_length(coor,r[0]))
                self.tie_dict.update({'端部':confine_text})
                self.confine_tie = Tie(self.tie_dict['端部'][1],0)
            if '中央' in tie_text or 'TIE' in tie_text:
                middle_text = min(temp_list,key = lambda r:self.cal_length(coor,r[0]))
                self.tie_dict.update({'中央':middle_text})
                self.middle_tie = Tie(self.tie_dict['中央'][1],0)
        if not self.tie_list:return
        temp_list = []
        outer_tie = max(self.tie_list,key=lambda tie:self.cal_length(tie[0],tie[1]))
        self.tie_list.remove(outer_tie)
        for tie in self.tie_list:
            x_diff = abs(tie[0].x - tie[1].x)
            y_diff = abs(tie[0].y - tie[1].y)
            if x_diff >= y_diff:
                self.x_tie += 1
            else:
                self.y_tie += 1
    def set_seq(self,floor_seq:list[str]):
        if not self.floor in floor_seq: 
            self.seq = -1000
            return
        self.seq = floor_seq.index(self.floor)
    def cal_rebar(self):
        copy_up_rebar = copy.deepcopy(self.total_rebar)
        copy_bot_rebar = copy.deepcopy(self.total_rebar)
        if self.up_column:
            # up_rebar = self.up_column.total_rebar
            if self.up_column.total_As > self.total_As:
                copy_up_rebar = copy.deepcopy(self.up_column.total_rebar)
        if self.bot_column:
            # bot_rebar = self.bot_column.total_rebar
            if self.bot_column.total_As > self.total_As:
                copy_bot_rebar = copy.deepcopy(self.bot_column.total_rebar)
        for rebar,text in copy_up_rebar:
            rebar.length = self.height/2
            self.rebar.append(rebar)
        for rebar,text in copy_bot_rebar:
            rebar.length = self.height/2
            self.rebar.append(rebar)
        self.cal_coupler(copy_up_rebar,copy_bot_rebar)  
        # copy_up_rebar.length = self.height/2
        # copy_bot_rebar.length = self.height/2
        # self.rebar.append(copy_up_rebar)
        # self.rebar.append(copy_bot_rebar)
        # self.coupler.update({(copy_up_rebar.size,copy_bot_rebar.size):min(copy_up_rebar.number,copy_bot_rebar.number)})
    def cal_coupler(self,up_rebar:list[Tuple[Rebar,str]],bot_rebar:list[Tuple[Rebar,str]]):
        #狀況1:上32-#8/8-#10 下24-#8/16-#10
        #狀況2:上32-#8/8-#10 下32-#8/16-#10
        #狀況3:上32-#8/8-#10 下36-#8/16-#10
        #狀況4:上32-#8/16-#10 下40-#8/8-#10
        temp_dict = {'up':{},'bot':{}}
        for rebar,text in bot_rebar:
            same_size_rebar = [r[0] for r in up_rebar if r[0].size == rebar.size]
            if same_size_rebar:
                if rebar.number >= same_size_rebar[0].number:
                    self.coupler.update({(same_size_rebar[0].size,rebar.size):same_size_rebar[0].number})
                    temp_dict['bot'].update({rebar.size:rebar.number - same_size_rebar[0].number})
                else:
                    self.coupler.update({(same_size_rebar[0].size,rebar.size):rebar.number})
                    temp_dict['up'].update({rebar.size:same_size_rebar[0].number - rebar.number})
            else:
                temp_dict['bot'].update({rebar.size:rebar.number})
        for size,number in temp_dict['up'].items():
            for bot_size,bot_number in temp_dict['bot'].items():
                if number > 0 and bot_number > 0:
                    self.coupler.update({(size,bot_size):min(number,bot_number)})
                    number -= number - min(number,bot_number)
                    bot_number -= min(number,bot_number)
                    temp_dict['bot'][bot_size] -= min(number,bot_number)

    def cal_tie(self):
        if '端部' in self.tie_dict:
            self.tie.append(Tie(self.tie_dict['端部'][1],(1/6)*self.height*2))
            
        if '中央' in self.tie_dict:
            self.tie.append(Tie(self.tie_dict['中央'][1],(4/6)*self.height))
        
        if self.floor_object.overlap_option["tight_tie"] == '是':
            for tie in self.tie:
                tie.change_spacing(10)
        
    def cal_material(self):
        if self.floor:
            self.concrete = self.x_size * self.y_size * self.floor_object.height
            self.formwork = (self.x_size + self.y_size) * 2 * self.floor_object.height
    def summary_count(self):
        for rebar in self.rebar:
            if not rebar.size in self.rebar_count : self.rebar_count[rebar.size] = 0 
            self.rebar_count[rebar.size] +=rebar.length * rebar.mass
        for tie in self.tie:
            if not tie.size in self.rebar_count : self.rebar_count[tie.size] = 0 
            self.rebar_count[tie.size] += tie.number * RebarInfo(tie.size) * ((self.x_tie + 2) * (self.x_size - 8) + (self.y_tie + 2) * (self.y_size - 8))
    def calculate_rebar(self):
        print(f'calculate map {self.floor} {self.serial}')
        self.cal_rebar()
        self.cal_tie()
        self.cal_material()
        self.summary_count()
        pass        
class Floor:
    height:float
    material_list:dict[str,float]
    column_list:list[Column]
    overlap_option:dict[str,str]
    rebar_count:dict[str,float]
    concrete_count:dict[str,float]
    formwork_count:float
    coupler:dict[str,float]
    floor_name:str
    def __init__(self,floor_name):
        if floor_name[-1] != 'F':
            floor_name += 'F'
        self.floor_name = floor_name
        self.rebar_count ={}
        self.column_list = []
        self.material_list = {}
        self.overlap_option ={}
        self.concrete_count ={}
        self.coupler = {}
        self.formwork_count = 0
        pass
    def set_prop(self,kwargs):
        self.material_list.update({'fc':kwargs["混凝土強度fc'(kgf/cm2)"]})
        self.material_list.update({'fy':kwargs["鋼筋強度fy(kgf/cm2)"]})
        self.overlap_option.update({"tight_tie":kwargs["全段緊密"],"coupler":kwargs["續接器"],"overlap":kwargs["續接方式"]})
        self.height = float(kwargs["樓高"])
    def add_column(self,c_list:list[Column]):
        if not c_list:return
        for c in c_list:
            c.set_prop(self)
            c.floor_object = self
        self.column_list.extend(c_list)
    def summary_rebar(self):
        for c in self.column_list:
            for size,count in c.rebar_count.items():
                if not size in self.rebar_count : self.rebar_count[size] = 0
                self.rebar_count[size] += round(count/1000/1000,2)
            for size,coupler in c.coupler.items():
                if size == ('',''):continue
                if not size in self.coupler : self.coupler[size] = 0
                self.coupler[size] += coupler
            if not c.fc in self.concrete_count:self.concrete_count[c.fc] = 0
            self.concrete_count[c.fc] += c.concrete
            self.formwork_count += c.formwork
        self.rebar_count['total'] = sum(self.rebar_count.values())
        
class Rebar:
    length = 0
    text = 0
    number=0
    size = ''
    As = 0
    mass = 0
    def __init__(self,rebar_text:str):
        self.text = rebar_text
        match_obj = re.search(r'(\d+).([#|D]\d+)',self.text)
        if match_obj:
            self.number = float(match_obj.group(1))
            self.size = match_obj.group(2)
            self.mass = self.number * RebarInfo(self.size)
            self.As = self.number * RebarArea(self.size)
class Tie:
    length = 0
    size = ''
    text = ''
    spacing = 0
    number = 0
    Ash = 0
    def __init__(self,tie_text:str,length:float):
        self.text = tie_text
        self.length = length
        match_obj = re.search(r'([#|D]\d+)[@](\d+)',self.text)
        if match_obj:
            self.spacing = float(match_obj.group(2))
            self.size = match_obj.group(1)
            self.number = length//self.spacing
            self.Ash = RebarArea(self.size)
    def change_spacing(self,new_spacing:float):
        self.spacing = new_spacing
        self.number = self.length//self.spacing
if __name__ == '__main__':
    c = Column()
    c.grid_coor={'left_bot':Point(12685.43,4833.29),'left_top':(12685.43,5146.29),'right_top':Point(12835.43,5146.29),'right_bot':Point(12835.43,4833.29)}
    print(c.in_grid((12717.91,4832.76)))