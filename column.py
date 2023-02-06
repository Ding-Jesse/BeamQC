from __future__ import annotations
import re
import pandas as pd
from rebar import RebarInfo
from beam import Point,Rebar
from math import pow,sqrt
class Column:
    size = ''
    serial = ''
    floor = ''
    rebar_text = ''
    rebar_text_coor =''
    rebar_coor:list[Point]
    rebar:list[Rebar]
    grid_coor:dict[str,Point]
    tie_list:list[tuple(Point,Point)]
    tie_text_list:list[tuple(Point,str)]
    tie_dict:dict[str,tuple(Point,str)]
    def __init__(self):
        self.x_row = 0
        self.y_row = 0
        self.x_tie = 0
        self.y_tie = 0
        self.rebar = []
        self.grid_coor = {}
        self.rebar_coor = []
        self.tie_list = []
        self.tie_text_list = []
        self.tie_dict = {}
        pass
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
        match_obj = re.search(r'(\d+)[x|X](\d+)',size)
        if match_obj:
            self.x_size = match_obj.group(1)
            self.y_size = match_obj.group(2)
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
        if not self.rebar_coor:return
        if self.rebar_text_coor =='':return
        self.rebar_coor.remove(min(self.rebar_coor,key=lambda r:abs(r[0]-self.rebar_text_coor[0])+abs(r[1]-self.rebar_text_coor[1])))
        self.x_row = len(set(map(lambda r:r[0],self.rebar_coor)))
        self.y_row = len(set(map(lambda r:r[1],self.rebar_coor)))
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
            if '端部' in tie_text:
                confine_text = min(temp_list,key = lambda r:self.cal_length(coor,r[0]))
                self.tie_dict.update({'端部':confine_text})
            if '中央' in tie_text:
                middle_text = min(temp_list,key = lambda r:self.cal_length(coor,r[0]))
                self.tie_dict.update({'中央':middle_text})
        if not self.tie_list:return
        temp_list = []
        outer_tie = max(self.tie_list,key=lambda tie:self.cal_length(tie[0],tie[1]))
        self.tie_list.remove(outer_tie)
        for tie in self.tie_list:
            x_diff = tie[0].x - tie[1].x
            y_diff = tie[0].y - tie[1].y
            if x_diff >= y_diff:
                self.x_tie += 1
            else:
                self.y_tie += 1
class Floor:
    height:float
    material_list:dict[str,float]
    column_list:list[Column]
    overlap_option:str
    rebar_count:dict[str,float]
    concrete_count:dict[str,float]
    formwork_count:dict[str,float]
    coupler:dict[str,float]
    floor_name:str
    def __init__(self,floor_name):
        self.floor_name =floor_name
        pass
    def set_prop(self,kwargs):
        pass

