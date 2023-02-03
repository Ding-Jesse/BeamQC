from __future__ import annotations
import re
import pandas as pd
from rebar import RebarInfo
from beam import Point,Rebar
class Column:
    size = ''
    serial = ''
    floor = ''
    rebar_text = ''
    rebar_text_coor =''
    rebar_coor:list[Point]
    rebar:list[Rebar]
    grid_coor:dict[str,Point]
    def __init__(self):
        self.x_row = 0
        self.y_row = 0
        self.rebar = []
        self.grid_coor = {}
        self.rebar_coor = []
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
    def add_rebar_coor(self,coor):
        self.rebar_coor.append(coor)
    def sort_rebar(self):
        if not self.rebar_coor:return
        if self.rebar_text_coor =='':return
        self.rebar_coor.remove(min(self.rebar_coor,key=lambda r:abs(r[0]-self.rebar_text_coor[0])+abs(r[1]-self.rebar_text_coor[1])))
        self.x_row = len(set(map(lambda r:r[0],self.rebar_coor)))
        self.y_row = len(set(map(lambda r:r[1],self.rebar_coor)))
