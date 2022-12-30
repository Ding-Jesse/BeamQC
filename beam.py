from __future__ import annotations
import re
import pandas as pd
class Point:
    x = 0
    y = 0
    def __init__(self,*pt):
        if len(pt) == 0:
            pass
        elif isinstance(pt[0],tuple):
            self.x = pt[0][0]
            self.y = pt[0][1]
class Rebar:
    start_pt = Point
    end_pt = Point
    length = 0
    text = 0
    number = 0
    size = ''
    def __init__(self,start_pt,end_pt,length,number,size,text):
        self.start_pt = Point(start_pt)
        self.end_pt = Point(end_pt)
        self.number = number
        self.size = size
        self.length = length
        self.text = text
        self.start_pt.x -= self.length/2
        self.end_pt.x += self.length/2

class Tie:
    # start_pt=Point
    count = 0
    tie_num = 0
    size = ''
    text = 0
    def __init__(self,tie,coor,tie_num,count,size):
        self.start_pt = Point(coor)
        self.count = count
        self.size = size
        self.text = tie
        self.tie_num = tie_num
class Beam:
    rebar_list:list[Rebar]
    tie_list:list[Tie]
    rebar:dict[str,list[Rebar]]
    tie:dict[str,Tie]
    serial = ''
    floor = ''
    depth = 0
    width = 0
    length = 0
    start_pt:Point
    end_pt:Point
    # coor = Point
    # bounding_box = (Point,Point)
    def __init__(self,serial,x,y):
        self.coor = Point()
        self.bounding_box = (Point(),Point())
        self.start_pt = Point()
        self.end_pt = Point()
        self.rebar_list=[]
        self.tie_list = []
        self.middle_tie = []
        self.rebar={
            'top_first':[],
            'top_second':[],
            'bot_first':[],
            'bot_second':[],
        }       
        self.tie ={
            'left':None,
            'middle':None,
            'right':None
        }
        # print(f'{serial}-{hex(id(self.coor))}')
        # print(f'{serial}-{hex(id(self.bounding_box[0]))}')
        self.serial = serial
        self.coor.x = x
        self.coor.y = y
        self.get_beam_info()
    def add_rebar(self,**kwargs):
        if 'E.F' in kwargs['text']:
            self.middle_tie.append(Rebar(**kwargs))
        else:
            self.rebar_list.append(Rebar(**kwargs))
    def add_tie(self,*tie):
        self.tie_list.append(Tie(*tie))
    def set_bounding_box(self,pt1_x,pt1_y,pt2_x,pt2_y):
        self.bounding_box[0].x = pt1_x
        self.bounding_box[0].y = pt1_y
        self.bounding_box[1].x = pt2_x
        self.bounding_box[1].y = pt2_y
    def get_bounding_box(self):
        return ((self.bounding_box[0].x,self.bounding_box[0].y),(self.bounding_box[1].x,self.bounding_box[1].y))
    def get_coor(self):
        return (self.coor.x,self.coor.y)
    def get_beam_info(self):
        self.floor = self.serial.split(' ')[0]
        matches= re.findall(r"\((.*?)\)",self.serial,re.MULTILINE)
        if len(matches) == 0 or 'X' not in matches[0]:return
        try:
            self.depth = int(matches[0].split('X')[1])
            self.width = int(matches[0].split('X')[0])
        except:
            self.depth = 0
            self.width = 0
    def sort_beam_rebar(self):

        if not self.rebar_list:return
        self.start_pt.x = min(self.rebar_list,key=lambda rebar:rebar.start_pt.x).start_pt.x
        self.end_pt.x = max(self.rebar_list,key=lambda rebar:rebar.end_pt.x).end_pt.x
        self.length = abs(self.start_pt.x - self.end_pt.x)
        self.rebar_list.sort(key=lambda rebar:(rebar.start_pt.y,rebar.start_pt.x))
        
        top_y = self.rebar_list[-1].start_pt.y
        bot_y = self.rebar_list[0].start_pt.y
        for rebar in self.rebar_list:
            if bot_y == rebar.start_pt.y:
                self.rebar['bot_first'].append(rebar)
            elif top_y == rebar.start_pt.y:
                self.rebar['top_first'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) < self.depth/2:
                self.rebar['bot_second'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) >= self.depth/2:
                self.rebar['top_second'].append(rebar)
    def sort_beam_tie(self):
        if not self.tie_list:return
        self.tie_list.sort(key=lambda tie:tie.start_pt.x)
        self.tie['left'] = self.tie_list[0]
        self.tie['middle'] = self.tie_list[0]
        self.tie['right'] = self.tie_list[0]
        for i,tie in enumerate(self.tie_list):
            if i == 1:
                self.tie['middle'] = tie
            if i == 2:
                self.tie['right'] = tie
    def write_beam(self,df:pd.DataFrame):
        pass