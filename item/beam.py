from __future__ import annotations
import re
import pandas as pd
from collections import defaultdict
from item.excepteions import BeamFloorNameError
from plan_to_beam import turn_floor_to_float,turn_floor_to_string
from typing import Tuple
from item.rebar import RebarInfo,RebarArea,RebarFy,RebarDiameter
from item import floor
from item.point import Point
from enum import Enum
commom_pattern = r'(,)|(、)'
stash_pattern = r'(\w+)[-|~](\w+)'
class Rebar:
    
    start_pt = Point
    end_pt = Point
    length = 0
    text:str
    number=0
    size = ''
    fy = 0
    As = 0
    def __init__(self,start_pt,end_pt,length,number,size,text,add_up=''):
        self.start_pt = Point(start_pt)
        self.end_pt = Point(end_pt)
        self.number = int(number)
        self.size = size
        self.length = length
        self.text = text
        self.start_pt.x -= self.length/2
        self.end_pt.x += self.length/2
        self.As = RebarArea(self.size) * self.number
        self.fy = RebarFy(self.size)
    def __str__(self) -> str:
        return self.text
    def __repr__(self) -> str:
        return self.text
class RebarType(Enum):
    Top = 'top'
    Bottom = 'bottom'
    Left = 'left'
    Middle = 'middle'
    Right = 'right'
class BeamType(Enum):
    FB = 'fbeam'
    Grider = 'beam'
    SB = 'sbeam'
    Other = 'other'
class Tie:
    # start_pt=Point
    count = 0
    tie_num = 0
    size = ''
    text = 0
    spacing = 0
    Ash = 0
    fy = 0
    def __init__(self,tie,coor,tie_num,count,size):
        self.start_pt = Point(coor)
        self.count = count
        self.size = size
        self.text = tie
        self.tie_num = tie_num
        self.Ash = RebarArea(self.size) * 2
        self.fy = RebarFy(self.size)
        match_obj = re.search(r'(\d*)([#|D]\d+)[@](\d+)',self.text)
        if match_obj:
            self.spacing = float(match_obj.group(3))
            if match_obj.group(1):
                self.Ash *= 2
        # self.spacing = float()
class Beam:
    
    middle_tie:list[Rebar]
    rebar_list:list[Rebar]
    rebar_add_list:list[Rebar] #line with no arrow
    rebar_bend_list:list[Rebar]
    tie_list:list[Tie]
    rebar:dict[str,list[Rebar]]
    rebar_table:dict[str,dict[str,list[Rebar]]]
    tie:dict[str,Tie]
    rebar_count:dict[str,float]
    tie_count:dict[str,float]
    floor_object:floor.Floor
    multi_floor:list[str]
    rebar_ratio:dict[Tuple[RebarType,RebarType],float]
    serial = ''
    floor = ''
    depth = 0
    width = 0
    length = 0
    left_column = 0
    right_column = 0
    concrete = 0
    formwork = 0
    start_pt:Point
    end_pt:Point
    beam_type:BeamType
    ng_message:list[str]
    protect_layer:int
    # coor = Point
    # bounding_box = (Point,Point)
    def __init__(self,serial,x,y):
        self.beam_type = BeamType.Other
        self.coor = Point()
        self.bounding_box = (Point(),Point())
        self.start_pt = Point()
        self.end_pt = Point()
        self.rebar_list=[]
        self.rebar_add_list= []
        self.rebar_bend_list = []
        self.tie_list = []
        self.middle_tie = []
        self.rebar_count = {}
        self.tie_count = {}
        self.ng_message = []
        self.multi_floor = []
        self.protect_layer = 9
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
        self.rebar_table={
            'top':{
                'left':[],
                'middle':[],
                'right':[]
            },
            'bottom':{
                'left':[],
                'middle':[],
                'right':[]
            },
            'top_length':{
                'left':[],
                'middle':[],
                'right':[]
            },
            'bottom_length':{
                'left':[],
                'middle':[],
                'right':[]
            }
        }
        self.rebar_ratio =defaultdict(lambda:0)
        self.serial = serial
        self.coor.x = x
        self.coor.y = y
        self.get_beam_info()

    def add_rebar(self,**kwargs):
        if 'add_up' in kwargs:
            if kwargs['add_up'] == 'bend':
                self.rebar_bend_list.append(Rebar(**kwargs))
            else:
                self.rebar_add_list.append(Rebar(**kwargs))
            return
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
        floor_serial_spacing_char = ' '
        def _get_floor_list(floor1:float,floor2:float):
            if floor1 >= floor2:
                l = list(range(int(floor1),int(floor2),-1))
                l.append(floor2)
                return l
            else:
                l = list(range(int(floor1),int(floor2),1))
                l.append(floor2)
                return l
        ## get beam floor
        if floor_serial_spacing_char not in self.serial:
            temp_matchobj = re.search(r'\((.*)\)(.*\(.*\))',self.serial)
            if temp_matchobj:
                self.floor = temp_matchobj.group(1)
                self.serial = temp_matchobj.group(2)
            else:
                raise BeamFloorNameError
        else:
            self.floor = self.serial.split(' ')[0]
            if self.floor == '':
                raise BeamFloorNameError
            self.serial = ''.join(self.serial.split(floor_serial_spacing_char)[1:])
        if re.search(commom_pattern,self.floor):
            sep = re.search(commom_pattern,self.floor).group(0)
            for floor_text in self.floor.split(sep):
                self.multi_floor.append(floor_text)
            self.floor = self.multi_floor[0]
        if re.search(stash_pattern,self.floor):
            try:
                floor_tuple = re.findall(stash_pattern ,self.floor)
                for floors in floor_tuple:
                    first_floor = turn_floor_to_float(floor=floors[0])
                    second_floor = turn_floor_to_float(floor=floors[-1])
                    if first_floor and second_floor and max(first_floor,second_floor) < 100:
                        for floor_float in _get_floor_list(second_floor,first_floor):
                            self.multi_floor.append(turn_floor_to_string(floor_float))
                        self.floor = self.multi_floor[0]
            except:
                pass
        if self.floor[-1] != 'F':
            self.floor += 'F'

        ## get beam width/depth
        # temp_serial = ''.join(self.serial.split(floor_serial_spacing_char)[1:])
        # self.serial = ''.join(self.serial.split(floor_serial_spacing_char)[1:])
        temp_serial = self.serial
        matches= re.findall(r"\((.*?)\)",self.serial,re.MULTILINE)
        if len(matches) == 0 or len(re.findall(r"X|x",matches[0],re.MULTILINE))==0:return
        split_char = re.findall(r"X|x",matches[0])[0]
        try:
            self.depth = int(matches[0].split(split_char)[1])
            self.width = int(matches[0].split(split_char)[0])
        except:
            self.depth = 0
            self.width = 0
        
        ## get beam serial/type
        match_obj = re.search(r'(.+)\((.*?)\)',temp_serial)
        if match_obj:
            serial = match_obj.group(1).replace(" ","")
            self.beam_type = BeamType.Other
            if re.search(r'^[B|G]',serial):
                self.beam_type = BeamType.Grider
            if re.search(r'^F',serial):
                self.beam_type = BeamType.FB
            if re.search(r'^b',serial):
                self.beam_type = BeamType.SB
            self.serial = serial 
    def get_loading(self,band_width):
        return self.floor_object.loading['SDL']* 0.1 * band_width + self.floor_object.loading['LL'] * 0.1* band_width + self.width * self.depth * 2.4 /1000 # t/m

    def sort_beam_rebar(self):
        min_diff = 30
        if not self.rebar_list:return
        self.start_pt.x = min(self.rebar_list,key=lambda rebar:rebar.start_pt.x).start_pt.x
        self.end_pt.x = max(self.rebar_list,key=lambda rebar:rebar.end_pt.x).end_pt.x
        if self.end_pt.x - self.bounding_box[1].x > min_diff:
            self.end_pt.x = min(self.rebar_list,key=lambda rebar:abs(rebar.end_pt.x - self.bounding_box[1].x)).end_pt.x
        self.length = abs(self.start_pt.x - self.end_pt.x)
        self.rebar_list.sort(key=lambda rebar:(rebar.start_pt.y,rebar.start_pt.x))
        
        top_y = self.rebar_list[-1].start_pt.y
        bot_y = self.rebar_list[0].start_pt.y
        for rebar in self.rebar_list:
            if rebar.end_pt.x > self.end_pt.x:
                rebar.end_pt.x = self.end_pt.x
                rebar.length -= abs(rebar.end_pt.x - self.end_pt.x)
            if bot_y == rebar.start_pt.y:
                self.rebar['bot_first'].append(rebar)
            elif top_y == rebar.start_pt.y:
                self.rebar['top_first'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) < self.depth/2:
                self.rebar['bot_second'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) >= self.depth/2:
                self.rebar['top_second'].append(rebar)

        for pos,rebar in self.rebar.items():
            # if 'second' in pos:
            #     if len(rebar):
            #         left_rebar = min(rebar,key=lambda r:r.start_pt.x)
            #         while left_rebar.start_pt.x > self.start_pt.x:
            #     continue
            if len(rebar) == 0: continue
            left_rebar = min(rebar,key=lambda r:r.start_pt.x)
            while left_rebar.start_pt.x > self.start_pt.x:
                connect_rebar = [r for r in self.rebar_add_list if abs(r.end_pt.x - left_rebar.start_pt.x) < 0.1 and r.start_pt.y == left_rebar.start_pt.y]
                if connect_rebar:
                    rebar.append(connect_rebar[0])
                    self.rebar_add_list.remove(connect_rebar[0])
                    left_rebar = min(rebar,key=lambda r:r.start_pt.x)
                else:
                    print(f'{self.floor}{self.serial}left rebar error')
                    break
            right_rebar = max(rebar,key=lambda r:r.end_pt.x)
            while right_rebar.end_pt.x < self.end_pt.x:
                connect_rebar = [r for r in self.rebar_add_list if abs(r.start_pt.x - right_rebar.end_pt.x) < 0.1 and r.start_pt.y == right_rebar.end_pt.y]
                if connect_rebar:
                    rebar.append(connect_rebar[0])
                    self.rebar_add_list.remove(connect_rebar[0])
                    right_rebar = max(rebar,key=lambda r:r.end_pt.x)
                else:
                    print(f'{self.floor}{self.serial}right rebar error')
                    break
            rebar.sort(key=lambda r:(r.start_pt.y,r.start_pt.x))
            for i in range(0,len(rebar)-1):
                if abs(rebar[i].end_pt.x - rebar[i+1].start_pt.x) > 50:
                    connect_rebar = [r for r in self.rebar_add_list if abs(r.start_pt.x - rebar[i].end_pt.x) < 0.1 and r.start_pt.y == rebar[i].end_pt.y]
                    if connect_rebar:
                        rebar.insert(i+1,connect_rebar[0])
                    else:
                        print(f'{self.serial}')
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
    def cal_rebar(self):
        ## if E.F in rebar_add_list , will show on floor rebar size
        for rebar_list in [self.rebar_list,self.rebar_add_list,self.rebar_bend_list]:
            for rebar in rebar_list:
                if rebar.size in self.rebar_count:
                    self.rebar_count[rebar.size] += rebar.length * rebar.number * RebarInfo(rebar.size)
                else:
                    self.rebar_count[rebar.size] = rebar.length * rebar.number * RebarInfo(rebar.size)
        for rebar in self.middle_tie:
            matchObj = re.search(r'[#|D]\d+',rebar.text)
            if matchObj:
                size = matchObj.group()
                if not size in self.rebar_count:
                    self.rebar_count[size] = 0
                if "E.F" in rebar.size:
                    self.rebar_count[size] += rebar.length * rebar.number * RebarInfo(size)*2
                else:
                    self.rebar_count[size] += rebar.length * rebar.number * RebarInfo(size)
                break #middle tie rebar number equal to rebar line number, so only count one middle tie
            pass
        for tie in self.tie_list:
            if tie.size in self.tie_count:
                self.tie_count[tie.size] += tie.count * RebarInfo(tie.size) * (self.depth - 10 + self.width - 10) * 2
            else:
                self.tie_count[tie.size] = tie.count * RebarInfo(tie.size) * (self.depth - 10 + self.width - 10) * 2
        self.concrete = (self.depth - 15)*self.width*self.length 
        self.formwork = (self.width + (self.depth - 15)*2)*self.length 
    def get_rebar_weight(self):
        temp = 0
        for size,rebar in self.rebar_count.items():
            temp += rebar
        return temp
    def get_tie_weight(self):
        temp = 0
        for size,rebar in self.tie_count.items():
            temp += rebar
        return temp
    def get_rebar_list(self):
        temp = []
        for rebar in self.rebar_list:
            temp.append(f'{rebar.text}:{rebar.length}')
        return temp
    def get_tie_list(self):
        temp = []
        for rebar in self.tie_list:
            temp.append(rebar.text)
        return temp
    def get_middle_tie(self):
        if(self.middle_tie):return self.middle_tie[0].text
        return 
    def get_concrete(self):
        return self.concrete
    def get_formwork(self):
        return self.formwork
    def write_beam(self,df:pd.DataFrame):
        pass
    def set_prop(self,floor:floor.Floor):
        self.height = floor.height
        self.fc = floor.material_list['fc']
        self.fy = floor.material_list['fy']
        self.floor_object = floor
    ##取得梁配筋面積
    def get_rebar_table(self,rebar_type1:RebarType,rebar_type2:RebarType) -> float:
        As = 0
        for rebar in self.rebar_table[rebar_type1.value][rebar_type2.value]:
            As += rebar.As
        return As
    ## 整理梁配筋成常用表格
    def sort_rebar_table(self):
        min_diff = 10
        for rebar in self.rebar['top_first']:
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff :
                self.rebar_table['top']['left'].append(rebar)
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['top']['right'].append(rebar)
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
                self.rebar_table['top']['middle'].append(rebar)
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff:
                self.rebar_table['top_length']['left'].append(rebar.length)
                continue
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['top_length']['right'].append(rebar.length)
                continue
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff):
                self.rebar_table['top_length']['middle'].append(rebar.length)
                continue
        for rebar in self.rebar['top_second']:
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff :
                self.rebar_table['top']['left'].append(rebar)
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['top']['right'].append(rebar)
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
                self.rebar_table['top']['middle'].append(rebar) 

        for rebar in self.rebar['bot_first']:
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff :
                self.rebar_table['bottom']['left'].append(rebar)
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['bottom']['right'].append(rebar)
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
                self.rebar_table['bottom']['middle'].append(rebar)
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff:
                self.rebar_table['bottom_length']['left'].append(rebar.length)
                continue
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['bottom_length']['right'].append(rebar.length)
                continue
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff):
                self.rebar_table['bottom_length']['middle'].append(rebar.length)
                continue
        for rebar in self.rebar['bot_second']:
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff :
                self.rebar_table['bottom']['left'].append(rebar)
            if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
                self.rebar_table['bottom']['right'].append(rebar)
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
                self.rebar_table['bottom']['middle'].append(rebar)
        if len(self.rebar_table['top']['middle']) == 0:
            if self.rebar_table['top_length']['left'] > self.rebar_table['top_length']['right']:
                self.rebar_table['top']['middle'].extend(self.rebar_table['top']['left'])
            else:
                self.rebar_table['top']['middle'].extend(self.rebar_table['top']['right'])
        if len(self.rebar_table['bottom']['middle']) == 0:
            if self.rebar_table['bottom_length']['left'] > self.rebar_table['bottom_length']['right']:
                self.rebar_table['bottom']['middle'].extend(self.rebar_table['bottom']['left'])
            else:
                self.rebar_table['bottom']['middle'].extend(self.rebar_table['bottom']['right'])
        self.cal_rebar_ratio()
        self.cal_ld_table()
    ##計算梁配筋比
    def cal_rebar_ratio(self):
        for rebar_type in [RebarType.Top,RebarType.Bottom]:
            for rebar_type2 in [RebarType.Left,RebarType.Middle,RebarType.Right]:
                try:
                    self.rebar_ratio[(rebar_type,rebar_type2)] = self.get_rebar_table(rebar_type1=rebar_type,
                                                                                  rebar_type2=rebar_type2)/(self.width * 
                                                                                                            (self.depth - self.protect_layer))
                except:
                    self.rebar_ratio[(rebar_type,rebar_type2)] = 0    
    def get_rebar_ratio(self):
        return [
                    self.rebar_ratio[(RebarType.Top,RebarType.Left)],
                    self.rebar_ratio[(RebarType.Top,RebarType.Middle)],
                    self.rebar_ratio[(RebarType.Top,RebarType.Right)],
                    self.rebar_ratio[(RebarType.Bottom,RebarType.Left)],
                    self.rebar_ratio[(RebarType.Bottom,RebarType.Middle)],
                    self.rebar_ratio[(RebarType.Bottom,RebarType.Right)],
                ]
    def cal_ld(self,rebar:Rebar,tie:Tie):
        from math import sqrt,ceil
        cover = 7.5
        fy = self.fy
        fc = self.fc
        fydb = RebarDiameter(rebar.size)
        fytdb = RebarDiameter(tie.size)
        spacing = tie.spacing
        if self.beam_type == BeamType.Grider:
            spacing = 10
        width_ = self.width
        fynum = rebar.number
        avh = RebarArea(tie.size)

        psitTop_ = 1.3
        psitBot_ = 1
        psie_ = 1
        lamda_ = 1
        psis_ = 1
        ld = fy / (sqrt(fc) * 3.5 * lamda_) * fydb
        if fydb >= 2:
            psis_ = 1
            ld_simple_top = fy * psitTop_ * psie_ / (sqrt(fc) * 5.3 * lamda_) * fydb
            ld_simple_bot = fy * psitBot_ * psie_ / (sqrt(fc) * 5.3 * lamda_) * fydb
        else:
            psis_ = 0.8
            ld_simple_top = fy * psitTop_ * psie_ / (sqrt(fc) * 6.6 * lamda_) * fydb
            ld_simple_bot = fy * psitBot_ * psie_ / (sqrt(fc) * 6.6 * lamda_) * fydb
        
        cs_ = ((width_ - fydb * fynum - fytdb*2 - cover * 2)/(fynum - 1) + fydb) / 2
        cc_ = cover + fytdb + fydb / 2
        if cs_ <= cc_ :
            pass
            cb_ = cs_
            atr_ = 2 * avh
            ktr_ = atr_ * 40 / (spacing * fynum)
        else:
            cb_ = cs_
            atr_ = avh
            ktr_ = atr_ * 40 / (spacing * fynum)
        
        botFactor = psitBot_ * psie_ * psis_ * lamda_ / min((cb_ + ktr_) / fydb, 2.5)
        topFactor = psitTop_ * botFactor

        bot_ld = botFactor * ld
        top_ld = topFactor * ld

        bot_lap_ld = ceil(1.3 * min(ld_simple_bot,bot_ld))
        top_lap_ld = ceil(1.3 * min(ld_simple_top,top_ld))

        return (top_lap_ld,bot_lap_ld)
    
    def cal_ld_table(self):
        self.ld_table = {}
        if self.rebar_table[RebarType.Top.value][RebarType.Left.value] and self.tie['left']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Top.value][RebarType.Left.value][0],
                                                tie = self.tie['left'])
            self.ld_table.update({(RebarType.Top,RebarType.Left):top_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Middle.value] and self.tie['middle']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Top.value][RebarType.Middle.value][0],
                                                tie = self.tie['middle'])
            self.ld_table.update({(RebarType.Top,RebarType.Middle):top_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Right.value] and self.tie['right']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Top.value][RebarType.Right.value][0],
                                                tie = self.tie['right'])
            self.ld_table.update({(RebarType.Top,RebarType.Right):top_lap_ld})

        if self.rebar_table[RebarType.Bottom.value][RebarType.Left.value] and self.tie['left']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Bottom.value][RebarType.Left.value][0],
                                                tie = self.tie['left'])
            self.ld_table.update({(RebarType.Bottom,RebarType.Left):bot_lap_ld})
        if self.rebar_table[RebarType.Bottom.value][RebarType.Middle.value] and self.tie['middle']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Bottom.value][RebarType.Middle.value][0],
                                                tie = self.tie['middle'])
            self.ld_table.update({(RebarType.Bottom,RebarType.Middle):bot_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Right.value] and self.tie['right']:
            top_lap_ld,bot_lap_ld = self.cal_ld(rebar = self.rebar_table[RebarType.Bottom.value][RebarType.Right.value][0],
                                                tie = self.tie['right'])
            self.ld_table.update({(RebarType.Bottom,RebarType.Right):bot_lap_ld})

        