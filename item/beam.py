from __future__ import annotations
import re
import pandas as pd
from collections import defaultdict
from item.excepteions import BeamFloorNameError
from typing import Tuple
from enum import Enum
from dataclasses import dataclass, field
from item.rebar import RebarInfo, RebarArea, RebarFy, RebarDiameter
from item.floor import Floor
from item.point import Point

commom_pattern = r'(,)|(、)'
stash_pattern = r'(\w+)[-|~](\w+)'


@dataclass(eq=True)
class Rebar:
    arrow_coor: Tuple[Tuple[float, float],
                      Tuple[float, float]] = field(default_factory=Tuple)
    start_pt: Point = None
    end_pt: Point = None
    length: float = 0
    text: str = ''
    number: float = 0
    size: str = ''
    fy: float = 0
    As: float = 0
    arrow_coor: tuple

    def __init__(self, start_pt, end_pt, length, number, size, text, arrow_coor, with_dim, add_up='', measure_type=''):
        self.start_pt = Point(start_pt)
        self.end_pt = Point(end_pt)
        self.number = int(number)
        self.size = size
        self.length = length
        self.text = text
        self.start_pt.x -= self.length/2
        self.end_pt.x += self.length/2
        self.arrow_coor = arrow_coor
        self.As = RebarArea(self.size) * self.number
        self.fy = RebarFy(self.size)
        self.dim = with_dim
        if measure_type == 'mm':
            self.length /= 10

    def __str__(self) -> str:
        return self.text

    def __repr__(self) -> str:
        return self.text

    def set_new_property(self, number, size):
        '''
        set rebar with new size and number
        '''
        self.size = size
        self.number = int(number)
        self.As = RebarArea(self.size) * self.number
        self.fy = RebarFy(self.size)
        self.text = f'{self.number}-{self.size}'


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


@dataclass(eq=True)
class Tie:
    start_pt: Point = None
    count: float = 0
    tie_num: float = 0
    size: str = ''
    text: str = ''
    spacing: float = 0
    Ash: float = 0
    fy: float = 0

    def __init__(self, tie, coor, tie_num, count, size):
        self.start_pt = Point(coor)
        self.count = count
        self.size = size
        self.text = tie
        self.tie_num = tie_num
        self.Ash = RebarArea(self.size) * 2
        self.fy = RebarFy(self.size)
        match_obj = re.search(r'(\d*)([#|D]\d+)[@](\d+)', self.text)
        if match_obj:
            self.spacing = float(match_obj.group(3))
            if match_obj.group(1):
                self.Ash *= 2

    def __str__(self) -> str:
        return self.text

    def __repr__(self) -> str:
        return self.text
        # self.spacing = float()


@dataclass(eq=True)
class Beam():

    middle_tie: list[Rebar] = field(default_factory=list)
    rebar_list: list[Rebar] = field(default_factory=list)
    rebar_add_list: list[Rebar] = field(
        default_factory=list)  # line with no arrow
    rebar_bend_list: list[Rebar] = field(default_factory=list)
    tie_list: list[Tie] = field(default_factory=list)
    rebar: dict[str, list[Rebar]] = field(default_factory=dict)
    rebar_table: dict[str, dict[str, list[Rebar]]
                      ] = field(default_factory=dict)
    tie: dict[str, Tie] = field(default_factory=dict)
    rebar_count: dict[str, float] = field(default_factory=dict)
    tie_count: dict[str, float] = field(default_factory=dict)
    floor_object: floor.Floor = None
    multi_floor: list[str] = field(default_factory=list)
    rebar_ratio: dict[Tuple[RebarType, RebarType],
                      float] = field(default_factory=dict)
    serial: str = ''
    floor: str = ''
    depth: float = 0
    width: float = 0
    length: float = 0
    left_column = 0
    right_column = 0
    concrete = 0
    formwork = 0
    fc: float = 0
    start_pt: Point = None
    end_pt: Point = None
    beam_type: BeamType = None
    ng_message: list[str] = field(default_factory=list)
    protect_layer: float = 0
    plan_count: float = 0
    top_y = 0
    bot_y = 0
    # coor = Point
    # bounding_box = (Point,Point)

    def __init__(self, serial, x, y):
        self.plan_count = 1
        self.detail_report = []
        self.beam_type = BeamType.Other
        self.coor = Point()
        self.bounding_box = (Point(), Point())
        self.start_pt = Point()
        self.end_pt = Point()
        self.rebar_list = []
        self.rebar_add_list = []
        self.rebar_bend_list = []
        self.tie_list = []
        self.middle_tie = []
        self.rebar_count = {}
        self.tie_count = {}
        self.ng_message = []
        self.multi_floor = []
        self.multi_serial = []
        self.protect_layer = 9
        self.rebar = {
            'top_first': [],
            'top_second': [],
            'bot_first': [],
            'bot_second': [],
        }
        self.tie = {
            'left': None,
            'middle': None,
            'right': None
        }
        self.rebar_table = {
            'top': {
                'left': [],
                'middle': [],
                'right': []
            },
            'bottom': {
                'left': [],
                'middle': [],
                'right': []
            },
            'top_length': {
                'left': [],
                'middle': [],
                'right': []
            },
            'bottom_length': {
                'left': [],
                'middle': [],
                'right': []
            }
        }
        self.rebar_ratio = defaultdict(lambda: 0)
        self.serial = serial
        self.coor.x = x
        self.coor.y = y
        self.fy = 0
        self.fc = 0
        self.measure_type = 'cm'
        self.height = 0
        # self.get_beam_info()

    def add_rebar(self, **kwargs):
        kwargs['measure_type'] = self.measure_type
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

    def add_tie(self, *tie):
        self.tie_list.append(Tie(*tie))

    def set_bounding_box(self, pt1_x, pt1_y, pt2_x, pt2_y):
        self.bounding_box[0].x = pt1_x
        self.bounding_box[0].y = pt1_y
        self.bounding_box[1].x = pt2_x
        self.bounding_box[1].y = pt2_y

    def get_bounding_box(self):
        return ((self.bounding_box[0].x, self.bounding_box[0].y), (self.bounding_box[1].x, self.bounding_box[1].y))

    def get_coor(self):
        return (self.coor.x, self.coor.y)

    def get_beam_info(self, floor_list: list[str],
                      measure_type='cm',
                      name_pattern: dict = None,
                      size_pattern: dict = None,
                      floor_pattern: str = ''):
        '''
        - Seperate Floor with pattern like "、|,|~"
        - assign beam name type, seperate floor, name and section type
        - inputs = {
            'Grider':[],
            'FB':[],
            'SB':[]
        }
        '''
        if name_pattern is None:
            return None
        if size_pattern is None:
            return None

        # floor_serial_spacing_char = ' '
        self.measure_type = measure_type

        if name_pattern:
            for beam_type, patterns in name_pattern.items():
                match_floor = ''
                match_serial = ''
                match_obj = None

                for pattern in patterns:
                    match_obj = re.search(pattern, self.serial)
                    if match_obj:
                        match_floor = match_obj.group(1)
                        match_floor = re.sub(r'\(|\)', '', match_floor)  # 去除()
                        match_serial = match_obj.group(2)
                        match_serial.replace(' ', '')  # 去除編號與尺寸的間隔
                        break

                if match_obj:

                    if beam_type == 'Grider':
                        self.beam_type = BeamType.Grider
                    if beam_type == 'FB':
                        self.beam_type = BeamType.FB
                        if match_floor == '':
                            match_floor = floor_list[-1]
                    if beam_type == 'SB':
                        self.beam_type = BeamType.SB

                    self.floor = match_floor
                    self.serial = match_serial

                    break

        # Seperate Floor with pattern like "、,"
        if re.search(commom_pattern, self.floor):
            sep = re.search(commom_pattern, self.floor).group(0)
            for floor_text in self.floor.split(sep):
                if floor_pattern:
                    match_obj = re.search(floor_pattern, floor_text)
                    if match_obj:
                        floor_text = match_obj.group()
                self.multi_floor.append(floor_text)
            self.floor = self.multi_floor[0]
        if re.search(stash_pattern, self.floor):
            try:
                floor_tuple = re.findall(stash_pattern, self.floor)
                for floors in floor_tuple:
                    first_floor = floors[0]
                    second_floor = floors[-1]
                    if floor_pattern:
                        match_obj = re.search(floor_pattern, first_floor)
                        if match_obj:
                            first_floor = match_obj.group()
                        match_obj = re.search(floor_pattern, second_floor)
                        if match_obj:
                            second_floor = match_obj.group()

                    if first_floor[-1] != 'F':
                        first_floor += 'F'

                    if second_floor[-1] != 'F':
                        second_floor += 'F'

                    first_index = min(floor_list.index(
                        first_floor), floor_list.index(second_floor))
                    second_index = max(floor_list.index(
                        first_floor), floor_list.index(second_floor))
                    self.multi_floor.extend(
                        floor_list[first_index:second_index + 1])
                    self.floor = self.multi_floor[0]
            except:
                pass

        if self.floor == '':
            return None

        if self.floor[-1] != 'F':
            self.floor += 'F'

        # regex the floor name to prevent special name
        self.floor = re.search(floor_pattern, self.floor).group()

        # get beam width/depth
        temp_serial = self.serial
        matches_size = re.search(size_pattern['pattern'], self.serial)
        if matches_size:
            text_depth = int(matches_size.group(size_pattern['depth']))
            text_width = int(matches_size.group(size_pattern['width']))
            if measure_type == 'mm':
                text_depth /= 10
                text_width /= 10
        else:
            text_depth = 0
            text_width = 0

        self.depth = text_depth
        self.width = text_width

        # get beam serial/type
        match_obj = re.search(r'(.+)\((.*?)\)', temp_serial)
        if match_obj:
            serial = match_obj.group(1).replace(" ", "")
            self.serial = serial
            if re.search(commom_pattern, self.serial):
                sep = re.search(commom_pattern, self.serial).group(0)
                for serial_text in self.serial.split(sep):
                    self.multi_serial.append(serial_text)
                self.serial = self.multi_serial[0]

        return self

    def get_loading(self, band_width):
        # t/m
        return self.floor_object.loading['SDL'] * 0.1 * band_width + self.floor_object.loading['LL'] * 0.1 * band_width + self.width * self.depth * 2.4 / 1000

    def sort_beam_rebar(self):
        factor = 1
        if self.measure_type == 'mm':
            factor = 10

        def check_rebar_dim(pos_list: list[Rebar], rebar: Rebar):

            if len(pos_list) > 0:
                prev_rebar = pos_list[-1]
            else:
                return
            if rebar.dim and not prev_rebar.dim:
                if rebar.start_pt.x != prev_rebar.end_pt.x:
                    prev_rebar.end_pt.x = rebar.start_pt.x
                    prev_rebar.length = abs(
                        prev_rebar.start_pt.x - prev_rebar.end_pt.x) / factor
            elif not rebar.dim and prev_rebar.dim:
                if rebar.start_pt.x != prev_rebar.end_pt.x:
                    rebar.start_pt.x = prev_rebar.end_pt.x
                    rebar.length = abs(rebar.start_pt.x -
                                       rebar.end_pt.x) / factor
            return

        if not self.rebar_list:
            return

        self.start_pt.x = min(
            self.rebar_list, key=lambda rebar: rebar.start_pt.x).start_pt.x
        self.end_pt.x = max(
            self.rebar_list, key=lambda rebar: rebar.end_pt.x).end_pt.x

        dim_start = self.start_pt.x
        dim_end = self.end_pt.x
        dim_pt = [r for r in self.rebar_list if r.dim]
        if dim_pt:
            dim_start = min(
                dim_pt, key=lambda rebar: rebar.start_pt.x).start_pt.x
            dim_end = max(dim_pt, key=lambda rebar: rebar.end_pt.x).end_pt.x

        if abs(self.start_pt.x - dim_start) < 150:
            self.start_pt.x = dim_start
        if abs(self.end_pt.x - dim_end) < 150:
            self.end_pt.x = dim_end

        # if self.end_pt.x - self.bounding_box[1].x > min_diff:
        #     self.end_pt.x = min(self.rebar_list,key=lambda rebar:abs(rebar.end_pt.x - self.bounding_box[1].x)).end_pt.x
        # if self.start_pt.x - self.bounding_box[0].x > min_diff and self.rebar_add_list:
        #     if min(self.rebar_add_list,key=lambda rebar:abs(rebar.start_pt.x - self.bounding_box[0].x)).start_pt.x < self.start_pt.x :
        #         self.start_pt.x = min(self.rebar_add_list,key=lambda rebar:abs(rebar.start_pt.x - self.bounding_box[0].x)).start_pt.x
        self.length = abs(self.start_pt.x - self.end_pt.x) / factor
        self.rebar_list.sort(key=lambda rebar: (
            round(rebar.arrow_coor[0][1]), round(rebar.arrow_coor[0][0])))

        top_y = self.rebar_list[-1].start_pt.y
        bot_y = self.rebar_list[0].start_pt.y
        for rebar in self.rebar_list:
            if rebar.end_pt.x > self.end_pt.x:
                rebar.end_pt.x = self.end_pt.x
                rebar.length -= abs(rebar.end_pt.x - self.end_pt.x) / factor
            if bot_y == rebar.start_pt.y:
                check_rebar_dim(self.rebar['bot_first'], rebar=rebar)
                self.rebar['bot_first'].append(rebar)
            elif top_y == rebar.start_pt.y:
                check_rebar_dim(self.rebar['top_first'], rebar=rebar)
                self.rebar['top_first'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) < self.depth/2:
                check_rebar_dim(self.rebar['bot_second'], rebar=rebar)
                self.rebar['bot_second'].append(rebar)
            elif abs(rebar.start_pt.y - bot_y) >= self.depth/2:
                check_rebar_dim(self.rebar['top_second'], rebar=rebar)
                self.rebar['top_second'].append(rebar)
        self.top_y = top_y
        self.bot_y = bot_y

    def sort_beam_tie(self):
        if not self.tie_list:
            return
        self.tie_list.sort(key=lambda tie: tie.start_pt.x)
        self.tie['left'] = self.tie_list[0]
        self.tie['middle'] = self.tie_list[0]
        self.tie['right'] = self.tie_list[0]
        for i, tie in enumerate(self.tie_list):
            if i == 1:
                self.tie['middle'] = tie
            if i == 2:
                self.tie['right'] = tie

    def sort_middle_tie(self):
        '''
        for the middle tie type is mutli line assign not one line assign
        '''
        if not self.middle_tie:
            return

        middle_tie = self.middle_tie[0]

        if middle_tie.number != len(self.middle_tie):
            middle_tie.set_new_property(number=len(
                self.middle_tie), size=middle_tie.size)

    def cal_rebar(self):
        # if E.F in rebar_add_list , will show on floor rebar size
        for rebar_list in [self.rebar_list, self.rebar_add_list, self.rebar_bend_list]:
            for rebar in rebar_list:
                if rebar.size in self.rebar_count:
                    self.rebar_count[rebar.size] += rebar.length * \
                        rebar.number * RebarInfo(rebar.size)
                else:
                    self.rebar_count[rebar.size] = rebar.length * \
                        rebar.number * RebarInfo(rebar.size)
                self.detail_report.append(
                    f'主筋:{rebar}= {rebar.length:.2f} (cm) * {rebar.number} * {RebarInfo(rebar.size):.2f} (kg)')
        for rebar in self.middle_tie:
            matchObj = re.search(r'[#|D]\d+', rebar.text)
            if matchObj:
                size = matchObj.group()
                if not size in self.rebar_count:
                    self.rebar_count[size] = 0
                if "E.F" in rebar.size:
                    self.rebar_count[size] += rebar.length * \
                        rebar.number * RebarInfo(size)*2
                else:
                    self.rebar_count[size] += rebar.length * \
                        rebar.number * RebarInfo(size)
                self.detail_report.append(
                    f'側筋:{rebar}= {rebar.length:.2f} (cm) * {rebar.number} * {RebarInfo(rebar.size):.2f} (kg) * 2')
                break  # middle tie rebar number equal to rebar line number, so only count one middle tie
        for tie in self.tie_list:
            if tie.size in self.tie_count:
                self.tie_count[tie.size] += tie.count * \
                    RebarInfo(tie.size) * (self.depth -
                                           10 + self.width - 10) * 2
            else:
                self.tie_count[tie.size] = tie.count * \
                    RebarInfo(tie.size) * (self.depth -
                                           10 + self.width - 10) * 2
            self.detail_report.append(
                f'箍筋:{tie}= {tie.count} * {(self.depth - 10 + self.width - 10)} (cm) * {RebarInfo(tie.size):.2f} (kg) * 2')
        self.concrete = (self.depth - 15)*self.width * \
            self.length / (100*100*100)
        self.formwork = (self.width + (self.depth - 15)*2) * \
            self.length / (100*100)

    def get_rebar_weight(self):
        temp = 0
        for size, rebar in self.rebar_count.items():
            temp += rebar
        return temp

    def get_tie_weight(self):
        temp = 0
        for size, rebar in self.tie_count.items():
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
        if (self.middle_tie):
            return self.middle_tie[0].text
        return

    def get_concrete(self):
        return self.concrete

    def get_formwork(self):
        return self.formwork

    def write_beam(self, df: pd.DataFrame):
        pass

    def set_prop(self, floor: Floor):
        self.height = floor.height
        self.fc = floor.material_list['fc']
        self.fy = floor.material_list['fy']
        self.floor_object = floor
    # 取得梁配筋面積

    def get_rebar_table(self, rebar_type1: RebarType, rebar_type2: RebarType) -> float:
        As = 0
        for rebar in self.rebar_table[rebar_type1.value][rebar_type2.value]:
            As += rebar.As
        return round(As, 2)
    # 整理梁配筋成常用表格

    def sort_rebar_table(self):
        min_diff = 30
        self.rebar['top_first'].sort(key=lambda r: r.arrow_coor[0][0])
        # for i,rebar in enumerate(self.rebar['top_first']):
        #     if i == 0:
        #         self.rebar_table['top']['left'].append(rebar)
        #         self.rebar_table['top_length']['left'].append(rebar.length)
        #     if i == 1:
        #         self.rebar_table['top']['middle'].append(rebar)
        #         self.rebar_table['top_length']['middle'].append(rebar.length)
        #     if i == 2:
        #         self.rebar_table['top']['right'].append(rebar)
        #         self.rebar_table['top_length']['right'].append(rebar.length)
        for rebar in self.rebar['top_first']:
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff:
                self.rebar_table['top']['left'].append(rebar)
            if abs(rebar.end_pt.x - self.end_pt.x) < min_diff:
                self.rebar_table['top']['right'].append(rebar)
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x) >= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
                self.rebar_table['top']['middle'].append(rebar)
            if abs(rebar.start_pt.x - self.start_pt.x) < min_diff:
                self.rebar_table['top_length']['left'].append(rebar.length)
                continue
            if abs(rebar.end_pt.x - self.end_pt.x) < min_diff:
                self.rebar_table['top_length']['right'].append(rebar.length)
                continue
            if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x) >= min_diff):
                self.rebar_table['top_length']['middle'].append(rebar.length)
                continue

        if len(self.rebar_table['top']['middle']) > 1 and len(self.rebar_table['top']['left']) == 0:
            self.rebar_table['top']['left'].append(
                self.rebar_table['top']['middle'].pop(0))
            self.rebar_table['top_length']['left'].append(
                self.rebar_table['top_length']['middle'].pop(0))
            self.start_pt.x = self.rebar_table['top']['left'][0].start_pt.x

        if len(self.rebar_table['top']['middle']) > 1 and len(self.rebar_table['top']['right']) == 0:
            self.rebar_table['top']['right'].append(
                self.rebar_table['top']['middle'].pop())
            self.rebar_table['top_length']['right'].append(
                self.rebar_table['top_length']['middle'].pop())
            self.end_pt.x = self.rebar_table['top']['right'][0].end_pt.x

        for rebar in self.rebar['top_second']:
            temp = min(list(self.rebar_table['top'].values()),
                       key=lambda r_table:
                       abs(rebar.arrow_coor[0][0] - r_table[0].arrow_coor[0]
                           [0]) if r_table else float('inf')
                       )
            diff_dis = abs(rebar.arrow_coor[0][0] - temp[0].arrow_coor[0][0])
            for r_list in [v for v in list(self.rebar_table['top'].values()) if v and
                           abs(rebar.arrow_coor[0][0] - v[0].arrow_coor[0][0]) == diff_dis]:
                r_list.append(rebar)

        # for rebar in self.rebar['top_second']:

        #     if abs(rebar.start_pt.x - self.start_pt.x) < min_diff:
        #         self.rebar_table['top']['left'].append(rebar)
        #     if abs(rebar.end_pt.x - self.end_pt.x) < min_diff:
        #         self.rebar_table['top']['right'].append(rebar)
        #     if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and
        #             abs(rebar.end_pt.x - self.end_pt.x) >= min_diff) or \
        #             (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
        #         self.rebar_table['top']['middle'].append(rebar)

        self.rebar['bot_first'].sort(key=lambda r: r.arrow_coor[0][0])

        for i, rebar in enumerate(self.rebar['bot_first']):
            if i == 0:
                self.rebar_table['bottom']['left'] = [rebar]
                self.rebar_table['bottom']['middle'] = [rebar]
                self.rebar_table['bottom']['right'] = [rebar]
                self.rebar_table['bottom_length']['left'].append(rebar.length)
            if i == 1:
                self.rebar_table['bottom']['middle'] = [rebar]
                self.rebar_table['bottom']['right'] = [rebar]
                self.rebar_table['bottom_length']['middle'].append(
                    rebar.length)
            if i == 2:
                self.rebar_table['bottom']['right'] = [rebar]
                self.rebar_table['bottom_length']['right'].append(rebar.length)

        for rebar in self.rebar['bot_second']:
            temp = min(self.rebar_table['bottom'].items(),
                       key=lambda r_table: abs(rebar.arrow_coor[0][0] - r_table[1][0].arrow_coor[0][0]))[1]
            diff_dis = abs(rebar.arrow_coor[0][0] - temp[0].arrow_coor[0][0])
            for r_list in [v for k, v in self.rebar_table['bottom'].items() if
                           abs(rebar.arrow_coor[0][0] - v[0].arrow_coor[0][0]) == diff_dis]:
                r_list.append(rebar)
            # temp.append(rebar)
        # for rebar in self.rebar['bot_second']:
        #     if abs(rebar.start_pt.x - self.start_pt.x) < min_diff :
        #         self.rebar_table['bottom']['left'].append(rebar)
        #     if abs(rebar.end_pt.x - self.end_pt.x)< min_diff:
        #         self.rebar_table['bottom']['right'].append(rebar)
        #     if (abs(rebar.start_pt.x - self.start_pt.x) >= min_diff and abs(rebar.end_pt.x - self.end_pt.x)>= min_diff) or (rebar.start_pt.x == self.start_pt.x and rebar.end_pt.x == self.end_pt.x):
        #         self.rebar_table['bottom']['middle'].append(rebar)
        if len(self.rebar_table['top']['middle']) == 0:
            if self.rebar_table['top_length']['left'] > self.rebar_table['top_length']['right']:
                self.rebar_table['top']['middle'].extend(
                    self.rebar_table['top']['left'])
            else:
                self.rebar_table['top']['middle'].extend(
                    self.rebar_table['top']['right'])

        if len(self.rebar_table['top']['right']) == 0:
            self.rebar_table['top']['right'].extend(
                self.rebar_table['top']['middle'])

        if len(self.rebar_table['top']['left']) == 0:
            self.rebar_table['top']['left'].extend(
                self.rebar_table['top']['middle'])

        if len(self.rebar_table['bottom']['middle']) == 0:
            if self.rebar_table['bottom_length']['left'] > self.rebar_table['bottom_length']['right']:
                self.rebar_table['bottom']['middle'].extend(
                    self.rebar_table['bottom']['left'])
            else:
                self.rebar_table['bottom']['middle'].extend(
                    self.rebar_table['bottom']['right'])

        if len(self.rebar_table['bottom']['right']) == 0:
            self.rebar_table['bottom']['right'].extend(
                self.rebar_table['bottom']['middle'])

        if len(self.rebar_table['bottom']['left']) == 0:
            self.rebar_table['bottom']['left'].extend(
                self.rebar_table['bottom']['middle'])

        self.cal_rebar_ratio()
        self.cal_ld_table()
    # 計算梁配筋比

    def cal_rebar_ratio(self):
        for rebar_type in [RebarType.Top, RebarType.Bottom]:
            for rebar_type2 in [RebarType.Left, RebarType.Middle, RebarType.Right]:
                try:
                    self.rebar_ratio[(rebar_type, rebar_type2)] = self.get_rebar_table(rebar_type1=rebar_type,
                                                                                       rebar_type2=rebar_type2)/(self.width *
                                                                                                                 (self.depth - self.protect_layer))
                except:
                    self.rebar_ratio[(rebar_type, rebar_type2)] = 0

    def get_rebar_ratio(self):
        return [
            self.rebar_ratio[(RebarType.Top, RebarType.Left)],
            self.rebar_ratio[(RebarType.Top, RebarType.Middle)],
            self.rebar_ratio[(RebarType.Top, RebarType.Right)],
            self.rebar_ratio[(RebarType.Bottom, RebarType.Left)],
            self.rebar_ratio[(RebarType.Bottom, RebarType.Middle)],
            self.rebar_ratio[(RebarType.Bottom, RebarType.Right)],
        ]

    def cal_ld(self, rebar: Rebar, tie: Tie):
        from math import sqrt, ceil
        if self.beam_type == BeamType.FB:
            cover = 7.5
        else:
            cover = 4
        fy = self.fy
        fc = self.fc
        fydb = RebarDiameter(rebar.size)
        fytdb = RebarDiameter(tie.size)
        spacing = tie.spacing
        if self.floor_object is not None:
            if self.beam_type == BeamType.Grider and self.floor_object.is_seismic:
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
            ld_simple_top = fy * psitTop_ * psie_ / \
                (sqrt(fc) * 5.3 * lamda_) * fydb
            ld_simple_bot = fy * psitBot_ * psie_ / \
                (sqrt(fc) * 5.3 * lamda_) * fydb
        else:
            psis_ = 0.8
            ld_simple_top = fy * psitTop_ * psie_ / \
                (sqrt(fc) * 6.6 * lamda_) * fydb
            ld_simple_bot = fy * psitBot_ * psie_ / \
                (sqrt(fc) * 6.6 * lamda_) * fydb

        cs_ = ((width_ - fydb * fynum - fytdb*2 -
               cover * 2)/(fynum - 1) + fydb) / 2
        cc_ = cover + fytdb + fydb / 2
        if cs_ <= cc_:
            cb_ = cs_
            atr_ = 2 * avh
            ktr_ = atr_ * 40 / (spacing * fynum)
        else:
            cb_ = cs_
            atr_ = avh
            ktr_ = atr_ * 40 / (spacing * fynum)

        botFactor = psitBot_ * psie_ * psis_ * \
            lamda_ / min((cb_ + ktr_) / fydb, 2.5)
        topFactor = psitTop_ * botFactor

        bot_ld = botFactor * ld
        top_ld = topFactor * ld

        bot_lap_ld = ceil(1.3 * min(ld_simple_bot, bot_ld))
        top_lap_ld = ceil(1.3 * min(ld_simple_top, top_ld))

        return (top_lap_ld, bot_lap_ld)

    def cal_ld_table(self):
        self.ld_table = {}
        if self.rebar_table[RebarType.Top.value][RebarType.Left.value] and self.tie['left']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Top.value][RebarType.Left.value][0],
                                                 tie=self.tie['left'])
            self.ld_table.update({(RebarType.Top, RebarType.Left): top_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Middle.value] and self.tie['middle']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Top.value][RebarType.Middle.value][0],
                                                 tie=self.tie['middle'])
            self.ld_table.update(
                {(RebarType.Top, RebarType.Middle): top_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Right.value] and self.tie['right']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Top.value][RebarType.Right.value][0],
                                                 tie=self.tie['right'])
            self.ld_table.update(
                {(RebarType.Top, RebarType.Right): top_lap_ld})

        if self.rebar_table[RebarType.Bottom.value][RebarType.Left.value] and self.tie['left']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Bottom.value][RebarType.Left.value][0],
                                                 tie=self.tie['left'])
            self.ld_table.update(
                {(RebarType.Bottom, RebarType.Left): bot_lap_ld})
        if self.rebar_table[RebarType.Bottom.value][RebarType.Middle.value] and self.tie['middle']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Bottom.value][RebarType.Middle.value][0],
                                                 tie=self.tie['middle'])
            self.ld_table.update(
                {(RebarType.Bottom, RebarType.Middle): bot_lap_ld})
        if self.rebar_table[RebarType.Top.value][RebarType.Right.value] and self.tie['right']:
            top_lap_ld, bot_lap_ld = self.cal_ld(rebar=self.rebar_table[RebarType.Bottom.value][RebarType.Right.value][0],
                                                 tie=self.tie['right'])
            self.ld_table.update(
                {(RebarType.Bottom, RebarType.Right): bot_lap_ld})
