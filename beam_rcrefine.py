from item.beam import Beam, RebarType, Rebar, Tie
from utils.demand import calculate_beam_earthquake_load, calculate_beam_gravity_load
from item.rebar import RebarArea, RebarFy, RebarDiameter, RebarInfo
from itertools import product
from math import ceil, sqrt
from copy import deepcopy, copy
'''
Every item has its score, like fc, rebar size, rebar arrange
To opt the result of the beam design (or better than origin design)
'''
opt_item = {
    "width": [5, -5, 0],
    "depth": [5, -5, 0],
    "fc": [70, -70, 0],
    "rebar_size": ["#7", "#8", "#10"],
    "rebar_db": [1, 1.5],
    "tie_size": ["#3", "#4", "#5"],
    "tie_spacing": [0, 10, 12, 15, 20]
}

score_item = {
    "#7": [5, 3, 1],
    "#8": [3, 1, 1],
    "#10": [1, 1, 1]
}

price_item = {
    "concrete": {},
    "rebar": {},
    "formwork": []
}


def main(beam_list: list[Beam]):
    # read beam data
    for beam in beam_list:
        w, Vg = calculate_beam_gravity_load(beam=beam)
        left_Meq, right_Meq, Veq, Vp = calculate_beam_earthquake_load(
            w, beam=beam, pos=RebarType.Top)
        for opt, item_list in opt_item.items():

            # cal beam score

            # refine by fc, width, depth, rebar_size

            # compare with origin data


def refine_beam(beam: Beam, parameter: dict, Meq, Mg, Veq, Vp, Vg):
    phi = 0.9
    left_Meq, right_Meq = Meq

    left_top_Mu = (left_Meq + Mg)
    right_top_Mu = (right_Meq + Mg)

    left_bot_Mu = (left_Meq)
    right_bot_Mu = (right_Meq)

    Vu = max(Veq + Vg, Vp)

    configs = []
    keys, values = zip(*opt_item.items())
    for combination in product(*values):
        configs += [dict(zip(keys, combination))]

    for config in configs:
        rebar_table = {
            RebarType.Top.value: {
                RebarType.Left.value: [],
                RebarType.Middle.value: [],
                RebarType.Right.value: [],
            },
            RebarType.Bottom.value: {
                RebarType.Left.value: [],
                RebarType.Middle.value: [],
                RebarType.Right.value: [],
            }
        }

        # refine longitudinal rebar
        for Mu, pos in zip([left_top_Mu, right_top_Mu, left_bot_Mu, right_bot_Mu],
                           [(RebarType.Top, RebarType.Left), (RebarType.Top, RebarType.Right), (RebarType.Bottom, RebarType.Left), (RebarType.Bottom, RebarType.Right)]):
            pos1, pos2 = pos
            AsFy = Mu / 0.9 / \
                (beam.depth - beam.protect_layer + config['depth'])
            origin_rebar = deepcopy(beam.rebar_table[pos1][pos2][0])
            As = AsFy / RebarFy(config['rebar_size'])
            rebar_num = ceil(As / RebarArea(config['rebar_size']))
            rebar_max_num = (beam.width + config['width'] - beam.protect_layer * 2) / (
                config['rebar_db'] * RebarDiameter(config['rebar_size'])) + 1
            if rebar_num > rebar_max_num and rebar_num <= rebar_max_num * 2:
                first = rebar_max_num
                second = rebar_num - rebar_max_num
                if second < 2:
                    first = rebar_num - 2
                    second = 2
                first_rebar = deepcopy(origin_rebar)
                first_rebar.set_new_property(first, config['rebar_size'])
                rebar_table[pos1.value][pos2.value].append(first_rebar)

                second_rebar = deepcopy(origin_rebar)
                second_rebar.set_new_property(second, config['rebar_size'])
                rebar_table[pos1.value][pos2.value].append(second_rebar)

            elif rebar_num <= rebar_max_num:
                first_rebar = deepcopy(origin_rebar)
                first_rebar.set_new_property(first, config['rebar_size'])
                rebar_table[pos1.value][pos2.value].append(first_rebar)
        rebar_table
        # refine shear rebar
        tie_table = {
            'left': None,
            'middle': None,
            'right': None
        }

        for pos in ['left', 'middle', 'right']:
            tie = beam.tie[pos]

            Vs = round(tie.Ash * tie.fy*(beam.depth -
                                         beam.protect_layer)/tie.spacing, 2)
            if pos == "middle":
                Vc = 0
            else:
                Vc = 0.53 * beam.width * \
                    (beam.depth - beam.protect_layer) * sqrt(beam.fc)

            Vs = (Vu - 0.75 * Vc) / 0.75
            Ash_demand = Vs / (beam.depth - beam.protect_layer)
            for tie_size in config['tie_size']:
                Ash = RebarArea(tie_size) * 2
                fy = RebarFy(tie_size)
                require_spacing = Ash * fy / Ash_demand
                real_spacing = min(
                    [spacing for spacing in config["tie_spacing"] if spacing < require_spacing])
                if real_spacing == 0:
                    require_spacing /= 2
                    real_spacing = min(
                        [spacing for spacing in config["tie_spacing"] if spacing < require_spacing])
                    if real_spacing != 0:
                        break
            real_count = ceil(tie.spacing * tie.count / real_spacing)
            new_tie = Tie(f"{tie_size}@{real_spacing}",
                          (0, 0), tie_size, real_count, tie_size)
            tie_table[pos] = new_tie


def calculate_material_cost(rebar_table, tie_table, width, depth, fc):
    for pos, rebars_list in rebar_table.items():
        for pos2, rebars in rebar_table[pos]:
            for rebar in rebars:
                self.rebar_count[rebar.size] = rebar.length * \
                    rebar.number * RebarInfo(rebar.size)
    if rebar.size in self.rebar_count:
        self.rebar_count[rebar.size] += rebar.length * \
            rebar.number * RebarInfo(rebar.size)
    else:
        self.rebar_count[rebar.size] = rebar.length * \
            rebar.number * RebarInfo(rebar.size)


def cal_Rebar_As(Mu, b, d, protect_layer, fy, fc):
    Mn = Mu / 0.9
    tho_b = 0.85 * fc / fy * 0.85


def refine_beam_rebar():
    pass


def refine_beam_fc():
    pass


def refine_beam_width():
    pass
