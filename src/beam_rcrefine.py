from item.beam import Beam, RebarType, Rebar, Tie
from utils.demand import calculate_beam_earthquake_load, calculate_beam_gravity_load
from item.rebar import RebarArea, RebarFy, RebarDiameter, RebarInfo
from itertools import product
from math import ceil, sqrt, floor
from copy import deepcopy, copy
from src.save_temp_file import read_temp
from beam_count import floor_parameter
'''
Every item has its score, like fc, rebar size, rebar arrange
To opt the result of the beam design (or better than origin design)
'''
# opt_item = {
#     "width": [5, -5, 0],
#     "depth": [5, -5, 0],
#     "fc": [70, -70, 0],
#     "rebar_size": ["#7", "#8", "#10"],
#     "rebar_db": [1, 1.5],
# }
opt_item = {
    "width": [0],
    "depth": [0],
    "fc": [0],
    "rebar_size": ["#7", "#8", "#10"],
    "rebar_db": [1, 1.5],
}
tie_item = {
    "tie_size": ["#3", "#4", "#5"],
    "tie_spacing": [0, 10, 12, 15, 20]
}

score_item = {
    "#7": [5, 3, 1],
    "#8": [3, 1, 1],
    "#10": [1, 1, 1]
}

# price_item = {
#     "concrete": {},
#     "rebar": {},
#     "formwork": []
# }
# Read Evaluation Data
price_item = {'rebar': 30000,
              'concrete': {280: 3400, 350: 3600, 420: 3800, 490: 4200},
              'formwork': 1000}


def main(beam_list: list[Beam]):
    deliemeter = '\n'
    # read beam data
    for beam in beam_list:
        beam.cal_rebar()
        beam.sort_rebar_table()
        w, Vg, Mg = calculate_beam_gravity_load(beam=beam)
        left_Meq, right_Meq, Veq, Vp = calculate_beam_earthquake_load(
            w, beam=beam, pos=RebarType.Top)
        origin_cost, results = refine_beam(beam=beam,
                                           opt_item=opt_item,
                                           Meq=(left_Meq, right_Meq),
                                           Mg=Mg,
                                           Veq=Veq,
                                           Vp=Vp,
                                           Vg=Vg)
        min_result = min(results, key=lambda result: result[1]['total'])
        print(f'{origin_cost}')
        print(f'config:{min_result[0]},material_cost:{min_result[1]}')
        print(
            f'origin:{beam.rebar_table}{deliemeter}after:{min_result[2].rebar_table}')
        return origin_cost, results
        # new_beam = deepcopy(beam)

        # cal beam score

        # refine by fc, width, depth, rebar_size

        # compare with origin data


def refine_beam(beam: Beam, opt_item: dict, Meq, Mg, Veq, Vp, Vg):
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

    results = []
    origin_cost = calculate_material_cost(beam=beam, para=price_item)

    for config in configs:
        status = 'available'
        new_beam = deepcopy(beam)
        if 'depth' in config:
            new_beam.depth = config['depth'] + beam.depth
        if 'width' in config:
            new_beam.width = config['width'] + beam.width
        if 'fc' in config:
            new_beam.fc = config['fc'] + beam.fc
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
        for Mu, pos in zip([left_top_Mu, right_top_Mu, left_bot_Mu, right_bot_Mu, Mg],
                           [(RebarType.Top, RebarType.Left),
                            (RebarType.Top, RebarType.Right),
                            (RebarType.Bottom, RebarType.Left),
                            (RebarType.Bottom, RebarType.Right),
                            (RebarType.Bottom, RebarType.Middle)]):
            pos1, pos2 = pos
            AsFy = Mu / phi / \
                (beam.depth - beam.protect_layer + config['depth'])
            origin_rebar = deepcopy(
                beam.rebar_table[pos1.value][pos2.value][0])
            As = AsFy / RebarFy(config['rebar_size'])
            rebar_num = ceil(As / RebarArea(config['rebar_size']))
            rebar_max_num = floor((beam.width + config['width'] - beam.protect_layer * 2) / (
                (1+config['rebar_db']) * RebarDiameter(config['rebar_size'])) + 1)
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
            else:
                status = "not available"
                print(f'{config} : {pos1.value}-{pos2.value}')
                break
            new_beam.rebar_table[pos1.value][pos2.value] = rebar_table[pos1.value][pos2.value]
        if status == "not available":
            continue
        # refine shear rebar
        tie_table = {
            'left': None,
            'middle': None,
            'right': None
        }
        real_spacing = 0
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
            for tie_size in tie_item['tie_size']:
                Ash = RebarArea(tie_size) * 2
                fy = RebarFy(tie_size)
                require_spacing = Ash * fy / Ash_demand
                real_spacing = min(
                    [spacing for spacing in tie_item["tie_spacing"] if spacing < require_spacing])
                if real_spacing == 0:
                    require_spacing /= 2
                    real_spacing = min(
                        [spacing for spacing in tie_item["tie_spacing"] if spacing < require_spacing])
                    if real_spacing != 0:
                        break
            if real_spacing == 0:
                status == 'not available'
                print(f'{config} : shear-{pos}')
                break
            real_count = ceil(tie.spacing * tie.count / real_spacing)
            new_tie = Tie(f"{tie_size}@{real_spacing}",
                          (0, 0), tie_size, real_count, tie_size)
            tie_table[pos] = new_tie
            new_beam.tie[pos] = tie_table[pos]
        if status == 'not available':
            continue

        new_beam.cal_rebar()
        material_cost = calculate_material_cost(beam=new_beam, para=price_item)
        results.append((config, material_cost, new_beam))
    return origin_cost, results


def calculate_material_cost(beam: Beam, para: dict):
    rebar_price = 0
    concrete_price = 0
    formwork_price = 0
    for size, amount in beam.rebar_count.items():
        rebar_price += para['rebar'] * amount / 100 / 7.85 / 1000
    concrete_price += beam.get_concrete() * para['concrete'][beam.fc]
    formwork_price += beam.get_formwork() * para['formwork']
    return {
        'rebar': rebar_price,
        'concrete': concrete_price,
        'formwork': formwork_price,
        'total': rebar_price + concrete_price + formwork_price
    }


def cal_Rebar_As(Mu, b, d, protect_layer, fy, fc):
    Mn = Mu / 0.9
    tho_b = 0.85 * fc / fy * 0.85


def refine_beam_rebar():
    pass


def refine_beam_fc():
    pass


def refine_beam_width():
    pass


if __name__ == "__main__":
    beam_list = read_temp(
        tmp_file=r'TEST\RCREFINE\2024-0522 247-2024-05-22-11-40-temp-beam_list.pkl')
    floor_parameter(beam_list=beam_list,
                    floor_parameter_xlsx=r'TEST\RCREFINE\基本資料表 _ 三重427.xlsx')
    main(beam_list=[beam_list[145]])
