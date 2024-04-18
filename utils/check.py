import pandas as pd
from math import sqrt
from item.beam import Beam
from item.beam import RebarType
from item.column import Column
from utils.demand import calculate_beam_earthquake_load, calculate_beam_gravity_load


def check_beam_shear_strength(beam: Beam):
    w, Vg = calculate_beam_gravity_load(beam=beam)
    left_Meq, right_Meq, Veq, Vp = calculate_beam_earthquake_load(
        w, beam=beam, pos=RebarType.Top.value)
    tie = beam.tie['middle']
    Vs = round(tie.Ash * tie.fy*(beam.depth -
                                 beam.protect_layer)/tie.spacing, 2)
    Vn = 0.53 * beam.width * \
        (beam.depth - beam.protect_layer) * sqrt(beam.fc) + Vs

    phi = 0.75

    ratio_Vp = phi * Vn / (Vp)

    ratio_Veq = phi * Vn / (Vg + Veq)

    return ratio_Vp, ratio_Veq
