from __future__ import annotations
import re
import pandas as pd
from rebar import RebarInfo
from beam import Point,Rebar
class Column:
    size = ''
    serial = ''
    floor = ''
    rebar:list[Rebar]
    def __init__(self):
        pass