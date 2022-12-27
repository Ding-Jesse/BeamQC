class Point:
    x = 0
    y = 0
    def __init__(self):
        pass
class Rebar:
    start_pt = Point
    length = 0
    text = 0
    def __init__(self):
        pass

class Tie:
    start_pt=Point
    length = 0
    text = 0
    def __init__(self):
        pass
class Beam:
    rebar={
        'top_first':list(Rebar),
        'top_second':list(Rebar),
        'bot_first':list(Rebar),
        'bot_second':list(Rebar),
    }
    tie ={
        'left':Tie,
        'middle':Tie,
        'right':Tie
    }
    serial = ''
    floor = ''
    depth = 0
    width = 0
    coor = Point
    def __init__(self):
        pass