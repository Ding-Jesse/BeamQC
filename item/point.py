class Point:
    x = 0
    y = 0
    def __init__(self,*pt):
        if len(pt) == 0:
            pass
        elif isinstance(pt[0],tuple):
            self.x = pt[0][0]
            self.y = pt[0][1]
        elif isinstance(pt,tuple):
            self.x = pt[0]
            self.y = pt[1]