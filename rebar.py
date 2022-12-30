

_rebar = {
    "#3":0.957,
    "#4":1.267,
    "#5":1.986,
    "#6":2.856,
    "#7":3.871,
    "#8":5.067,
    "#10":8.143,
    "#11":10.07
}

def RebarInfo(size="#3"):
    if "E.F" in size:return _rebar[size.replace("E.F","")]*2
    return _rebar[size]*7.85

if __name__ == '__main__':
    l = [0]
    print(l[0])