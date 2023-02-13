

_rebar = {
    "#3":0.957,
    "#4":1.267,
    "#5":1.986,
    "#6":2.856,
    "#7":3.871,
    "#8":5.067,
    "#10":8.143,
    "#11":10.07,
    "3":0.957,
    "4":1.267,
    "5":1.986,
    "6":2.856,
    "7":3.871,
    "8":5.067,
    "10":8.143,
    "11":10.07
}

def RebarInfo(size="#3"):
    if "(E.F)" in size:return _rebar[size.replace("(E.F)","").replace(" ","")]*2
    if "E.F" in size:return _rebar[size.replace("E.F","")]*2
    return _rebar[size]*7.85
def RebarArea(size="#3"):
    if "(E.F)" in size:return _rebar[size.replace("(E.F)","").replace(" ","")]*2
    if "E.F" in size:return _rebar[size.replace("E.F","")]*2
    return _rebar[size]

if __name__ == '__main__':

    class testclass:
        def __init__(self,i) -> None:
            self.val = i
            pass
        def add_value(self,i):
            print('in map')
            self.val += i 
        def __str__(self) -> str:
            return self.val
            pass
    # def full2half(c: str) -> str:
    #     return chr(ord(c)-65248)


    # def half2full(c: str) -> str:
    #     return chr(ord(c)+65248)
    
    # print('(45X90)'.replace(' ', '').replace('X','x'))
    # test_list = [testclass(1),testclass(5),testclass(10)]
    # list(map(lambda i:i.add_value(5),test_list))
    # for t in test_list:
    #     print(t.val)
    # pass
    # temp = {'1':{'10':10,'11':11,'12':12},'2':{'100':100,'110':110,'120':120}}
    # for i,j in temp['1'].items():
    #     # j += 1
    #     temp['1'][i] += 1
    #     print(j)
    #     print(temp['1'][i]) 
    #     pass
    # print(temp)
    print(list(['1','2']))