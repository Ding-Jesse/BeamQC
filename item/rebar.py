import pandas as pd
# import item.floor

_rebar = {
    "#3": 0.957,
    "#4": 1.267,
    "#5": 1.986,
    "#6": 2.856,
    "#7": 3.871,
    "#8": 5.067,
    "#10": 8.143,
    "#11": 10.07,
    "3": 0.957,
    "4": 1.267,
    "5": 1.986,
    "6": 2.856,
    "7": 3.871,
    "8": 5.067,
    "10": 8.143,
    "11": 10.07
}
_rebar_dia = {
    "#3": 0.9525,
    "#4": 1.27,
    "#5": 1.5785,
    "#6": 1.905,
    "#7": 2.222,
    "#8": 2.54,
    "#10": 3.22,
    "#11": 3.58,
    "3": 0.9525,
    "4": 1.27,
    "5": 1.5785,
    "6": 1.905,
    "7": 2.222,
    "8": 2.54,
    "10": 3.226,
    "11": 3.581
}
rebar_dict = {}


def RebarInfo(size="#3"):
    if "(E.F.)" in size:
        size = size.replace("(E.F.)", "").replace(" ", "")
    if "E.F." in size:
        size = size.replace("E.F.", "").replace(" ", "")
    if "(E.F)" in size:
        size = size.replace("(E.F)", "").replace(" ", "")
    if "E.F" in size:
        size = size.replace("E.F", "").replace(" ", "")
    global rebar_dict
    if size in rebar_dict:
        return rebar_dict[size]['重量（kg/m）'] / 100 * 1000
    return _rebar[size]*7.85


def RebarArea(size="#3"):
    if "(E.F.)" in size:
        size = size.replace("(E.F.)", "").replace(" ", "")
    if "E.F." in size:
        size = size.replace("E.F.", "").replace(" ", "")
    if "(E.F)" in size:
        size = size.replace("(E.F)", "").replace(" ", "")
    if "E.F" in size:
        size = size.replace("E.F", "").replace(" ", "")
    global rebar_dict
    if size in rebar_dict:
        return rebar_dict[size]['截面積（cm²）']
    # if "(E.F)" in size:return _rebar[size.replace("(E.F)","").replace(" ","")]*2
    # if "E.F" in size:return _rebar[size.replace("E.F","").replace(" ","")]*2
    return _rebar[size]


def RebarDiameter(size="#3"):
    global rebar_dict
    if size in rebar_dict:
        return rebar_dict[size]['直徑（cm）']
    return _rebar_dia[size]


def RebarFy(size="#3"):
    global rebar_dict
    if size in rebar_dict:
        return rebar_dict[size]['fy(kgf/cm2)']
    return 4200


def isRebarSize(size):
    try:
        RebarInfo(size=size)
        return True
    except:
        return False


def readRebarExcel(file_path: str):
    global rebar_dict
    # rebar_df:pd.DataFrame = item.floor.read_parameter_df(read_file=file_path,sheet_name="鋼筋資料表")
    rebar_df: pd.DataFrame = pd.read_excel(
        file_path, sheet_name="鋼筋資料表", header=[0])
    rebar_df.set_index('鋼筋尺寸', inplace=True)
    rebar_dict = rebar_df.to_dict('index')
    pass


def PrintRebarDict():
    global rebar_dict
    print(rebar_dict)


if __name__ == '__main__':
    readRebarExcel(r'D:\Desktop\BeamQC\file\樓層參數_floor.xlsx')
    PrintRebarDict()
    # print(','.join([]))
    # import pandas as pd
    # ng_df = pd.DataFrame(columns = ['樓層','編號','備註'],index=[])
    # for i in range(10):
    #     temp_df = pd.DataFrame(data={'樓層':1,'編號':1,'備註':1},index=['temp'])
    #     ng_df = pd.concat([ng_df,temp_df],ignore_index=True)
    # print(ng_df)
    # class testclass:
    #     def __init__(self,i) -> None:
    #         self.val = i
    #         pass
    #     def add_value(self,i):
    #         print('in map')
    #         self.val += i
    #     def __str__(self) -> str:
    #         return self.val
    #         pass
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
