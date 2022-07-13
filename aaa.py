import pandas as pd

filename = '0713.xlsx'
df = pd.read_excel(filename)
new_list = [('task1', '0.1%', '0.1%', '0.1%', '0.1%', '0.1%', '0.1%', '0.1%')]
dfNew=pd.DataFrame(new_list, columns = ['名稱' , 'in plan not in beam 大梁', 'in plan not in beam 小梁','in beam not in plan 大梁', 'in plan not In beam 小梁', '執行時間', '執行日期' , '備註'])
df=pd.concat([df, dfNew], axis=0, ignore_index = True)
df.to_excel('0713.xlsx')