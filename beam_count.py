import win32com.client
import pythoncom
import time

def vtFloat(l): #要把點座標組成的list轉成autocad看得懂的樣子？
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)

def error(error_message): # 把錯誤訊息印到error.log裡面
    f = open(error_file, 'a', encoding = 'utf-8')
    localtime = time.asctime(time.localtime(time.time()))
    f.write(f'{localtime} | {error_message}\n')
    f.close
    return

def progress(message, progress_file): # 把進度印到progress裡面，在app.py會對這個檔案做事
    f = open(progress_file, 'a', encoding = 'utf-8')
    f.write(f'{message}\n')
    f.close
    return

def read_beam(beam_filename, rebar_layer, rebar_data_layer, tie_text_layer, block_layer, beam_text_layer, result_filename, tie_file, progress_file):
    error_count = 0
    progress('開始讀取梁配筋圖', progress_file)
    # Step 1. 打開應用程式
    flag = 0
    while not flag and error_count <= 10:
        try:
            wincad_beam = win32com.client.Dispatch("AutoCAD.Application")
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 1: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 1/15', progress_file)

    # Step 2. 匯入檔案
    flag = 0
    while not flag and error_count <= 10:
        try:
            doc_beam = wincad_beam.Documents.Open(beam_filename)
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 2: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 2/15', progress_file)

    # Step 3. 匯入modelspace
    flag = 0
    while not flag and error_count <= 10:
        try:
            msp_beam = doc_beam.Modelspace
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 3: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 3/15', progress_file)

    # # Step 4 解鎖所有圖層 -> 不然不能刪東西
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         layer_count = doc_beam.Layers.count
    #         for x in range(layer_count):
    #             layer = doc_beam.Layers.Item(x)
    #             layer.Lock = False
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 4: {e}, error_count = {error_count}.')
    # progress('梁配筋圖讀取進度 4/15', progress_file)

    # # Step 5. (1) 遍歷所有物件 -> 炸圖塊; (2) 刪除我們不要的條件 -> 省時間  
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         count = 0
    #         total = msp_beam.Count
    #         progress(f'正在炸梁配筋圖的圖塊及篩選判斷用的物件，梁配筋圖上共有{total}個物件，大約運行{int(total / 9000) + 1}分鐘，請耐心等候', progress_file)
    #         layer_list = [rebar_layer, rebar_data_layer]
    #         for object in msp_beam:
    #             count += 1
    #             if object.EntityName == "AcDbBlockReference" and object.Layer in layer_list:
    #                 object.Explode()
    #             if object.Layer not in layer_list:
    #                 object.Delete()
    #             if count % 1000 == 0:
    #                 progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
    #         flag = 1
        
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 5: {e}, error_count = {error_count}.')
    #         msp_beam = doc_beam.Modelspace
    # progress('梁配筋圖讀取進度 5/15', progress_file)

    # # Step 6. 重新匯入modelspace
    # flag = 0
    # while not flag and error_count <= 10:
    #     try:
    #         msp_beam = doc_beam.Modelspace
    #         flag = 1
    #     except Exception as e:
    #         error_count += 1
    #         time.sleep(5)
    #         error(f'read_beam error in step 6: {e}, error_count = {error_count}.')
    # progress('梁配筋圖讀取進度 6/15', progress_file)
    
    # Step 7. 遍歷所有物件 -> 完成 floor_to_beam_set，格式為(floor, beam, coor, size)
    progress('正在遍歷梁配筋圖上的物件並篩選出有效信息，運行時間取決於梁配筋圖大小，請耐心等候', progress_file)
    coor_to_rebar_list = [] # (頭座標，尾座標，長度)
    coor_to_bend_rebar_list = [] # (直的端點，橫的端點，長度)
    coor_to_data_list = [] # (字串，座標)
    coor_to_arrow_dic = {} # 尖點座標 -> 文字連接處座標
    coor_to_tie_list = [] # (下座標，上座標，長度) 
    coor_to_tie_text_list = [] # (字串，座標)
    coor_to_block_list = [] # ((左下，右上), rebar_length_dic, tie_count_dic)
    coor_to_beam_list = [] # (string, midpoint, list of tie, tie_count_dic)
    flag = 0
    while not flag and error_count <= 10:
        try:
            count = 0
            total = msp_beam.Count
            progress(f'梁配筋圖上共有{total}個物件，大約運行{int(total / 5500)}分鐘，請耐心等候', progress_file)
            for object in msp_beam:
                count += 1
                if count % 1000 == 0:
                    progress(f'梁配筋圖已讀取{count}/{total}個物件', progress_file)
                # 抓鋼筋的字的座標
                if object.Layer == rebar_data_layer and object.ObjectName == "AcDbMText":
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2))
                    coor_to_data_list.append((object.TextString, coor))
                # 抓箭頭座標
                elif object.Layer == rebar_data_layer and object.ObjectName == "AcDbLeader":
                    # object.Coordinates 有九個參數 -> 箭頭尖點座標，直角的座標，文字接出去的座標，都有x, y, z
                    coor_to_arrow_dic[(round(object.Coordinates[0], 2), round(object.Coordinates[1], 2))] = (round(object.Coordinates[6], 2), round(object.Coordinates[7], 2))
                # 抓鋼筋本人和箍筋本人
                elif object.Layer == rebar_layer and object.ObjectName == 'AcDbPolyline':
                    # object.Coordinates 橫的和直的有四個參數 -> 兩端點的座標，都只有x, y; 彎的有八個參數 -> 直的端點，直的轉角，橫的轉角，橫的端點
                    if round(object.Length, 4) > 4: # 太短的是分隔線 -> 不要
                        if len(object.Coordinates) == 4 and round(object.Coordinates[1], 2) == round(object.Coordinates[3], 2): # 橫的 -> y 一樣 -> 鋼筋
                            coor_to_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4))) 
                        elif len(object.Coordinates) == 4 and round(object.Coordinates[0], 2) == round(object.Coordinates[2], 2): # 直的 -> x 一樣 -> 箍筋
                            coor_to_tie_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[2], 2), round(object.Coordinates[3], 2)), round(object.Length, 4))) 
                        elif len(object.Coordinates) == 8: # 彎的 -> 直的端點，橫的端點
                            coor_to_bend_rebar_list.append(((round(object.Coordinates[0], 2), round(object.Coordinates[1], 2)), (round(object.Coordinates[6], 2), round(object.Coordinates[7], 2)), round(object.Length, 4))) 
                # 抓箍筋文字座標
                elif object.Layer == tie_text_layer and object.ObjectName == 'AcDbText':
                    coor = (round(object.InsertionPoint[0], 2), round(object.InsertionPoint[1], 2))
                    coor_to_tie_text_list.append((object.TextString, coor))
                # 抓圖框
                elif object.Layer == block_layer and object.EntityName == "AcDbBlockReference":
                    coor1 = (round(object.GetBoundingBox()[0][0], 2), round(object.GetBoundingBox()[0][1], 2))
                    coor2 = (round(object.GetBoundingBox()[1][0], 2), round(object.GetBoundingBox()[1][1], 2))
                    coor_to_block_list.append(((coor1, coor2), {}, {}))
                # 抓梁的字的座標
                elif object.Layer == beam_text_layer and object.ObjectName == 'AcDbText':
                    midpoint = (round((object.GetBoundingBox()[0][0] + object.GetBoundingBox()[1][0]) / 2, 2), round((object.GetBoundingBox()[0][1] + object.GetBoundingBox()[1][1]) / 2, 2))
                    coor_to_beam_list.append((object.TextString, midpoint, [], {}))
                    
            flag = 1
        except Exception as e:
            error_count += 1
            time.sleep(5)
            error(f'read_beam error in step 7: {e}, error_count = {error_count}.')
    progress('梁配筋圖讀取進度 7/15', progress_file)

    # 在這之後就沒有while迴圈了，所以錯超過10次就出去
    if error_count > 10:
        try:
            doc_beam.Close(SaveChanges=False)
        except:
            pass
        return False

    # Step 8-15 是在處理鋼筋的部分
    
    # Step 8. 對應箭頭跟鋼筋
    new_coor_to_arrow_dic = {}
    for x in coor_to_arrow_dic: #此時的coor_to_arrow_dic為尖點座標->文字端坐標
        arrow_coor = x
        min_diff = 100 # 先看y是不是最近，再看x有沒有被夾到
        min_head_coor = ''
        min_tail_coor = ''
        min_length = ''
        min_mid_coor = ''
        for y in coor_to_rebar_list: # (頭座標，尾座標，長度)
            head_coor = y[0]
            tail_coor = y[1]
            mid_coor = (round((head_coor[0] + tail_coor[0]) / 2, 2), head_coor[1])
            length = y[2]
            y_diff = abs(mid_coor[1] - arrow_coor[1])
            if y_diff < min_diff and (head_coor[0] - arrow_coor[0]) * (tail_coor[0] - arrow_coor[0]) <= 0:
                min_diff = y_diff
                min_head_coor = head_coor
                min_tail_coor = tail_coor
                min_length = length
                min_mid_coor = mid_coor
        
        if min_head_coor != '':
            new_coor_to_arrow_dic[x] = (coor_to_arrow_dic[x], min_length, min_mid_coor) # 新的coor_to_arrow_dic為尖點座標 -> (文字端坐標，鋼筋長度，鋼筋中點座標)
            coor_to_rebar_list.remove((min_head_coor, min_tail_coor, min_length))
    
    coor_to_arrow_dic = new_coor_to_arrow_dic
    progress('梁配筋圖讀取進度 8/15', progress_file)
        
    # Step 9. 對應箭頭跟文字，並完成head_to_data_dic, tail_to_data_dic
    new_coor_to_arrow_dic = {}
    head_to_data_dic = {} # 座標 -> (number, size)
    tail_to_data_dic = {}
    for x in coor_to_arrow_dic: # 新的coor_to_arrow_dic為尖點座標 -> (文字端坐標，鋼筋長度，鋼筋中點座標)
        if len(coor_to_arrow_dic[x]) == 3:
            arrow_coor = coor_to_arrow_dic[x][0]
            length = coor_to_arrow_dic[x][1]
            rebar_mid_coor = coor_to_arrow_dic[x][2]
            min_diff = 100
            min_data = ''
            min_data_coor = ''
            for y in coor_to_data_list: # for 鋼筋的 (字串，座標)
                data = y[0]
                data_coor = y[1]
                x_diff = abs(arrow_coor[0] - data_coor[0])
                y_diff = abs(arrow_coor[1] - data_coor[1])
                total = x_diff + y_diff
                if total < min_diff:
                    min_diff = total
                    min_data = data
                    min_data_coor = data_coor
            if min_data != '':
                if '-' in min_data:
                    number = min_data.split('-')[0]
                    size =  min_data.split('-')[1]    
                    new_coor_to_arrow_dic[x] = (arrow_coor, length, rebar_mid_coor, number, size, min_data_coor) # 新的coor_to_arrow_dic為尖點座標 -> (箭頭文字端坐標，鋼筋長度，鋼筋中點座標，數量，尺寸，文字座標)
                    head_to_data_dic[(rebar_mid_coor[0] - length / 2, rebar_mid_coor[1])] = (number, size)
                    tail_to_data_dic[(rebar_mid_coor[0] + length / 2, rebar_mid_coor[1])] = (number, size)
                else:
                    error(f"There are no '-' in {min_data}. ")
    
    coor_to_arrow_dic = new_coor_to_arrow_dic
    progress('梁配筋圖讀取進度 9/15', progress_file)
    
    # Step 10. 統計目前的type跟size
    for x in coor_to_arrow_dic: # 新的coor_to_arrow_dic為尖點座標 -> (箭頭文字端坐標，鋼筋長度，鋼筋中點座標，數量，尺寸，文字座標)
        # 先找在哪個block裡面
        for y in coor_to_block_list: # ((左下，右上), rebar_length_dic, tie_count_dic)
            if x[0] > y[0][0][0] and x[0] < y[0][1][0] and x[1] > y[0][0][1] and x[1] < y[0][1][1]:
                # y[1] 是該格的rebar_length_dic: size -> length * number
                if coor_to_arrow_dic[x][4] not in y[1]:
                    y[1][coor_to_arrow_dic[x][4]] = float(coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
                else:
                    y[1][coor_to_arrow_dic[x][4]] += float(coor_to_arrow_dic[x][1]) * float(coor_to_arrow_dic[x][3])
    
    progress('梁配筋圖讀取進度 10/15', progress_file)
    # Step 11. 拿剩下的直的去找跟誰接在一起
    coor_to_rebar_list_straight_left = [] # (頭座標，尾座標，長度，number，size)
    coor_to_rebar_list_straight_right = []
    for x in coor_to_rebar_list: # (頭座標，尾座標，長度)
        if x[1] in head_to_data_dic: # 座標 -> (number, size)
            coor_to_rebar_list_straight_right.append((x[0], x[1], x[2], head_to_data_dic[x[1]][0], head_to_data_dic[x[1]][1]))
        elif x[0] in tail_to_data_dic:
            coor_to_rebar_list_straight_left.append((x[0], x[1], x[2], tail_to_data_dic[x[0]][0], tail_to_data_dic[x[0]][1]))
    
    progress('梁配筋圖讀取進度 11/15', progress_file)
    # Step 12. 拿彎的去找跟誰接在一起
    new_coor_to_bend_rebar_list = [] # 新的：(直的端點，橫的端點，長度，number，size)
    for x in coor_to_bend_rebar_list: # (直的端點，橫的端點，長度)
        if x[1] in head_to_data_dic:
            new_coor_to_bend_rebar_list.append((x[0], x[1], x[2], head_to_data_dic[x[1]][0], head_to_data_dic[x[1]][1]))
        elif x[1] in tail_to_data_dic:
            new_coor_to_bend_rebar_list.append((x[0], x[1], x[2], tail_to_data_dic[x[1]][0], tail_to_data_dic[x[1]][1]))
    coor_to_bend_rebar_list = new_coor_to_bend_rebar_list 
    
    progress('梁配筋圖讀取進度 12/15', progress_file)
    #Step 13. 直的彎的對對碰 (此處bug: 如果有畫不下的情況，會直接繼承前面的性質，不會去找沒畫完的在哪兒)
    # 左直找右直和右彎，剩的直接繼承前面的性質
    for x in coor_to_rebar_list_straight_left: # (頭座標，尾座標，長度，number，size)
        founded = 0
        for y in coor_to_rebar_list_straight_right: # (頭座標，尾座標，長度，number，size)
            if y[0][1] == x[0][1] and x[0][0] < y[0][0] and y[0][0] < x[1][0]: # 棒棒打棒棒
                if x[3] == y[3] and x[4] == y[4]:
                    # 先找在哪個block裡面
                    for z in coor_to_block_list: # ((左下，右上), rebar_length_dic, tie_count_dic)
                        if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                            # y[1] 是該格的rebar_length_dic: size -> length * number
                            if x[4] not in z[1]:
                                z[1][x[4]] = float(x[2]) * float(x[3])
                            else:
                                z[1][x[4]] += float(x[2]) * float(x[3])
                else:
                    error(f'{x[0]}, {x[1]}, {y[0]}, {y[1]}: 左右鋼筋不一致')
                founded = 1
                break
        if founded:
            coor_to_rebar_list_straight_right.remove(y)
        else:
            for y in coor_to_bend_rebar_list: # (直的端點，橫的端點，長度，number，size)
                if y[1][1] == x[0][1] and x[0][0] < y[1][0] and y[1][0] < x[1][0]: # 彎彎打棒棒
                    if x[3] >= y[3] and x[4] == y[4]:
                        # 先找在哪個block裡面
                        for z in coor_to_block_list: # ((左下，右上), rebar_length_dic, tie_count_dic)
                            if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                                # y[1] 是該格的rebar_length_dic: size -> length * number
                                if x[4] not in z[1]:
                                    z[1][x[4]] = (float(x[2]) * float(y[3]) + float(y[2]) * (float(x[3]) - float(y[3])))
                                else:
                                    z[1][x[4]] += (float(x[2]) * float(y[3]) + float(y[2]) * (float(x[3]) - float(y[3])))
                    else:
                        error(f'{x[0]}, {x[1]}, {y[0]}, {y[1]}: 左右鋼筋不一致或鋼筋量有誤')
                    founded = 1
                    break
            if founded:
                coor_to_bend_rebar_list.remove(y)
            else:
                error(f'warning: {x[0]}, {x[1]}: 只有單邊 -> 真的結束了or空間太小畫不下')
                # 先找在哪個block裡面
                for z in coor_to_block_list: # ((左下，右上), rebar_length_dic, tie_count_dic)
                    if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                        # y[1] 是該格的rebar_length_dic: size -> length * number
                        if x[4] not in z[1]:
                            z[1][x[4]] = float(x[2]) * float(x[3])
                        else:
                            z[1][x[4]] += float(x[2]) * float(x[3])
            
    # 右直找左彎，剩的直接繼承前面的性質
    for x in coor_to_rebar_list_straight_right: # (頭座標，尾座標，長度，number，size)
        founded = 0
        for y in coor_to_bend_rebar_list:
            if y[1][1] == x[0][1] and x[0][0] < y[1][0] and y[1][0] < x[1][0]: # 彎彎打棒棒
                if x[3] >= y[3] and x[4] == y[4]:
                    # 先找在哪個block裡面
                    for z in coor_to_block_list:
                        if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                            # y[1] 是該格的rebar_length_dic: size -> length * number
                            if x[4] not in z[1]:
                                z[1][x[4]] = (float(x[2]) * float(y[3]) + float(y[2]) * (float(x[3]) - float(y[3])))
                            else:
                                z[1][x[4]] += (float(x[2]) * float(y[3]) + float(y[2]) * (float(x[3]) - float(y[3])))
                else:
                    error(f'{x[0]}, {x[1]}, {y[0]}, {y[1]}: 左右鋼筋不一致或鋼筋量有誤')
                founded = 1
                break
        if founded:
            coor_to_bend_rebar_list.remove(y)
        else:
            error(f'warning: {x[0]}, {x[1]}: 只有單邊 -> 真的結束了or空間太小畫不下')
            # 先找在哪個block裡面
            for z in coor_to_block_list:
                if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                    # y[1] 是該格的rebar_length_dic: size -> length * number
                    if x[4] not in z[1]:
                        z[1][x[4]] = float(x[2]) * float(x[3])
                    else:
                        z[1][x[4]] += float(x[2]) * float(x[3])
    
    # 剩的彎直接繼承前面的性質
    for x in coor_to_bend_rebar_list: 
        error(f'warning: {x[0]}, {x[1]}: 只有單邊 -> 真的結束了or空間太小畫不下')
        # 先找在哪個block裡面
        for z in coor_to_block_list:
            if x[0][0] > z[0][0][0] and x[0][0] < z[0][1][0] and x[0][1] > z[0][0][1] and x[0][1] < z[0][1][1]:
                # y[1] 是該格的rebar_length_dic: size -> length * number
                if x[4] not in z[1]:
                    z[1][x[4]] = float(x[2]) * float(x[3])
                else:
                    z[1][x[4]] += float(x[2]) * float(x[3])
    
    # # DEBUG # 畫線把文字跟鋼筋中點連在一起
    # date = time.strftime("%Y-%m-%d", time.localtime())
    # layer_beam = doc_beam.Layers.Add(f"S-CLOUD_{date}")
    # doc_beam.ActiveLayer = layer_beam
    # layer_beam.color = 111
    # layer_beam.Linetype = "Continuous"
    
    # for x in coor_to_arrow_dic:
    #     rebar_mid_coor = coor_to_arrow_dic[x][2]
    #     data_coor = coor_to_arrow_dic[x][5]
    #     if rebar_mid_coor != '' and data_coor != '':
    #         coor_list = [rebar_mid_coor[0], rebar_mid_coor[1], 0, data_coor[0], data_coor[1], 0]
    #         points = vtFloat(coor_list)
    #         line = msp_beam.AddPolyline(points)
    #         line.SetWidth(0, 2, 2)
    
    progress('梁配筋圖讀取進度 13/15', progress_file)
    
    # Step 14-15 和 16 為箍筋部分，14-15在算框框內的數量，16在算每個梁的總長度，兩者獨立
        
    # Step 14. 算箍筋
    for x in coor_to_tie_text_list: # (字串，座標)
        if '-' in x[0] and int(x[0].split('-')[0]): # 已經算好有幾根就直接用
            count = int(x[0].split('-')[0])
            size = (x[0].split('-')[1]).split('@')[0] # 用'-'和'@'來切
            if size.split('#')[0].isdigit():
                count *= int(size.split('#')[0])
                size = f"#{size.split('#')[1]}"
            for y in coor_to_block_list:
                if x[1][0] > y[0][0][0] and x[1][0] < y[0][1][0] and x[1][1] > y[0][0][1] and x[1][1] < y[0][1][1]:
                    # y[2] 是該格的tie_count_dic: size -> number
                    if size not in y[2]:
                        y[2][size] = count
                    else:
                        y[2][size] += count

        else: # 沒算好自己算
            left_diff = 1000 # 找左邊最近的箍筋
            min_left_coor = ''
            min_right_coor = ''
            right_diff = 1000 # 找右邊最近的箍筋
            for y in coor_to_tie_list: # (下座標，上座標，長度) 
                if y[0][0] < x[1][0] and x[1][0] - y[0][0] < left_diff and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字左邊且diff最小且文字有被上下的y夾住
                    left_diff = x[1][0] - y[0][0]
                    min_left_coor = y[0]
                elif y[0][0] > x[1][0] and y[0][0] - x[1][0] < right_diff and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字右邊且diff最小且文字有被上下的y夾住
                    right_diff = y[0][0] - x[1][0]
                    min_right_coor = y[0]
            if left_diff != 1000 and right_diff != 1000:
                size = x[0].split('@')[0] # 用'@'來切
                bound = int(x[0].split('@')[1])
                count = int((left_diff + right_diff) / bound)
                if size.split('#')[0].isdigit():
                    count *= int(size.split('#')[0])
                    size = f"#{size.split('#')[1]}"
                for y in coor_to_block_list:
                    if x[1][0] > y[0][0][0] and x[1][0] < y[0][1][0] and x[1][1] > y[0][0][1] and x[1][1] < y[0][1][1]:
                        # y[2] 是該格的tie_count_dic: size -> number
                        if size not in y[2]:
                            y[2][size] = count
                        else:
                            y[2][size] += count
                
                # DEBUG # 畫線把文字跟左右的線連在一起
                coor_list1 = [min_left_coor[0], min_left_coor[1], 0, x[1][0], x[1][1], 0]
                coor_list2 = [min_right_coor[0], min_right_coor[1], 0, x[1][0], x[1][1], 0]
                points1 = vtFloat(coor_list1)
                points2 = vtFloat(coor_list2)
                line1 = msp_beam.AddPolyline(points1)
                line2 = msp_beam.AddPolyline(points2)
                line1.SetWidth(0, 2, 2)
                line2.SetWidth(0, 2, 2)
                line1.color = 101
                line2.color = 101

    progress('梁配筋圖讀取進度 14/15', progress_file)
    # Step 15. 印出每個框框的結果然後加在一起
    rebar_length_dic = {}
    tie_count_dic = {}
    f = open(result_filename, "w", encoding = 'utf-8')
    
    for x in coor_to_block_list:
        if len(x[1]) != 0 or len(x[2]) != 0:
            f.write(f'統計左下角為{x[0][0]}，右上角為{x[0][1]}的框框內結果：\n')
            if len(x[1]) != 0:
                f.write('鋼筋計算：\n')
                for y in x[1]:
                    f.write(f'{y}: 總長度(長度*數量)為 {x[1][y]}\n')
                    if y in rebar_length_dic:
                        rebar_length_dic[y] += x[1][y]
                    else:
                        rebar_length_dic[y] = x[1][y]
            else:
                f.write('此圖框內沒有鋼筋\n')
                
            if len(x[2]) != 0:    
                f.write('箍筋計算：\n')
                for y in x[2]:
                    f.write(f'{y}: 總數量為 {x[2][y]}\n')
                    if y in tie_count_dic:
                        tie_count_dic[y] += x[2][y]
                    else:
                        tie_count_dic[y] = x[2][y]
            else:
                f.write('此圖框內沒有箍筋\n')
                
            f.write('\n')
                
    f.write(f'統計所有結果：\n')
    f.write('鋼筋計算：\n')
    for y in rebar_length_dic:
        f.write(f'{y}: 總長度(長度*數量)為 {rebar_length_dic[y]}\n')
        
    f.write('箍筋計算：\n')
    for y in tie_count_dic:
        f.write(f'{y}: 總數量為 {tie_count_dic[y]}\n')
    
    f.close
    
    progress('梁配筋圖讀取進度 15/15', progress_file)
    
    # Step 16. 把箍筋跟beam字串綁在一起
    # 先判斷beam字串在上面還是下面 -> 看最高的beam字串跟tie_text誰比較高
    highest = 0
    down = 1 # 預設在下
    for x in coor_to_tie_text_list: # (字串，座標)
        if x[1][1] > highest:
            highest = x[1][1]
    for x in coor_to_beam_list: # (string, midpoint, list of tie)
        if x[1][1] > highest:
            down = 0
            break
    for x in coor_to_tie_text_list: # (字串，座標)
        # 先算出tie的根數和尺寸
        if '-' in x[0] and int(x[0].split('-')[0]): # 已經算好有幾根就直接用
            count = int(x[0].split('-')[0])
            size = (x[0].split('-')[1]).split('@')[0] # 用'-'和'@'來切
            if size.split('#')[0].isdigit():
                count *= int(size.split('#')[0])
                size = f"#{size.split('#')[1]}"

        else: # 沒算好自己算
            left_diff = 1000
            right_diff = 1000
            for y in coor_to_tie_list: # (下座標，上座標，長度) 
                if y[0][0] < x[1][0] and x[1][0] - y[0][0] < left_diff and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字左邊且diff最小且文字有被上下的y夾住
                    left_diff = x[1][0] - y[0][0]
                elif y[0][0] > x[1][0] and y[0][0] - x[1][0] < right_diff and y[0][1] < x[1][1] and x[1][1] < y[1][1]: # 箍筋在文字右邊且diff最小且文字有被上下的y夾住
                    right_diff = y[0][0] - x[1][0]
            if left_diff != 1000 and right_diff != 1000:
                size = x[0].split('@')[0] # 用'@'來切
                bound = int(x[0].split('@')[1])
                count = int((left_diff + right_diff) / bound)
                if size.split('#')[0].isdigit():
                    count *= int(size.split('#')[0])
                    size = f"#{size.split('#')[1]}"
                    
        min_diff = 1000
        min_string = ''
        for y in coor_to_beam_list:
            if (down == 1 and x[1][1] > y[1][1]) or (down == 0 and x[1][1] < y[1][1]):
                x_diff = abs(x[1][0] - y[1][0])
                y_diff = abs(x[1][1] - y[1][1])
                if x_diff + y_diff < min_diff:
                    min_diff = x_diff + y_diff
                    min_string = y[0]
        if min_string != '':
            for y in coor_to_beam_list:
                if y[0] == min_string:
                    y[2].append(x[0])
                    if size not in y[3]:
                        y[3][size] = count
                    else:
                        y[3][size] += count
    
        f = open(tie_file, "w", encoding = 'utf-8')  
        tie_length_dic = {}
        for x in coor_to_beam_list: # (string, midpoint, list of tie)
            try:
                f.write(f'{x[0]}的箍筋列表：\n')  
                for y in x[2]:
                    f.write(f'  {y}\n')
                f.write(f'統計：\n')  
                for y in x[3]:
                    f.write(f'  {y}: 總數量為 {x[3][y]}\n')
                size = x[0].replace('X', 'x')
                size = (size.split('(')[1]).split(')')[0]
                num1 = int(size.split('x')[0]) - 10
                num2 = int(size.split('x')[1]) - 10
                total_len = (num1 + num2) * 2
                for y in x[3]:
                    ans = x[3][y] * total_len
                    f.write(f'  {y}: 總長度為 {ans}\n')
                    if y in tie_length_dic:
                        tie_length_dic[y] += ans
                    else:
                        tie_length_dic[y] = ans
            except:
                pass
                
    f.write(f'統計所有結果：\n')
    f.write('箍筋計算：\n')
    for y in tie_length_dic:
        f.write(f'{y}: 總長度(長度*數量)為 {tie_length_dic[y]}\n')
    
    f.close
    progress('梁配筋圖讀取完成', progress_file)
    return

error_file = './result/error_log.txt' # error_log.txt的路徑

if __name__=='__main__':

    # 檔案路徑區
    # 跟AutoCAD有關的檔案都要吃絕對路徑
    beam_filename = r"K:\100_Users\EI 202208 Bamboo\BeamQC\task27-CD\XS-BEAM.dwg"#sys.argv[1] # XS-BEAM的路徑
    progress_file = './result/tmp'#sys.argv[14]
    rebar_file = './result/rebar.txt' # rebar.txt的路徑 -> 計算鋼筋和箍筋總量
    tie_file = './result/tie.txt' # rebar.txt的路徑 -> 把箍筋跟梁綁在一起

    # 在beam裡面自訂圖層
    rebar_data_layer = 'S-LEADER' # 箭頭和鋼筋文字的塗層
    rebar_layer = 'S-REINF' # 鋼筋和箍筋的線的塗層
    tie_text_layer = 'S-TEXT' # 箍筋文字圖層
    block_layer = '0' # 框框的圖層
    beam_text_layer = 'S-RC' # 梁的字串圖層

    read_beam(beam_filename, rebar_layer, rebar_data_layer, tie_text_layer, block_layer, beam_text_layer, rebar_file, tie_file, progress_file)