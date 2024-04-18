import logging
import win32com.client
import time
import pythoncom

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
main_logger = logging.getLogger(__name__)


def vtFloat(l):
    '''
    要把點座標組成的list轉成autocad看得懂的樣子
    '''
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)


def open_new_joint_plan():

    draw_plan_logger = main_logger.getChild('open_new_joint_plan')
    wincad_acad = None
    error_count = 0
    while not wincad_acad:
        if error_count > 10:
            break
        try:
            wincad_acad = win32com.client.Dispatch("AutoCAD.Application")
            if not wincad_acad.IsEarlyBind:  # 判断是否是EarlyBind，如果不是则打开Earlybind模式
                wincad_acad.TurnOnEarlyBind()
        except Exception as ex:
            error_count += 1
            draw_plan_logger.error(
                f'Autocad Open Fail, times = {error_count} , {ex}')
            time.sleep(5)

    draw_plan_logger.info('Open Autocad Success')

    doc_joint_plan = None
    error_count = 0
    while not doc_joint_plan:
        if error_count > 10:
            break
        try:
            # doc_joint_plan = wincad_joint_plan.Documents.Open(plan_filename)
            doc_joint_plan = wincad_acad.Documents.Add()
        except Exception as ex:
            error_count += 1
            draw_plan_logger.error(
                f'Autocad Add Fail, times = {error_count} , {ex}')
            time.sleep(5)

    draw_plan_logger.info('Add Plan Success')

    msp_joint_plan = None
    error_count = 0
    while not msp_joint_plan:
        if error_count > 10:
            break
        try:
            msp_joint_plan = doc_joint_plan.Modelspace
        except Exception as ex:
            error_count += 1
            draw_plan_logger.error(
                f'Autocad Open Modelspace Fail, times = {error_count} , {ex}')
            time.sleep(5)

    draw_plan_logger.info('Open Modelspace Success')

    return doc_joint_plan, msp_joint_plan


def setup_drawing_layer(doc_joint_plan, msp_joint_plan, layer_config: dict[str, dict]):
    '''
    layer_config = {
        "Name":{
            "ColorIndex"="",
            "Linetype" = "",
            "Lineweight" = "",
        }
    }
    '''
    setup_plan_logger = main_logger.getChild('setup_drawing_layer')
    linetypes = doc_joint_plan.Linetypes
    text_styles = doc_joint_plan.TextStyles
    new_style = text_styles.Add("myStandard")
    new_style.fontFile = "simplex.shx"
    new_style.BigFontFile = "lsp.shx"
    # doc_joint_plan.LinetypeScale = 100
    error_count = 0
    while msp_joint_plan:
        if error_count > 10:
            break
        try:
            for layer_name, layer in layer_config.items():
                doc_layer = doc_joint_plan.Layers.Add(layer_name)
                try:
                    linetypes.Item(layer["Linetype"])
                except pythoncom.com_error:
                    linetypes.Load(layer["Linetype"], "acad.lin")
                doc_layer.color = layer["ColorIndex"]
                doc_layer.Linetype = layer["Linetype"]
                doc_layer.Lineweight = layer["Lineweight"]
            break
        except Exception as ex:
            error_count += 1
            setup_plan_logger.error(
                f'Layer {layer} Add Fail, times = {error_count} , {ex}')
            time.sleep(5)

    setup_plan_logger.info('Layers Add Success')


def draw_beam_line(msp_joint_plan, mline_list, layer_config: dict[str, dict]):
    import joint_scan
    mline: joint_scan.MlineObject
    draw_beam_line_logger = main_logger.getChild('draw_beam_line')
    for mline in mline_list:
        draw_beam_line_logger.debug(f"Add {mline.beam_serial}")

        if not mline.beam_serial:
            continue
        if mline.xy_direction == "x":
            points = [mline.start[0], mline.mid[1],
                      0, mline.end[0], mline.mid[1], 0]
        else:
            points = [mline.mid[0], mline.start[1],
                      0, mline.mid[0], mline.end[1], 0]
        ms_line = msp_joint_plan.AddMLine(vtFloat(points))
        ms_line.Layer = "Beam"
        ms_line.MLineScale = mline.scale
        ms_line.LinetypeScale = 100
        ms_line.Update()

        points = [mline.mid[0], mline.mid[1], 0]
        ms_beam_text = msp_joint_plan.AddText(
            mline.beam_serial[1], vtFloat(points), 30)
        ms_beam_text.Rotation = mline.beam_serial[2]
        ms_beam_text.Layer = "BeamText"
        ms_beam_text.Alignment = 1
        ms_beam_text.TextAlignmentPoint = vtFloat(points)
        # ms_beam_text.VerticalAlignment = 2
        # ms_beam_text.InsertionPoint = vtFloat(points)

        if mline.left_column:
            ms_polyline = msp_joint_plan.AddPolyline(
                vtFloat(mline.left_column.get_corner()))
            ms_polyline.Layer = "Column"
            points = [(mline.left_column.start[0] + mline.left_column.end[0]) / 2,
                      (mline.left_column.start[1] +
                       mline.left_column.end[1]) / 2,
                      0]
            ms_column_text = msp_joint_plan.AddText(
                mline.left_column.column_serial[1], vtFloat(points), 20)
            ms_column_text.Layer = "ColumnText"
            ms_column_text.Alignment = 13
            ms_column_text.TextAlignmentPoint = vtFloat(points)

        if mline.right_column:
            ms_polyline = msp_joint_plan.AddPolyline(
                vtFloat(mline.right_column.get_corner()))
            ms_polyline.Layer = "Column"
            points = [(mline.right_column.start[0] + mline.right_column.end[0]) / 2,
                      (mline.right_column.start[1] +
                       mline.right_column.end[1]) / 2,
                      0]
            ms_column_text = msp_joint_plan.AddText(
                mline.right_column.column_serial[1], vtFloat(points), 20)
            ms_column_text.Layer = "ColumnText"
            ms_column_text.Alignment = 13
            ms_column_text.TextAlignmentPoint = vtFloat(points)

    draw_beam_line_logger.info('Beam Create Success')


def draw_rebar_data(msp_joint_plan, mline_list):
    # Constants for alignment types (AcAlignmentEnum):
    #   0 = Left, 1 = Center, 2 = Right, 3 = Aligned, 4 = Middle,
    #   5 = Fit, 6 = Top Left, 7 = Top Center, 8 = Top Right,
    #   9 = Middle Left, 10 = Middle Center, 11 = Middle Right,
    #   12 = Bottom Left, 13 = Bottom Center, 14 = Bottom Right
    import joint_scan
    mline: joint_scan.MlineObject
    alignment_dict = {
        ("top", "left"): 6,
        ("top", "middle"): 7,
        ("top", "right"): 8,
        ("bottom", "left"): 12,
        ("bottom", "middle"): 13,
        ("bottom", "right"): 14,
    }
    draw_rebar_data_logger = main_logger.getChild('draw_rebar_data')
    for mline in mline_list:
        if mline.beam_data:
            beam_data = mline.beam_data
            if mline.xy_direction == "x":
                for pos1, rebars in [("top", beam_data.rebar_table["top"]), ("bottom", beam_data.rebar_table["bottom"])]:
                    for pos2, rebar in rebars.items():
                        if pos1 == "top":
                            prefix = "上層:"
                            y = mline.mid[1] + mline.scale / 2
                        if pos1 == "bottom":
                            prefix = "下層:"
                            y = mline.mid[1] - mline.scale / 2
                        if pos2 == "left":
                            x = mline.start[0]
                        if pos2 == "middle":
                            x = mline.mid[0]
                        if pos2 == "right":
                            x = mline.end[0]
                        ms_beam_text = msp_joint_plan.AddText(
                            prefix + "+".join([r.text for r in rebar]), vtFloat([x, y, 0]), 10)
                        ms_beam_text.Alignment = alignment_dict[(pos1, pos2)]
                        ms_beam_text.TextAlignmentPoint = vtFloat([x, y, 0])
                        ms_beam_text.StyleName = "myStandard"
                        ms_beam_text.Layer = "RebarText"
            if mline.xy_direction == "y":
                for pos1, rebars in [("top", beam_data.rebar_table["top"]), ("bottom", beam_data.rebar_table["bottom"])]:
                    for pos2, rebar in rebars.items():
                        if pos1 == "top":
                            prefix = "上層:"
                            x = mline.mid[0] + mline.scale / 2
                        if pos1 == "bottom":
                            prefix = "下層:"
                            x = mline.mid[0] - mline.scale / 2
                        if pos2 == "left":
                            y = mline.start[1]
                        if pos2 == "middle":
                            y = mline.mid[1]
                        if pos2 == "right":
                            y = mline.end[1]
                        ms_beam_text = msp_joint_plan.AddText(
                            prefix + "+".join([r.text for r in rebar]), vtFloat([x, y, 0]), 10)
                        ms_beam_text.Alignment = alignment_dict[(pos1, pos2)]
                        ms_beam_text.Rotation = mline.beam_serial[2]
                        ms_beam_text.TextAlignmentPoint = vtFloat([x, y, 0])
                        ms_beam_text.StyleName = "myStandard"
                        ms_beam_text.Layer = "RebarText"
            if mline.left_column and mline.left_column.column_data:
                ms_column_dcr_text = msp_joint_plan.AddText(f"DCR:{mline.left_column.column_data.DCR}",
                                                            vtFloat([x, y, 0]), 10)
                ms_column_dcr_text.Alignment = 7
                ms_column_dcr_text.TextAlignmentPoint = vtFloat(
                    [*mline.left_column.mid])
                ms_column_dcr_text.StyleName = "myStandard"
                ms_column_dcr_text.Layer = "RebarText"

    draw_rebar_data_logger.info("Rebar Data Create Success")


def save_new_file(doc_joint_plan, plan_filename):
    doc_joint_plan.SaveAs(plan_filename)
    doc_joint_plan.Close(SaveChanges=True)
    main_logger.info(f"{plan_filename} save Success")


def create_joint_plan_view(plan_filename, mline_list, layer_config):
    doc_joint_plan, msp_joint_plan = open_new_joint_plan()

    setup_drawing_layer(doc_joint_plan=doc_joint_plan,
                        msp_joint_plan=msp_joint_plan,
                        layer_config=layer_config)

    draw_beam_line(msp_joint_plan=msp_joint_plan,
                   mline_list=mline_list,
                   layer_config=layer_config)

    draw_rebar_data(msp_joint_plan=msp_joint_plan,
                    mline_list=mline_list)

    save_new_file(doc_joint_plan=doc_joint_plan,
                  plan_filename=plan_filename)


if __name__ == "__main__":
    doc_joint_plan, msp_joint_plan = open_new_joint_plan()

    setup_drawing_layer(doc_joint_plan=doc_joint_plan,
                        msp_joint_plan=msp_joint_plan,
                        layer_config={
                            "Beam": {
                                "ColorIndex": 2,
                                "Linetype": "HIDDEN",
                                "Lineweight": 0.5
                            }
                        })
