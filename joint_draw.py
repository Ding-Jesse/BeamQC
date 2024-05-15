import logging
import win32com.client
import time
import pythoncom
from math import sqrt
from collections import defaultdict
from logger import setup_custom_logger
# logging.basicConfig(level=logging.DEBUG,
#                     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# main_logger = logging.getLogger(__name__)
global main_logger
# main_logger = setup_custom_logger(__name__)
split_line = "\n"


def vtFloat(l):
    '''
    要把點座標組成的list轉成autocad看得懂的樣子
    '''
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, l)


def open_new_joint_plan():
    pythoncom.CoInitialize()
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
    visited = set()
    for count, mline in enumerate(mline_list):
        if count % 500 == 0:
            draw_beam_line_logger.info(f"繪製梁{count} / {len(mline_list)}")

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
            mline.beam_serial[1], vtFloat(points), 20)
        ms_beam_text.Rotation = mline.beam_serial[2]
        ms_beam_text.Layer = "BeamText"
        ms_beam_text.Alignment = 10
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
            ms_column_text.StyleName = "myStandard"
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
            ms_column_text.StyleName = "myStandard"
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
        ("bottom", "left"): 0,
        ("bottom", "middle"): 1,
        ("bottom", "right"): 2,
    }

    draw_rebar_data_logger = main_logger.getChild('draw_rebar_data')
    for mline in mline_list:
        if mline.beam_data:
            beam_data = mline.beam_data
            beam_section = f"({mline.beam_data.width}X{mline.beam_data.depth})"
            beam_section_text_rotation = 0
            # Add Beam Rebar Data
            if mline.xy_direction == "x":

                section_x = mline.mid[0]
                section_y = mline.mid[1] - mline.scale / 2

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
                        if alignment_dict[(pos1, pos2)]:
                            ms_beam_text.TextAlignmentPoint = vtFloat(
                                [x, y, 0])
                        ms_beam_text.StyleName = "myStandard"
                        ms_beam_text.Layer = "RebarText"
            if mline.xy_direction == "y":

                section_x = mline.mid[0] + mline.scale / 2
                section_y = mline.mid[1]
                beam_section_text_rotation = mline.beam_serial[2]

                for pos1, rebars in [("top", beam_data.rebar_table["top"]), ("bottom", beam_data.rebar_table["bottom"])]:
                    for pos2, rebar in rebars.items():
                        if pos1 == "top":
                            prefix = "上層:"
                            x = mline.mid[0] - mline.scale / 2
                        if pos1 == "bottom":
                            prefix = "下層:"
                            x = mline.mid[0] + mline.scale / 2
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
                        if alignment_dict[(pos1, pos2)]:
                            ms_beam_text.TextAlignmentPoint = vtFloat(
                                [x, y, 0])
                        ms_beam_text.StyleName = "myStandard"
                        ms_beam_text.Layer = "RebarText"
            # Add Beam Size Data
            ms_beam_section_text = msp_joint_plan.AddText(
                beam_section, vtFloat([section_x, section_y, 0]), 10)
            ms_beam_section_text.Alignment = 7
            ms_beam_section_text.TextAlignmentPoint = vtFloat(
                [section_x, section_y, 0])
            ms_beam_section_text.StyleName = "myStandard"
            ms_beam_section_text.Layer = "BeamText"
            ms_beam_section_text.Rotation = beam_section_text_rotation

    draw_rebar_data_logger.info("Rebar Data Create Success")


def draw_column_data(msp_joint_plan, column_block_list):
    from joint_scan import ColumnBlock, UserDefineWarning
    column_block: ColumnBlock
    draw_column_data = main_logger.getChild('draw_column_data')
    summary_dict: dict[str, list] = {}
    for column_block in column_block_list:
        if column_block.column_data:

            points = [*column_block.mid]
            for pos, result in column_block.column_data.joint_result.items():
                ms_column_dcr_text = msp_joint_plan.AddMText(vtFloat(points),
                                                             10,
                                                             f"{pos}:DCR={result['DCR']}{split_line}{pos}:Code={result['design_code']}")
                ms_column_dcr_text.AttachmentPoint = 2
                ms_column_dcr_text.Height = 8
                ms_column_dcr_text.StyleName = "myStandard"
                ms_column_dcr_text.Layer = "ColumnText"
                ms_column_dcr_text.InsertionPoint = vtFloat(points)
                points[1] -= 30
                if result['DCR'] > 1:
                    column_block.warning.append(
                        UserDefineWarning.ColumnJointShearFail)
                if column_block.column_data.floor not in summary_dict:
                    summary_dict[column_block.column_data.floor] = []
                summary_dict[column_block.column_data.floor].append(
                    result['DCR'])

    draw_column_data.info("Column Data Create Success")
    return summary_dict


def draw_warning_plan(msp_joint_plan, column_block_list):
    from joint_scan import ColumnBlock
    column_block: ColumnBlock
    for column_block in column_block_list:
        if column_block.warning:
            radius = sqrt((column_block.end[0] - column_block.start[0])
                          ** 2 + (column_block.end[1] - column_block.start[1]) ** 2)
            ms_circle = msp_joint_plan.AddCircle(
                vtFloat([*column_block.mid]), radius * 1.2)
            ms_circle.Layer = "Warning"
            ms_circle.LinetypeScale = 100


def change_block_ratio(block_list, ratio):
    def inner_point(pt1, pt2, m, n):
        x1, y1, z1 = pt1
        x2, y2, z2 = pt2
        return ((x2 * m + x1 * n) / (m + n), (y2 * m + y1 * n) / (m + n), 0)
    visited = set()
    for floor, block_data in block_list:
        block, floor_text = block_data
        if floor_text in visited:
            continue
        visited.add(floor_text)

        mid = [(block[0][0] + block[1][0]) / 2,
               (block[0][1] + block[1][1]) / 2, 0]
        new_block = (inner_point(block[0], mid, 0.5 - ratio / 2, ratio / 2),
                     inner_point(mid, block[1], ratio / 2, 0.5 - ratio / 2))
        block_data[0] = new_block


def draw_block(msp_joint_plan, block_list, summary_dict):
    plot_table = defaultdict(list)
    for floor, block_data in block_list:
        block, floor_text = block_data

        points = [*block[0],
                  block[0][0], block[1][1], 0,
                  *block[1],
                  block[1][0], block[0][1], 0,
                  *block[0]]
        ms_block = msp_joint_plan.AddPolyline(vtFloat(points))
        ms_block.Layer = "Block"

        ms_floor_text = msp_joint_plan.AddText(
            floor_text, vtFloat([*block[0]]), 300)
        ms_floor_text.Layer = "Block"

        # plot joint summary
        if floor in summary_dict:
            joint_list = summary_dict[floor]
            fail_list = [dcr for dcr in joint_list if dcr >= 1]
            plot_table[block].append(
                f"DCR >= 1{split_line}{floor_text}:{len(fail_list)}{split_line}DCR < 1{split_line}{floor_text}:{len(joint_list) - len(fail_list)}")
    for block, text_list in plot_table.items():
        points = [block[1][0], block[0][1], 0]
        for text in text_list:
            ms_column_dcr_text = msp_joint_plan.AddMText(
                vtFloat(points), 500, text)
            ms_column_dcr_text.Height = 60
            ms_column_dcr_text.AttachmentPoint = 9
            ms_column_dcr_text.InsertionPoint = vtFloat(points)
            ms_column_dcr_text.Layer = "Warning"
        points[1] += 300


def save_new_file(doc_joint_plan, plan_filename):
    error_count = 0
    while error_count < 3:
        try:
            doc_joint_plan.SaveAs(plan_filename)
            doc_joint_plan.Close(SaveChanges=True)
            main_logger.info(f"{plan_filename} save Success")
            break
        except pythoncom.com_error:
            main_logger.info(f"saving plan file:{plan_filename} error")
            error_count += 1
            time.sleep(3)


def create_joint_plan_view(plan_filename,
                           mline_list,
                           column_block_list,
                           block_list,
                           layer_config, client_id):
    global main_logger
    main_logger = setup_custom_logger(__name__, client_id)
    doc_joint_plan, msp_joint_plan = open_new_joint_plan()

    setup_drawing_layer(doc_joint_plan=doc_joint_plan,
                        msp_joint_plan=msp_joint_plan,
                        layer_config=layer_config)

    draw_beam_line(msp_joint_plan=msp_joint_plan,
                   mline_list=mline_list,
                   layer_config=layer_config)

    draw_rebar_data(msp_joint_plan=msp_joint_plan,
                    mline_list=mline_list)

    summary_dict = draw_column_data(msp_joint_plan=msp_joint_plan,
                                    column_block_list=column_block_list)

    change_block_ratio(block_list=block_list, ratio=0.6)

    draw_warning_plan(msp_joint_plan=msp_joint_plan,
                      column_block_list=column_block_list)

    draw_block(msp_joint_plan=msp_joint_plan,
               block_list=block_list, summary_dict=summary_dict)

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
