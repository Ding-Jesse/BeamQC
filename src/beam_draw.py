import ezdxf
import ezdxf.path
import os
import src.save_temp_file as save_temp_file
from item.beam import Beam, RebarType
from ezdxf import readfile
from ezdxf.document import Drawing
from ezdxf.layouts.layout import Modelspace
from ezdxf.enums import TextEntityAlignment
from ezdxf.math import Vec2
from typing import Literal
from ezdxf.tools.standards import linetypes  # some predefined linetypes
from itertools import product


def draw_beam_rebar_dxf(output_folder: str = r'TEST\2024-1008',
                        beam_list: list = None,
                        dxf_file_name: str = None,
                        beam_tmp_file: str = r'TEST\2024-1008\梁\2024-1008-20241009_165315-beam-object.pkl'):
    '''
    Draw Beam Rebar Dxf , if beam list then use beam list , else use temp file
    '''
    if beam_list is None:
        if os.path.exists(beam_tmp_file):
            beam_list = save_temp_file.read_temp(beam_tmp_file)
        else:
            return
    # Create a new DXF document
    doc = ezdxf.new()

    init_doc_layers(doc=doc)

    # Create a new drawing in the document
    msp = doc.modelspace()
    # beam_list = [beam for beam in beam_list if beam.serial in ['FB9b']]
    for b in beam_list:
        try:
            draw_data = draw_beam(b)

            draw_text(draw_data['text'], msp)

            draw_polyline(draw_data['polyline'], msp)

            draw_dim(draw_data['dim'], msp)
        except ZeroDivisionError:
            pass
    # Save the DXF file
    if dxf_file_name is None:
        dxf_file_name = f'{output_folder}\\redraw-{os.path.splitext(os.path.basename(beam_tmp_file))[0]}.dxf'
    doc.saveas(dxf_file_name)
    return dxf_file_name


def init_doc_layers(doc: Drawing):
    doc.linetypes.add(
        name="DASHED",
        pattern=[0.2, 0.2, -0.2],
        description="DASHED",

    )
    layers = [
        {
            'name': 'Page',
            'color': 0
        },
        {
            'name': 'S-RC',
            'color': 130
        },
        {
            'name': 'S-REINFD',
            'color': 6
        },
        {
            'name': 'S-REINF',
            'color': 5
        },
        {
            'name': 'S-TEXT',
            'color': 2
        },
        {
            'name': 'S-TABLE',
            'color': 1
        },
        {
            'name': 'S-REINFH',
            'color': 1
        },
        {
            'name': 'S-LINE',
            'color': 34,
            'linetype': "DASHED",
        },
        {
            'name': 'S-DIM',
            'color': 1
        },

    ]
    for layer in layers:
        doc.layers.add(**layer)

    doc.styles.new("myStandard", dxfattribs={
                   "font": "simplex.shx", "bigfont": "lsp.shx"})
    dimstyle = doc.dimstyles.new(
        'MyDimStyle',  # Dimension style name
        dxfattribs={
            'dimclrd': 1,  # Dimension line color
            'dimclre': 2,  # Extension line color
            'dimexe': 0.5,  # Extension line extension beyond dimension line
            'dimexo': 0.25,  # Extension line offset from origin
            'dimasz': 5,  # Arrow size
            'dimtxt': 12,  # Text height
            'dimtxsty': 'myStandard',  # Text style
            'dimclrt': 2
            # Add more dimension style parameters as needed
        }
    )
    mleaderstyle = doc.mleader_styles.duplicate_entry("Standard", "myStandard")
    mleaderstyle.set_mtext_style("myStandard")
    mleaderstyle.dxf.char_height = 12  # set the default char height of MTEXT


def draw_beam(b: Beam):
    b.sort_rebar_table()
    text_list = []
    polyline_list: list[tuple] = []
    bounding_box_list: list[tuple] = []
    rebar_list: list[tuple] = []
    dim_list: list[tuple] = []
    dim_line_list = []
    tie_list = []
    middle_tie_list = []
    # Name
    text_list.append(
        (f'{b.floor} {b.serial} ({b.width}x{b.depth})', b.get_coor(), 20))
    # Block
    polyline_list.append([(b.start_pt.x, b.bot_y - 6),
                          (b.start_pt.x, b.top_y + 6),
                          (b.end_pt.x, b.top_y + 6),
                          (b.end_pt.x, b.bot_y - 6),
                          (b.start_pt.x, b.bot_y - 6)])
    bounding_box = b.get_bounding_box()
    bounding_box_list.append([bounding_box[0],
                              (bounding_box[0][0], bounding_box[1][1]),
                              bounding_box[1],
                              (bounding_box[1][0], bounding_box[0][1]),
                              bounding_box[0]])
    # Rebar
    for pos, pos2 in product([RebarType.Top.value, RebarType.Bottom.value],
                             [RebarType.Left.value, RebarType.Middle.value, RebarType.Right.value]):
        for i, rebar in enumerate(b.rebar_table[pos][pos2], start=0):
            text_list.append((rebar.text, rebar.arrow_coor[0]))
            rebar_list.append([(rebar.start_pt.x, rebar.start_pt.y),
                               (rebar.end_pt.x, rebar.end_pt.y)])
            mid_x = (rebar.start_pt.x + rebar.end_pt.x) / 2
            # dim_line_list.append(
            #     [(mid_x, rebar.start_pt.y), rebar.arrow_coor[0]])

            if b.rebar_table[f'{pos}_length'][pos2] and i == 0:
                dim_list.append(((rebar.start_pt.x, rebar.start_pt.y),
                                (rebar.end_pt.x, rebar.end_pt.y),
                                str(b.rebar_table[f'{pos}_length'][pos2][0]),
                                pos, 30))
    # Tie
    for pos in [RebarType.Left.value, RebarType.Middle.value, RebarType.Right.value]:
        if b.tie[pos]:
            rebar = b.tie[pos]
            text_list.append(
                (rebar.text, (rebar.start_pt.x, rebar.start_pt.y)))

    for middle_tie in b.middle_tie:
        text_list.append((middle_tie.text, middle_tie.arrow_coor[0]))
        middle_tie_list.append([(middle_tie.start_pt.x, middle_tie.start_pt.y),
                                (middle_tie.start_pt.x + middle_tie.length, middle_tie.end_pt.y)])
    # Dim

    return {
        'text': {
            'S-TEXT': text_list},
        'polyline': {
            'S-REINFD': rebar_list,
            'S-REINF': tie_list,
            'S-REINFH': middle_tie_list,
            'S-RC': polyline_list,
            'S-DIM': dim_line_list,
            'Page': bounding_box_list},
        'dim': {
            'S-DIM': dim_list}
    }


def draw_text(text_dict: dict, msp: Modelspace):
    '''
    text_list = [(content , insert_point , text_height)]
    '''
    for layer, text_list in text_dict.items():
        for text in text_list:
            if len(text) == 3:
                text_content, text_insert, text_height = text
            else:
                text_height = 8
                text_content, text_insert = text
            text_dxf = msp.add_text(text_content, dxfattribs={
                                    'height': text_height, 'layer': layer, "style": "myStandard"})
            # Set the insertion point and alignment
            text_dxf.set_placement(
                text_insert, align=TextEntityAlignment.MIDDLE_CENTER)


def draw_polyline(polyline_dict: dict,
                  msp: Modelspace):
    for layer, ply_list in polyline_dict.items():
        for ply in ply_list:
            msp.add_lwpolyline(ply, dxfattribs={"layer": layer})


def draw_dim(dim_dict: dict[str, list], msp: Modelspace):
    direction: Literal["top", "bottom", "left", "right"]
    for layer, dim_list in dim_dict.items():
        for i, dim in enumerate(dim_list):
            # spacing = 35 * ((i % 2) + 1) + 15 * (i % 2)
            # spacing = 35
            p1, p2, text, direction, spacing = dim
            if direction == "top":
                angle = 0
                text_rotation = 0
                base = (p1[0], p1[1] + spacing)
            if direction == "bottom":
                angle = 0
                text_rotation = 0
                base = (p1[0], p1[1] - spacing)
            if direction == "left":
                angle = 90
                base = (p1[0] - spacing, p1[1])
                text_rotation = 90
            if direction == "right":
                angle = 90
                base = (p1[0] + spacing, p1[1])
                text_rotation = 90
            # Add a linear dimension
            d = msp.add_linear_dim(
                base=base,  # Start point
                p1=p1,    # First extension line point
                p2=p2,    # Second extension line point
                text=text,
                angle=angle,
                dimstyle='MyDimStyle',
                text_rotation=text_rotation,
                dxfattribs={"layer": layer}
            )
            d.render()


if __name__ == '__main__':
    draw_beam_rebar_dxf(output_folder=r'D:\Desktop\BeamQC\TEST\2024-1021',
                        beam_tmp_file=r'TEST\2024-1021\廍子社宅-20241021_142536-beam-object.pkl')
