import docx
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT

# Function to add an OMML equation to a paragraph


def add_omml_equation(paragraph, parts):
    run = paragraph.add_run()
    omml = OxmlElement('m:oMathPara')
    math = OxmlElement('m:oMath')
    for part in parts:
        if isinstance(part, str):
            part = create_math_element(part)
        math.append(part)
    # math.append(equation)
    omml.append(math)
    run._r.append(omml)
# Function to create a superscript element


def create_superscript(base, superscript):
    sSup = OxmlElement('m:sSup')
    e = OxmlElement('m:e')
    e_r = OxmlElement('m:r')
    e_t = OxmlElement('m:t')
    e_t.text = base
    e_r.append(e_t)
    e.append(e_r)

    sup = OxmlElement('m:sup')
    sup_r = OxmlElement('m:r')
    sup_t = OxmlElement('m:t')
    sup_t.text = superscript
    sup_r.append(sup_t)
    sup.append(sup_r)

    sSup.append(e)
    sSup.append(sup)
    return sSup

# Function to create a subscript element


def create_subscript(base, subscript):
    sSub = OxmlElement('m:sSub')
    e = OxmlElement('m:e')
    e_r = OxmlElement('m:r')
    e_t = OxmlElement('m:t')
    e_t.text = base
    e_r.append(e_t)
    e.append(e_r)

    sub = OxmlElement('m:sub')
    sub_r = OxmlElement('m:r')
    sub_t = OxmlElement('m:t')
    sub_t.text = subscript
    sub_r.append(sub_t)
    sub.append(sub_r)

    sSub.append(e)
    sSub.append(sub)
    return sSub


def create_superscript_subscript(base, superscript, subscript):
    sSubSup = OxmlElement('m:sSubSup')

    # Base element
    e = OxmlElement('m:e')
    e_r = OxmlElement('m:r')
    e_t = OxmlElement('m:t')
    e_t.text = base
    e_r.append(e_t)
    e.append(e_r)

    # Superscript element
    sup = OxmlElement('m:sup')
    sup_r = OxmlElement('m:r')
    sup_t = OxmlElement('m:t')
    sup_t.text = superscript
    sup_r.append(sup_t)
    sup.append(sup_r)

    # Subscript element
    sub = OxmlElement('m:sub')
    sub_r = OxmlElement('m:r')
    sub_t = OxmlElement('m:t')
    sub_t.text = subscript
    sub_r.append(sub_t)
    sub.append(sub_r)

    sSubSup.append(e)
    sSubSup.append(sub)
    sSubSup.append(sup)
    return sSubSup


def create_fraction(numerators, denominators):
    f = OxmlElement('m:f')

    # Numerator element
    num = OxmlElement('m:num')
    # num_r = OxmlElement('m:r')
    # num_t = OxmlElement('m:t')
    # num_t.text = numerator
    # num_r.append(num_t)
    for numerator in numerators:
        num.append(numerator)

    # Denominator element
    den = OxmlElement('m:den')
    # den_r = OxmlElement('m:r')
    # den_t = OxmlElement('m:t')
    # den_t.text = denominator
    # den_r.append(den_t)
    for denominator in denominators:
        den.append(denominator)

    f.append(num)
    f.append(den)
    return f


def create_math_element(text):
    r = OxmlElement('m:r')
    t = OxmlElement('m:t')
    t.text = text
    r.append(t)
    return r


# Create a new Document
doc = docx.Document()

# Add a title
doc.add_heading('耐震接頭檢討', level=1)

# Add a section for "接頭剪力設計原則"
doc.add_heading('接頭剪力設計原則', level=2)
doc.add_paragraph('梁撓曲鋼筋之應力假設為1.25fy')

# Add mathematical formulas
doc.add_heading('計算Ts1及Mpr-(忽略壓力筋效應)', level=3)
p = doc.add_paragraph()
add_omml_equation(p, [create_subscript('T', 's1'),
                      create_math_element('='),
                      create_subscript('A', 's1'),
                      create_math_element('×1.25fy')])
p.add_run('，由斷面上之力平衡知')
add_omml_equation(p, [create_subscript('C', 'c1'),
                      create_math_element('='),
                      create_subscript('T', 's1')])
p = doc.add_paragraph()
mpr_minus_equation = [create_superscript_subscript('M', '-', 'pr'),
                      create_math_element('='),
                      create_subscript('T', 's1'),
                      create_math_element('×(d-'),
                      create_fraction([create_math_element('1')],
                                      [create_math_element('2')]),
                      create_fraction([create_subscript('T', 's1')],
                                      [create_math_element('0.85'),
                                       create_superscript_subscript(
                                          'f', '\'', 'c'),
                                       create_math_element('b')]),
                      create_math_element(')')]
add_omml_equation(p, mpr_minus_equation)
# doc.save('耐震接頭檢討2.docx')

doc.add_heading('計算Ts2及Mpr+(忽略壓力筋效應)', level=3)
p = doc.add_paragraph()
add_omml_equation(p, [create_subscript('T', 's2'),
                      create_math_element('='),
                      create_subscript('A', 's2'),
                      create_math_element('×1.25fy')])
p.add_run('，由斷面上之力平衡知')
add_omml_equation(p, [create_subscript('C', 'c2'),
                      create_math_element('='),
                      create_subscript('T', 's2')])
p = doc.add_paragraph()
mpr_plus_equation = [create_superscript_subscript('M', '+', 'pr'),
                     create_math_element('='),
                     create_subscript('T', 's2'),
                     create_math_element('×(d-'),
                     create_fraction([create_math_element('1')],
                                     [create_math_element('2')]),
                     create_fraction([create_subscript('T', 's2')],
                                     [create_math_element('0.85'),
                                      create_superscript_subscript(
                                         'f', '\'', 'c'),
                                      create_math_element('b')]),
                     create_math_element(')')]
add_omml_equation(p, mpr_plus_equation)
doc.add_paragraph('計算柱端剪力')
p = doc.add_paragraph()
add_omml_equation(p, [create_subscript('V', 'h'),
                      create_math_element('='),
                      create_fraction([
                          create_superscript_subscript('M', '+', 'pr'),
                          create_math_element('+'),
                          create_superscript_subscript('M', '-', 'pr')
                      ], [
                          create_fraction([
                              create_subscript('H', '1'),
                              create_math_element('+'),
                              create_subscript('H', '2')
                          ], [
                              create_math_element('2')
                          ])
                      ])]
                  )

doc.add_paragraph('計算設計剪力')
p = doc.add_paragraph()
add_omml_equation(p, [create_subscript('V', 'u'),
                      create_math_element('='),
                      create_subscript('T', 's1'),
                      create_math_element('+'),
                      create_subscript('C', 'c2'),
                      create_math_element('-'),
                      create_subscript('V', 'h')])
doc.add_heading('接頭剪力設計例', level=2)
doc.add_paragraph('以新規範401-110檢討梁柱接頭剪力強度，其適用樣態如下表所示：')

# Create a table
table = doc.add_table(rows=1, cols=2, style='Table Grid')
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Set table background color
shading_elm = OxmlElement('w:shd')
shading_elm.set(qn('w:fill'), 'D9EAD3')
for cell in table.rows[0].cells:
    cell._element.get_or_add_tcPr().append(shading_elm)

# Header row formatting
hdr_cells = table.rows[0].cells
hdr_cells[0].text = '項目'
hdr_cells[1].text = '數值'
for cell in hdr_cells:
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = paragraph.runs[0]
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)  # Black color

# Add rows to the table with formulas
rows = [
    (['平行剪力方向X向柱深', [create_subscript('h', 'c')]], '120 cm'),
    (['梁邊與柱邊距離', [create_subscript('x', '1')]], '0 cm'),
    (['梁邊與柱邊距離', [create_subscript('x', '2')]], '35 cm，取120/4 = 30 cm'),
    (['接頭有效抗剪寬度', [create_subscript('b', 'j'),
                   create_math_element('='),
                   create_subscript('b', 'b'),
                   create_math_element('+'),
                   create_subscript('x', '1'),
                   create_math_element('+'),
                   create_subscript('x', '2')]], '95 cm'),
    (['有效抗剪水平斷面積', [create_subscript('A', 'j'),
                    create_math_element('='),
                    create_subscript('b', 'j'),
                    create_math_element('x'),
                    create_subscript('h', 'c')]], '11400 cm²'),
    ([[create_subscript('T', 's1'), "=",
       create_subscript('C', 'c1'), "=",
       create_subscript('A', 's1'), "x1.25fy"]], '513.0 tf'),
    ([mpr_minus_equation], '311.9 tf-m'),
    ([[create_subscript('T', 's2'), "=",
       create_subscript('C', 'c2'), "=",
       create_subscript('A', 's2'), "x1.25fy"]], '345.8 tf'),
    ([mpr_plus_equation], '260.3 tf-m'),
    ([[create_fraction([
        create_subscript('H', '1'),
        create_math_element('+'),
        create_subscript('H', '2')
    ], [
        create_math_element('2')
    ])
    ]], '3.4 m'),
    ('Vh=Mpr++Mpr-H1+H22', '168.3 tf'),
    ('設計剪力Vu=Ts1+Cc2-Vh', '690.5 tf'),
    ('剪力強度∅Vnx=0.85×3.9λfc\'Aj', '0.85×3.9490×11400=836.5 tf'),
    ('∅Vnx≥Vu，DCR=0.83→OK，檢核通過', '')
]

for items, values in rows:
    row_cells = table.add_row().cells
    p_item = row_cells[0].paragraphs[0]
    p_value = row_cells[1].paragraphs[0]

    # run_item = p_item.add_run(item)
    p_item.alignment = WD_ALIGN_PARAGRAPH.CENTER
    row_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for item in items:
        if isinstance(item, str):
            p_item.add_run(item)
        elif isinstance(item, list):
            add_omml_equation(p_item, item)

    for value in values:
        if isinstance(value, str):
            p_value.add_run(value)
        elif isinstance(value, list):
            add_omml_equation(p_value, value)
    # if '×' in value or '²' in value or 'λ' in value or '≥' in value:
    #     add_omml_equation(p_value, [create_math_element(value)])
    # else:
    #     run_value = p_value.add_run(value)

    p_value.alignment = WD_ALIGN_PARAGRAPH.CENTER
    row_cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # run_item.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    # if '×' in value or '²' in value or 'λ' in value or '≥' in value:
    #     for run in p_value.runs:
    #         run.font.color.rgb = RGBColor(0, 0, 0)  # Black color

# Apply border to the table


def set_cell_border(cell, **kwargs):
    """
    Set cell border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        bottom={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        left={"sz": 12, "val": "single", "color": "000000", "space": "0"},
        right={"sz": 12, "val": "single", "color": "000000", "space": "0"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f"w:{edge}"
            element = OxmlElement(tag)
            for key in ["sz", "val", "color", "space"]:
                if key in edge_data:
                    element.set(qn(f"w:{key}"), str(edge_data[key]))
            tcPr.append(element)


for row in table.rows:
    for cell in row.cells:
        set_cell_border(cell, top={"sz": 12, "val": "single", "color": "000000"},
                        bottom={"sz": 12, "val": "single", "color": "000000"},
                        left={"sz": 12, "val": "single", "color": "000000"},
                        right={"sz": 12, "val": "single", "color": "000000"})

# Save the document
doc.save('耐震接頭檢討2.docx')
