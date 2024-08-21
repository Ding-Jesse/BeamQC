import numpy as np
import string
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from matplotlib.font_manager import FontProperties
import matplotlib.pyplot as plt
from matplotlib.axes import Axes
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
# plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
# plt.rcParams['axes.unicode_minus'] = False
from fpdf import FPDF
from fpdf.fonts import FontFace
from fpdf.enums import XPos, YPos

plt.rcParams['font.sans-serif'] = ['SimHei']


class PDF(FPDF):
    report_type = 'Rebar report'

    def __init__(self, report_type='Rebar report'):
        super().__init__()
        self.WIDTH = 210
        self.HEIGHT = 297
        self.report_type = report_type

    def header(self):
        # Custom logo and positioning
        # Create an `assets` folder and put any wide and short image inside
        # Name the image `logo.png`
        self.image('assets/logo.png', 10, 8, 33)
        self.set_font('helvetica', 'B', 16)
        self.cell(self.WIDTH - 80)
        self.set_font("標楷體", size=12)
        self.cell(0, 1, self.report_type,
                  new_x=XPos.RMARGIN, new_y=YPos.TOP, align='R')
        self.set_font('helvetica', 'B', 16)
        self.ln(20)

    def footer(self):
        # Page numbers in the footer
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')
        # self.cell(0, 10, 'Page ' + str(self.page_no()), new_x=self.XPos.RIGHT, new_y=self.YPos.TOP)

    def page_body(self, images):
        # Determine how many plots there are per page and set positions
        # and margins accordingly
        if len(images) == 3:
            self.image(images[0], 15, 25, self.WIDTH - 30)
            self.image(images[1], 15, self.WIDTH / 2 + 5, self.WIDTH - 30)
            self.image(images[2], 15, self.WIDTH / 2 + 90, self.WIDTH - 30)
        elif len(images) == 2:
            self.image(images[0], 15, 25, self.WIDTH - 30)
            self.image(images[1], 15, self.WIDTH / 2 + 5, self.WIDTH - 30)
        else:
            self.image(images[0], 15, 25, self.WIDTH - 30)

    def print_page(self, images):
        # Generates the report
        self.add_page()
        self.page_body(images)

    def add_table(self, TABLE_DATA, table_title, font: str, col_widths: list = [], bold_last: bool = False):
        xlen = len(TABLE_DATA[0]) - 1
        ylen = len(TABLE_DATA) - 1
        self.add_text(table_title)
        blue = (0, 0, 255)
        grey = (228, 240, 239)
        yellow = (255, 255, 0)
        self.set_font(font, size=10)
        headings_style = FontFace(color=blue, fill_color=grey)
        cell_style = None
        if len(TABLE_DATA[0]) > len(col_widths) and col_widths:
            col_widths.extend([col_widths[-1]] *
                              (len(TABLE_DATA[0]) - len(col_widths)))
        if col_widths:
            col_widths = tuple(item / sum(col_widths) *
                               self.epw for item in col_widths)
        with self.table(headings_style=headings_style, text_align="CENTER", col_widths=col_widths) as table:
            # with self.table(**table_prop) as table:
            # pdf.set_font(style="B")
            for y, data_row in enumerate(TABLE_DATA):
                row = table.row()
                for x, datum in enumerate(data_row):
                    if x == xlen and y == ylen and bold_last:
                        self.set_font("Times", style="B", size=20)
                        cell_style = FontFace(color=blue, fill_color=yellow)
                    if isinstance(datum, float) or isinstance(datum, int):
                        row.cell(str(round(datum, 2)), style=cell_style)
                        cell_style = None
                    else:
                        row.cell(datum)
                    if x == xlen and y == ylen and bold_last:
                        self.set_font(font, size=10)
        self.ln()

    def add_text(self, texts, align='C'):
        # self.set_y(0)FPDF te
        self.set_font("標楷體", size=12)
        # self.cell(w=self.epw, align=align, txt=texts, border=0)
        # Add title above the image, using multi_cell for auto-wrapping
        if isinstance(texts, list):
            texts = "\n".join(texts)
        self.multi_cell(w=0, txt=texts, align=align,
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln()

    def add_prop(self, prop_dict: dict, font: str):
        self.ln(10)
        self.set_font(font, size=12)
        for key, item in prop_dict.items():
            self.cell(w=self.epw*1/4, align='L', txt=key)
            self.cell(w=self.epw*3/4, align='L', txt=item)
            self.ln(10)
        self.add_dashed_line()

    def add_dashed_line(self):
        self.dashed_line(x1=self.get_x(),
                         x2=self.get_x() + self.epw,
                         y1=self.get_y(),
                         y2=self.get_y())
        self.ln()

    def add_image(self, image_path, title='', page_width='', page_height=''):

        # Get the page dimensions
        if not page_width:
            page_width = self.w - self.l_margin - self.r_margin
        if not page_height:
            page_height = self.h - self.t_margin - self.b_margin - 35
        with Image.open(image_path) as img:
            width, height = img.size

        # Calculate scale factor
        width_ratio = page_width / width
        height_ratio = page_height / height
        scale_factor = min(width_ratio, height_ratio)

        # Calculate new dimensions
        new_width = width * scale_factor
        new_height = height * scale_factor

        # Add title above the first image, centered
        if title != '':
            self.cell(0, 10, title,
                      new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

        # Add the image while maintaining its original scale
        self.image(image_path, x="C", w=new_width, h=new_height)
        self.ln()


def create_scan_pdf(rebar_df: pd.DataFrame,
                    concrete_df: pd.DataFrame,
                    formwork_df: pd.DataFrame,
                    scan_df: pd.DataFrame,
                    ng_sum_df: pd.DataFrame,
                    beam_ng_df: pd.DataFrame,
                    scan_list: list,
                    project_prop: dict,
                    pdf_filename: str,
                    item_name: str,
                    **kwargs):
    '''
    Create scan pdf report \n
    Args:
        project_prop=
        {
            '專案名稱:':"測試案例",\n
            '測試日期:':"YYYY/MM/DD",\n
            '測試人員:':"XXX",\n
        }\n
        rebar_df = (    
            ("Story","#3", "#4", "#5", "#6","#7","#8","#10"	,"#11","total"),\n
            ("3F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0")\n
        )\n
        scan_df = (
            ("樓層","編號","檢核項目", "結果"),\n
            ("3F","B1-1",	"【0204】請確認左端下層筋下限，是否符合規範 3.6 規定","0204:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:10.134"),\n
        )
    '''
    pdf = PDF(item_name)
    pdf.add_font('標楷體', '', r'assets\msjhbd.ttc', True)
    pdf.add_page()
    pdf.add_prop(prop_dict=project_prop, font="標楷體")
    pdf.multi_cell(w=80, h=10, txt="數量統計不包含:\n-工作筋\n-穿孔補強\n-僅供參考")
    pdf.ln()
    pdf.add_table(TABLE_DATA=trans_df_to_table(rebar_df, 'Story'),
                  table_title=f"{item_name}鋼筋統計表(tf)", font="標楷體", bold_last=True)
    pdf.add_page()
    pdf.add_table(TABLE_DATA=trans_df_to_table(concrete_df, 'Story'),
                  table_title=f"{item_name}混凝土統計表(m3)", font="標楷體", bold_last=True)
    pdf.add_page()
    pdf.add_table(TABLE_DATA=trans_df_to_table(formwork_df, 'Story'),
                  table_title=f"{item_name}模板統計表(m2)", font="標楷體", bold_last=True)
    if 'header_list' in kwargs and 'ratio_dict' in kwargs:
        if kwargs['report_type'].casefold() == 'beam':
            item_name = '梁'
            pdf.add_page(orientation="landscape")
            # pdf.add_text(texts="鋼筋比層樓分布", align='C')
            try:
                top_png_file, bot_png_file = survey(
                    results=kwargs['ratio_dict'], category_names=kwargs['header_list'])
                pdf.add_image(top_png_file, title="鋼筋比層樓分布(上層)")
                pdf.add_page(orientation="landscape")
                pdf.add_image(bot_png_file, title="鋼筋比層樓分布(下層)")
                # pdf.image(top_png_file, h=pdf.eph - 35, w=pdf.epw, x='C')
                # pdf.add_page(orientation="landscape")
                # pdf.add_text(texts="鋼筋比層樓分布", align='C')
                # pdf.image(bot_png_file, h=pdf.eph - 35, w=pdf.epw, x='C')
            except:
                pass
        if kwargs['report_type'].casefold() == 'column':
            item_name = '柱'
            pdf.add_page()
            # pdf.add_text(texts="鋼筋比層樓分布", align='C')
            png_file = column_survey(
                results=kwargs['ratio_dict'], category_names=kwargs['header_list'])
            if png_file:
                pdf.add_image(png_file, "鋼筋比層樓分布")
        # pdf.image(top_png_file,h=pdf.eph - 35,keep_aspect_ratio=True)

    if 'header_list' in kwargs and 'ratio_dict' in kwargs:
        # pdf.add_page(orientation="landscape")
        pdf.add_page()
        png_file = plot_rebar_stack_percentage_bar(
            dataset_dict=rebar_df.T.to_dict())
        pdf.add_image(png_file, '號數樓層分布', page_height=(
            pdf.h - pdf.t_margin - pdf.b_margin - 50) / 2)
        # pdf.image(png_file, h=pdf.eph - 35, w=pdf.epw, x='C')
        png_file = plot_rebar_pie_chart(dataset_dict=rebar_df.T.to_dict())
        pdf.add_image(png_file, '號數分布', page_height=(
            pdf.h - pdf.t_margin - pdf.b_margin - 50) / 2)
        # pdf.image(png_file, h=pdf.eph - 35, w=pdf.epw, x='C')
    pdf.add_page()
    pdf.add_table(TABLE_DATA=trans_df_to_table(ng_sum_df, 'Scan Item'),
                  table_title=f"{item_name}檢核表", font="標楷體", col_widths=[4, 1, 1])
    pdf.add_dashed_line()
    pdf.add_page()
    pdf.add_table(TABLE_DATA=trans_df_to_table(
        beam_ng_df), table_title=f"{item_name}檢核表", font="標楷體", col_widths=[1, 1, 5, 1])
    pdf.add_dashed_line()
    pdf.add_page()
    match_index_with_serial(scan_df=scan_df, scan_list=scan_list)
    pdf.add_table(TABLE_DATA=trans_df_to_table(
        scan_df), table_title=f"{item_name}檢核表", font="標楷體", col_widths=[1, 1, 5, 5])
    pdf.ln(10)

    pdf.add_text(texts=["備註:依照",
                        "1. “建築技術規則”，內政部，最新版。",
                        "2. “混凝土結構設計規範”，內政部，100 年 7 月。",
                        "3. “結構混凝土施工規範”，內政部，110 年 9 月。"], align='L')
    pdf.ln(10)
    pdf.add_text('--------報告結束--------')
    pdf.add_page()
    if 'detail_report' in kwargs:
        for details in kwargs['detail_report']:
            pdf.add_text(texts=details, align='L')
    pdf.output(r'assets\contents.pdf')

    add_cover(cover_pdf_path=r'assets\封面.pdf',
              content_pdf_path=r'assets\contents.pdf', output_pdf=pdf_filename)


def trans_df_to_table(df: pd.DataFrame, reset_name=""):
    table = []
    if reset_name:
        df = df.rename_axis(reset_name).reset_index()
    else:
        df = df.reset_index(drop=True)
    table.append(list(df.columns))
    list_of_row = df.to_numpy().tolist()
    table.extend(list_of_row)
    return table


def survey(results: dict[str, dict], category_names: list):
    '''
    Parameters
    ----------
    results : dict
        A mapping from question labels to a list of answers per category.
        It is assumed all lists contain the same number of entries and that
        it matches the length of *category_names*.
    category_names : list of str
        The category labels.
    return img file path
    '''
    custom_text = [chr(i)
                   for i in range(ord('A'), ord('A') + len(category_names))]
    file_path_top = r'assets/top.png'
    file_path_bot = r'assets/bot.png'
    title = {
        (0, 0): '左端上層',
        (0, 1): '中央上層',
        (0, 2): '右端上層',
        (1, 0): '左端下層',
        (1, 1): '中央下層',
        (1, 2): '右端下層',
    }
    labels = list(results.keys())
    fig, ax0 = plt.subplots(1, 3, figsize=(29.7, 21))
    fig2, ax1 = plt.subplots(1, 3, figsize=(29.7, 21))
    category_colors = cm.get_cmap('jet')(
        np.linspace(0, 1, len(category_names)))
    for dir, (x, y) in enumerate([(0, 0), (0, 1), (0, 2), (1, 0), (1, 1), (1, 2)]):
        ar = np.zeros((len(labels), len(category_names)))
        for i, values in enumerate(results.values()):
            temp_ar = np.array(list(values.values()))
            ar[i] = temp_ar[:, dir]
        data = ar
        data_sum = data.sum(axis=1)
        data_sum[data_sum == 0] = 1
        data = (data.T/data_sum).T * 100
        data_cum = data.cumsum(axis=1)
        # category_colors = cm.get_cmap('RdYlGn_r')(
        #     np.linspace(0, 1, data.shape[1]))
        if x == 1:
            ax = ax1
        else:
            ax = ax0
        ax[y].invert_yaxis()
        ax[y].tick_params(axis='y', labelsize=30)
        ax[y].xaxis.set_visible(True)
        ax[y].set_xlim(0, np.sum(data, axis=1).max())
        ax[y].set_xlabel('percentage(%)', fontsize='xx-large')
        ax[y].set_title(f'{title[(x,y)]}', fontsize=30)

        for i, (colname, color) in enumerate(zip(category_names, category_colors)):
            widths = data[:, i]
            starts = data_cum[:, i] - widths
            # colname = colname.replace("鋼筋比", "ratio")
            ax[y].barh(labels, widths, left=starts, height=0.5,
                       label=colname, color=color)

            # r, g, b, _ = color
            # text_color = 'white' if r * g * b < 0.5 else 'darkgrey'
            # ax.bar_label(rects, label_type='center', color=text_color)
        custom_plot(ax=ax[y],
                    custom_text=custom_text,
                    labels=labels)
    fig.tight_layout()
    fig2.tight_layout()
    custom_legend(ax=ax0[1],
                  custom_text=custom_text,
                  ncols=len(category_names),
                  bbox_to_anchor=(0.5, 1.05),
                  loc='lower center',
                  fontsize='xx-large'
                  )
    custom_legend(ax=ax1[1],
                  custom_text=custom_text,
                  ncols=len(category_names),
                  bbox_to_anchor=(0.5, 1.05),
                  loc='lower center',
                  fontsize='xx-large'
                  )
    fig.savefig(file_path_top, bbox_inches='tight')
    fig2.savefig(file_path_bot, bbox_inches='tight')
    # plt.savefig(file_path, bbox_inches='tight')
    return file_path_top, file_path_bot


def plot_rebar_stack_percentage_bar(dataset_dict: dict[str, dict[str, float]]):
    image_path = r'assets/rebar_stack_percentage_bar.png'

    # Data
    # dataset_dict = {
    #     'PF': {'#3': 0.0, '#4': 1.0, '#5': 12.0},
    #     'RF': {'#3': 1.0, '#4': 2.0, '#5': 12.0},
    #     '2F': {'#3': 2.0, '#4': 3.0, '#5': 22.0},
    #     '1F': {'#3': 3.0, '#4': 2.0, '#5': 24.0},
    # }
    # Transfer Data
    categories = dataset_dict.keys()
    dataset = list(dataset_dict.values())
    rebar_dataset = {}
    for data in dataset:
        for key, item in data.items():
            if key not in ['#6', '#7', '#8', '#10', '#11']:
                continue
            if key not in rebar_dataset:
                rebar_dataset[key] = []
            rebar_dataset[key].append(item)

    # Calculate percentages for each category
    total_values = [sum(x) if sum(
        x) > 0 else 1 for x in zip(*rebar_dataset.values())]
    percent_datasets = {key: [v / total * 100 for v, total in zip(
        data, total_values)] for key, data in rebar_dataset.items()}

    # Create an array of x values for each category
    x = np.arange(len(categories))

    # Create a figure
    fig, ax = plt.subplots(figsize=(29.7, 21))
    category_colors = cm.get_cmap('jet')(
        np.linspace(0, 1, len(percent_datasets)))
    # Create stacked bar charts for each dataset
    for i, (key, data) in enumerate(percent_datasets.items()):
        ax.bar(x, data, label=key, bottom=np.sum(list(percent_datasets.values())[
            :list(percent_datasets.keys()).index(key)], axis=0), color=category_colors[i])

    # Set the x-axis labels
    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.tick_params(axis='y', labelsize=30)
    ax.tick_params(axis='x', labelsize=30)
    ax.set_xlabel('樓層', fontsize=30)

    # Set the y-axis label
    ax.set_ylabel('百分比', fontsize=30)

    # Add a legend
    ax.legend(fontsize=30)

    # Add a title
    # ax.set_title('號數樓層分布', fontsize=30)

    fig.tight_layout()
    fig.savefig(image_path, bbox_inches='tight')

    return image_path


def plot_rebar_pie_chart(dataset_dict: dict[str, dict[str, float]]):
    image_path = r'assets/rebar_pie_chart.png'

    # Data
    # dataset_dict = {
    #     'PF': {'#3': 0.0, '#4': 1.0, '#5': 12.0},
    #     'RF': {'#3': 1.0, '#4': 2.0, '#5': 12.0},
    #     '2F': {'#3': 2.0, '#4': 3.0, '#5': 22.0},
    #     '1F': {'#3': 3.0, '#4': 2.0, '#5': 24.0},
    # }
    # Transfer Data
    # categories = dataset_dict.keys()
    dataset = list(dataset_dict.values())
    rebar_dataset = {}
    for data in dataset:
        for key, item in data.items():
            if key == 'total':
                continue
            if key not in rebar_dataset:
                rebar_dataset[key] = []
            rebar_dataset[key].append(item)

    # Calculate the sum of values across categories for each dataset
    sum_values = {key: sum(data)
                  for key, data in rebar_dataset.items() if sum(data) > 0}
    # Create fig and axes
    fig, ax = plt.subplots(1, 1, figsize=(29.7, 21))

    ax.pie(sum_values.values(), labels=sum_values.keys(),
           autopct='%1.2f%%', startangle=90, textprops={'fontsize': 40})
    # ax.title('Sum of Values Across Categories')
    ax.axis('equal')
    # ax.set_title('號數分布', pad=50, fontsize=30)
    ax.legend(fontsize=30)

    fig.tight_layout()
    fig.savefig(image_path, bbox_inches='tight')
    return image_path


def column_survey(results: dict[str, dict], category_names: list):
    '''
    Parameters
    ----------
    results : dict
        A mapping from question labels to a list of answers per category.
        It is assumed all lists contain the same number of entries and that
        it matches the length of *category_names*.
    category_names : list of str
        The category labels.
    return img file path
    '''

    custom_text = [chr(i)
                   for i in range(ord('A'), ord('A') + len(category_names))]
    labels = list(results.keys())
    file_path = r'assets/column.png'
    if not results:
        return
    # fig = plt.figure(figsize=(29.7, 21))
    category_colors = cm.get_cmap('jet')(
        np.linspace(0, 1, len(category_names)))
    ar = np.zeros((len(labels), len(category_names)))
    for i, values in enumerate(results.values()):
        temp_ar = np.array(list(values.values()))
        ar[i] = temp_ar[:, 0]
        data = ar
        data_sum = data.sum(axis=1)
        data_sum[data_sum == 0] = 1
        data = (data.T/data_sum).T * 100
        data_cum = data.cumsum(axis=1)
    fig, ax = plt.subplots(1, 1, figsize=(21, 29.7))
    ax.invert_yaxis()
    ax.xaxis.set_visible(True)
    ax.tick_params(axis='y', labelsize=30)
    ax.set_xlim(0, 100)
    ax.set_xlabel('percentage(%)', fontsize=30)
    for i, (colname, color) in enumerate(zip(category_names, category_colors)):
        widths = data[:, i]
        starts = data_cum[:, i] - widths
        # colname = colname.replace("鋼筋比", "ratio")
        ax.barh(labels, widths, left=starts, height=0.5,
                label=colname, color=color)
    fig.tight_layout()
    custom_plot(ax=ax,
                custom_text=custom_text,
                labels=labels)
    custom_legend(ax=ax,
                  custom_text=custom_text,
                  ncols=len(category_names)//3,
                  bbox_to_anchor=(0.5, 1.02),
                  loc='lower center',
                  fontsize=30)
    # ax.legend(ncols=len(category_names)//3,bbox_to_anchor=(0.5, 1.05),
    #         loc='lower center', fontsize='xx-large')

    fig.savefig(file_path, bbox_inches='tight')
    return file_path


def custom_plot(ax: Axes, custom_text: list, labels: list):
    for i, bar in enumerate(ax.patches):
        if bar.get_width() > 0:
            face_color = bar.get_facecolor()
            r, g, b, _ = face_color
            text_color = 'white' if sum(face_color) / 3 < 0.75 else 'black'
            x = bar.get_width()/2+bar.get_x()
            y = bar.get_height()/2+bar.get_y()
            label = custom_text[i // len(labels)]
            ax.text(x, y, label, ha='center', va='center',
                    color=text_color, fontsize=20)
    # custom_labels = [
    #     f'{label} ({text})' for label, text in zip(ax.get_legend_handles_labels()[1], custom_text)]
    # ax.legend(handles=ax.containers, labels=custom_labels,**kwargs)


def custom_legend(ax: Axes, custom_text: list, **kwargs):
    custom_labels = [
        f'{label} ({text})' for label, text in zip(ax.get_legend_handles_labels()[1], custom_text)]
    ax.legend(handles=ax.containers, labels=custom_labels, **kwargs)


def match_index_with_serial(scan_list: list, scan_df: pd.DataFrame):
    item_df = scan_df['檢核項目'].drop_duplicates()
    scan_dict = {}
    for item in item_df:
        scan = [sc for sc in scan_list if sc.scan_index == int(item)]
        scan_dict.update({item: scan[0]})
    scan_df['檢核項目'] = scan_df['檢核項目'].apply(lambda x: scan_dict[x].ng_message)
    # row['檢核項目'] = scan_dict[row['檢核項目'] ].ng_message
    pass


def add_cover(cover_pdf_path, content_pdf_path, output_pdf):
    import fitz  # PyMuPDF
    # Load the cover PDF and the content PDF
    cover_pdf = fitz.open(cover_pdf_path)
    content_pdf = fitz.open(content_pdf_path)
    # Create a new PDF to combine them
    combined_pdf = fitz.open()

    # Add the cover page
    combined_pdf.insert_pdf(cover_pdf)

    # Add the content pages
    combined_pdf.insert_pdf(content_pdf)

    # Save the combined PDF
    combined_pdf.save(output_pdf)


if __name__ == '__main__':
    from numpy import arange, array
    from itertools import cycle
    import random
# from itertools import cycle
    cycol = cycle('bgrcmk')
    project_prop = {
        '專案名稱:': "測試案例",
        '測試日期:': "YYYY/MM/DD",
        '測試人員:': "XXX",
    }
    TABLE_DATA = (
        ("Story", "#3", "#4", "#5", "#6", "#7", "#8", "#10", "#11", "total"),
        ("3F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42", "0", "0"),
        ("2F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42", "0", "0"),
        ("1F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42", "0", "0"),
    )

    TABLE_DATA2 = (
        ("樓層", "編號", "檢核項目", "結果"),
        ("3F", "B1-1",	"【0204】請確認左端下層筋下限，是否符合規範 3.6 規定",
         "0204:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:10.134"),
        ("B1F",	"B2-3",	"【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",
         "0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("B2F",	"B2-3", "【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",
         "0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("3F", "B1-1",	"【0204】請確認左端下層筋下限，是否符合規範 3.6 規定",
         "0204:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:10.134"),
        ("B1F",	"B2-3",	"【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",
         "0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("B2F",	"B2-3", "【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",
         "0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
    )

    ratio_upper_bound_group = list(arange(0.005, 0.03, 0.005))
    ratio_lower_bound_group = list(arange(0, 0.025, 0.005))
    header_list = list(map(
        lambda r, p: f'{p*100}% ≤ 鋼筋比 < {r*100}%', ratio_upper_bound_group, ratio_lower_bound_group))
    header_list.append(f'≥ {ratio_upper_bound_group[-1]*100}%')
    temp = dict()
    for header in header_list:
        temp.update({header: [random.randint(0, 10), random.randint(0, 10), random.randint(0, 10),
                              random.randint(0, 10), random.randint(0, 10), random.randint(0, 10)]})
    temp2 = dict()
    for header in header_list:
        temp2.update({header: [random.randint(0, 10), random.randint(0, 10), random.randint(0, 10),
                               random.randint(0, 10), random.randint(0, 10), random.randint(0, 10)]})
    TABLE_DATA3 = {
        '15F': temp,
        '14F': temp2,
        '13F': temp,
        '12F': temp2,
    }
    # TABLE_DATA4 = {
    #     'right':[random.uniform(0, 0.025) for _ in range(100)],
    #     'middle':[random.uniform(0, 0.015) for _ in range(100)],
    #     'left':[random.uniform(0, 0.025) for _ in range(100)]
    # }

    # import matplotlib.pyplot as plt
    # fig, axs = plt.subplots(1, 1, sharey=True, tight_layout=True)
    column_survey(TABLE_DATA3, header_list)
    # labels = list(TABLE_DATA3.keys())
    # data = array(list(TABLE_DATA3.values()))
    # data_cum = data.cumsum(axis=1)
    # axs.bar(TABLE_DATA3)
    # top,bot = survey(TABLE_DATA3,header_list)
    # plt.savefig('foo2.png', bbox_inches='tight')
    # plt.savefig('foo2.png')
    # plt.show()
    # pdf = PDF()
    # pdf.add_page()
    # pdf.add_font('標楷體','',r'D:\Desktop\BeamQC\assets\msjhbd.ttc',True)
    # pdf.add_prop(prop_dict=project_prop,font="標楷體")
    # pdf.add_table(TABLE_DATA=TABLE_DATA,table_title="鋼筋統計表",font="標楷體")
    # # pdf.cur_orientation = 'L'
    # pdf.add_page(orientation="landscape")
    # pdf.image(top,h=pdf.eph - 25,w=pdf.epw,x='C')
    # pdf.add_page(orientation="landscape")
    # pdf.image(bot,h=pdf.eph - 25,keep_aspect_ratio=True)
    # pdf.add_page(orientation="landscape")
    # pdf.add_table(TABLE_DATA=TABLE_DATA2,table_title="梁檢核表",font="標楷體",col_widths=[1,1,4])

    # pdf.output(r'assets\table.pdf')
