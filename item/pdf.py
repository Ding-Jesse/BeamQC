import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace
class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.WIDTH = 210
        self.HEIGHT = 297
        
    def header(self):
        # Custom logo and positioning
        # Create an `assets` folder and put any wide and short image inside
        # Name the image `logo.png`
        self.image('assets/logo.png', 10, 8, 33)
        self.set_font('Arial', 'B', 16)
        self.cell(self.WIDTH - 80)
        self.cell(60, 1, 'Test report', 0, 0, 'R')
        self.ln(20)
        
    def footer(self):
        # Page numbers in the footer
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

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
    def add_table(self, TABLE_DATA,table_title,font:str,col_widths:tuple=()):
        self.add_text(table_title)
        blue = (0, 0, 255)
        grey = (228, 240, 239)
        self.set_font(font, size=12)
        headings_style = FontFace(color=blue, fill_color=grey)
        if col_widths:
            col_widths = tuple(item / sum(col_widths) * self.epw for item in col_widths)
        with self.table(headings_style=headings_style,text_align="CENTER",col_widths=col_widths) as table:
        # with self.table(**table_prop) as table:
            for data_row in TABLE_DATA:
                row = table.row()
                for datum in data_row:
                    if isinstance(datum,float):
                        row.cell(str(round(datum,2)))
                    else:
                        row.cell(datum)
        self.ln()
    def add_text(self, texts,align='C'):
        # self.set_y(0)FPDF te
        self.set_font("標楷體", size=12)
        self.cell(w=self.epw,align=align,txt=texts,border=0)
        self.ln()
    def add_prop(self,prop_dict:dict(),font:str):
        self.ln(10)
        self.set_font(font, size=12)
        for key,item in prop_dict.items():
            self.cell(w=self.epw*1/4,align='L',txt=key)
            self.cell(w=self.epw*3/4,align='L',txt=item)
            self.ln(10)
        self.add_dashed_line()

    def add_dashed_line(self):
        self.dashed_line(x1=self.get_x(),
                         x2=self.get_x() + self.epw,
                         y1=self.get_y(),
                         y2=self.get_y())
        self.ln()

def create_scan_pdf(rebar_df:tuple,scan_df:tuple,scan_list:list,project_prop:dict,pdf_filename:str):
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
    pdf = PDF()
    pdf.add_page()
    pdf.add_font('標楷體','',r'D:\Desktop\BeamQC\assets\msjhbd.ttc',True)
    pdf.add_prop(prop_dict=project_prop,font="標楷體")
    pdf.add_table(TABLE_DATA=trans_df_to_table(rebar_df,'Story'),table_title="鋼筋統計表",font="標楷體")
    pdf.add_dashed_line()
    match_index_with_serial(scan_df=scan_df,scan_list=scan_list)
    pdf.add_table(TABLE_DATA=trans_df_to_table(scan_df),table_title="梁檢核表",font="標楷體",col_widths=(1,1,4,4))
    pdf.add_text(texts= "備註:依照",align='L')
    pdf.add_text(texts= "1. “建築技術規則”，內政部，最新版。",align='L')
    pdf.add_text(texts= "2. “混凝土結構設計規範”，內政部，100 年 7 月。",align='L')
    pdf.add_text(texts= "3. “結構混凝土施工規範”，內政部，110 年 9 月。",align='L')
    pdf.ln(10)
    pdf.add_text(f'--------報告結束--------')
    pdf.output(pdf_filename)
def trans_df_to_table(df:pd.DataFrame,reset_name=""):
    table = []
    if reset_name:
        df = df.rename_axis(reset_name).reset_index()
    else:
        df = df.reset_index(drop=True)
    table.append(list(df.columns))
    list_of_row = df.to_numpy().tolist()
    table.extend(list_of_row)
    return table
def match_index_with_serial(scan_list:list,scan_df:pd.DataFrame):
    item_df = scan_df['檢核項目'].drop_duplicates()
    scan_dict = {}
    for item in item_df:
        scan = [sc for sc in scan_list if sc.scan_index == int(item)]
        scan_dict.update({item:scan[0]})
    scan_df['檢核項目'] = scan_df['檢核項目'].apply(lambda x:scan_dict[x].ng_message)
    # row['檢核項目'] = scan_dict[row['檢核項目'] ].ng_message
    pass

if __name__ == '__main__':
    project_prop = {
        '專案名稱:':"測試案例",
        '測試日期:':"YYYY/MM/DD",
        '測試人員:':"XXX",
    }
    TABLE_DATA = (
        ("Story","#3", "#4", "#5", "#6","#7","#8","#10"	,"#11","total"),
        ("3F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
        ("2F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
        ("1F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
    )

    TABLE_DATA2 = (
        ("樓層","編號","檢核項目", "結果"),
        ("3F","B1-1",	"【0204】請確認左端下層筋下限，是否符合規範 3.6 規定","0204:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:10.134"),
        ("B1F",	"B2-3",	"【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",	"0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("B2F",	"B2-3", "【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",	"0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("3F","B1-1",	"【0204】請確認左端下層筋下限，是否符合規範 3.6 規定","0204:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:10.134"),
        ("B1F",	"B2-3",	"【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",	"0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
        ("B2F",	"B2-3", "【0201】請確認左端上層筋下限，是否符合規範 3.6 規定",	"0201:max(code3_3:11.22cm2 ,code3_4:10.5cm2) > 鋼筋總面積:7.74"),
    )
    pdf = PDF()
    pdf.add_page()
    pdf.add_font('標楷體','',r'D:\Desktop\BeamQC\assets\msjhbd.ttc',True)
    pdf.add_prop(prop_dict=project_prop,font="標楷體")
    pdf.add_table(TABLE_DATA=TABLE_DATA,table_title="鋼筋統計表",font="標楷體")
    pdf.add_table(TABLE_DATA=TABLE_DATA2,table_title="梁檢核表",font="標楷體",col_widths=(1,1,4,4))
    pdf.output(r'D:\Desktop\BeamQC\assets\table.pdf')