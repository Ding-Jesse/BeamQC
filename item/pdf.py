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
        self.set_font('Arial', 'B', 11)
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
    def add_table(self, TABLE_DATA,table_title):
        self.add_text(table_title)
        blue = (0, 0, 255)
        grey = (228, 240, 239)
        self.set_font("Times", size=12)
        headings_style = FontFace(color=blue, fill_color=grey)
        with self.table(headings_style=headings_style,text_align="CENTER") as table:
            for data_row in TABLE_DATA:
                row = table.row()
                for datum in data_row:
                    row.cell(datum)
    def add_text(self, texts):
        # self.set_y(0)
        self.cell(w=self.epw,align='C',txt=texts,border=0)
        self.ln()
    def add_prop(self,prop_dict:dict()):
        self.ln(10)
        self.set_font("標楷體", size=12)
        for key,item in prop_dict.items():
            self.cell(w=self.epw*1/3,align='C',txt=key)
            self.cell(w=self.epw*2/3,align='L',txt=item)
            self.ln(10)
        # self.cell(txt="-----------------------------")
TABLE_DATA = (
    ("Story","#3", "#4", "#5", "#6","#7","#8","#10"	,"#11","total"),
    ("3F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
    ("2F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
    ("1F",	"0",	"6.12",	"0",	"0",	"1.52",	"8.8",	"10.42","0","0"),
)
pdf = PDF()
pdf.add_page()
pdf.add_font('標楷體','',r'D:\Desktop\BeamQC\assets\msjhbd.ttc',True)
pdf.add_prop({
    '專案名稱':"測試案例",
    '測試日期':"YYYY/MM/DD",
    '測試人員':"XXX",
})
pdf.add_table(TABLE_DATA=TABLE_DATA,table_title="鋼筋統計表")
# pdf.set_font("Times", size=16)
pdf.output('assets/table.pdf')