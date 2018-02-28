import xlrd
import xlwt
from xlwt import *

class ExcelReader:
    def __init__(self):
        self.workbook = self.worksheet = self.sheet_rows = None
        self.start_row = 0

    def open(self, file_path: str, sheet_num: int):
        self.workbook = xlrd.open_workbook(file_path)  # 打开xls文件
        self.worksheet = self.workbook.sheets()[sheet_num]  # 打开第X张表
        self.sheet_rows = self.worksheet.nrows  # 获取表的行数

    def seek(self, row_num: int):
        self.start_row = row_num

    # 读入一行
    def read_row(self, row_num: int) -> [str]:
        return self.worksheet.row_values(self.start_row + row_num)

    def read_col(self, col_num: int) -> [str]:
        s = []
        for i in range(self.sheet_rows - self.start_row):  # 循环逐行
            val = self.read_row(i)[col_num]
            if val != "":
                s.append(val)
        return s

class ExcelStyleProvider:
    black, white, light_red, light_green, blue, yellow, magenta = (0, 1, 29, 17, 4, 5, 6)

    def __init__(self):
        # 实线边框
        self.borders = Borders()
        self.borders.left = 1
        self.borders.right = 1
        self.borders.top = 1
        self.borders.bottom = 1

    def create_style(self, bg_color: int, fg_color: int) -> XFStyle:
        style = XFStyle()

        pattern = Pattern()  # 创建一个模式
        pattern.pattern = Pattern.SOLID_PATTERN  # 设置其模式为实型
        # 设置单元格背景颜色
        pattern.pattern_fore_colour = bg_color
        style.pattern = pattern

        fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
        fnt.colour_index = fg_color  # 设置其字体颜色
        style.font = fnt  # 将赋值好的模式参数导入Style

        style.borders = self.borders

        return style

    def default_style(self) -> XFStyle:
        return self.create_style(self.white, self.black)

class ExcelWriter:
    def __init__(self):
        self.workbook = self.worksheet = self.cur_row = None
        self.file_path = None

    def open(self, file_path: str):
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet('1', cell_overwrite_ok=True)
        self.cur_row = 0
        self.file_path = file_path


    # 依次写入新行
    def write_sequence_row(self, row_data: [str], style = ExcelStyleProvider().default_style()):
        n = len(row_data)
        for i in range(n):
            self.worksheet.write(self.cur_row, i, row_data[i], style)
        self.cur_row += 1

    def save(self):
        self.workbook.save(self.file_path)
