# coding=utf-8
from abc import abstractmethod
import xlrd
import xlwt
from xlwt import *
import sys
__metaclass__ = type

# 主表title
m_title = {'序号': 0, '群昵称': 1, '订单时间': 2, '订单编号': 3,
           '商品名称': 4, '订单状态': 5, '金额': 6, '返利金额': 7,
           '结算状态': 8, '备注': 9, '比例': 10, '佣金': 11, '绑定用户': 12}

# 从系统导出的新表title
s_title = {'订单时间': 0, '订单编号': 1, '商品名称': 2, '订单状态': 3,
           '金额': 4, '比例': 5, '佣金': 6, '绑定用户': 7}


class ExcelStyleProvider:
    # black, white, light_red, light_green, blue, yellow, magenta = (0, 1, 29, 17, 4, 5, 6)
    black, white, red, light_green, blue, yellow, magenta = (0, 1, 2, 17, 4, 5, 6)

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

class ExcelReader:
    def __init__(self):
        self.workbook = self.worksheet = self.sheet_rows = None

    def open(self, file_path: str, sheet_num: int):
        self.workbook = xlrd.open_workbook(file_path)  # 打开xls文件
        self.worksheet = self.workbook.sheets()[sheet_num]  # 打开第X张表
        self.sheet_rows = self.worksheet.nrows  # 获取表的行数

    # 读入一行
    def read_row(self, row_num: int) -> [str]:
        return self.worksheet.row_values(row_num)


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


class ExcelMergeRule:
    def __init__(self, reader1: ExcelReader, reader2: ExcelReader, writer1: ExcelWriter):
        self.main_reader = reader1
        self.sec_reader = reader2
        self.writer = writer1

    @abstractmethod
    def run(self):
        pass


# 规则：系统导出的新表中，如果某个订单在主表中已经存在，并且新表的订单状态
# 和主表相比已经发生变化，就要把这条订单更新到主表中。如果新表中的订单编号
# 在主表中不存在，就追加到主表的后面。订单编号+商品名称唯一确定一条记录
#　　　　主表　　　　　　　　　　系统导出的新表　　　　　　　整合后的表　　　　　　　　　　　　　　
# ∣—————————∣　　∣—————————∣　　∣—————————∣
# ∣订单编号∣状　态　∣　　∣订单编号∣状　态　∣　　∣订单编号∣状　态　∣
# ∣————————--∣　　∣————————--∣　　∣————————--∣
# ∣0 0 0 1 ∣　　　　∣　　∣0 0 2 3 ∣　　　　∣　　∣0 0 0 1 ∣　　　　∣
# ∣0 0 0 2 ∣　　　　∣　　∣0 0 1 8 ∣　　　　∣　　∣0 0 0 2 ∣　　　　∣
# ∣0 0 0 5 ∣付　款　∣　　∣0 0 0 5 ∣失　效　∣　　∣0 0 0 5 ∣失　效　∣
# ∣0 0 1 5 ∣　　　　∣+　 ∣0 1 1 4 ∣　　　　∣=　 ∣0 0 1 5 ∣　　　　∣
# ∣　　　　∣　　　　∣　　∣　　　　∣　　　　∣　　∣0 0 2 3 ∣　　　　∣
# ∣　　　　∣　　　　∣　　∣　　　　∣　　　　∣　　∣0 0 1 8 ∣　　　　∣
# ∣　　　　∣　　　　∣　　∣　　　　∣　　　　∣　　∣0 1 1 4 ∣　　　　∣
# ∣________⊥________∣　　∣________⊥________∣　　∣________⊥________∣


class ExcelMergeOrderState(ExcelMergeRule):
    @staticmethod
    def __diff(list_a, list_b):
        return list(set(list_a) - set(list_b))

    @staticmethod
    def __create_row_style(row_data: [str]) -> XFStyle:
        style_provider = ExcelStyleProvider()
        bg = ExcelStyleProvider.white
        fg = ExcelStyleProvider.black
        # if row_data[m_title['绑定用户']] == "【锁定】":
        #    bg = ExcelStyleProvider.light_red
        if row_data[m_title['订单状态']] == "订单失效":
            fg = ExcelStyleProvider.blue
        if row_data[m_title['订单状态']] == "订单结算":
            fg = ExcelStyleProvider.red

        style = style_provider.create_style(bg, fg)
        return style

    def run(self):
        found = []
        for i in range(self.main_reader.sheet_rows):  # 循环逐行
            if i == 0:  # 跳过第一行,输出title
                writer.write_sequence_row(self.main_reader.read_row(0))
                continue
            main_row = self.main_reader.read_row(i)
            for j in range(self.sec_reader.sheet_rows):
                if j == 0:
                    continue
                sec_row = self.sec_reader.read_row(j)
                if main_row[m_title['订单编号']] == sec_row[s_title['订单编号']] and \
                   main_row[m_title['商品名称']] == sec_row[s_title['商品名称']]:
                    main_row[m_title['订单状态']] = sec_row[s_title['订单状态']]
                    found.append(j)
                    break
            # 更新完一行，写入结果表中
            style = self.__create_row_style(main_row)
            self.writer.write_sequence_row(main_row, style)

        # 系统导出的新表中没有写入的行追加到结果表
        not_found = self.__diff(range(self.sec_reader.sheet_rows), found)
        for i in not_found:
            if i == 0:
                continue
            row = self.sec_reader.read_row(i)
            row.insert(m_title['序号'], "")
            row.insert(m_title['群昵称'], "")
            row.insert(m_title['返利金额'], "")
            row.insert(m_title['结算状态'], "")
            row.insert(m_title['备注'], "")
            style = self.__create_row_style(row)
            self.writer.write_sequence_row(row, style)
        self.writer.save()

if len(sys.argv) < 3:
    print("Usage: %s main_excel second_excel merged_excel" % (sys.argv[0]))
else:
    print("Start process ...")
    main_reader = ExcelReader()
    main_reader.open(sys.argv[1], 0)
    sec_reader = ExcelReader()
    sec_reader.open(sys.argv[2], 0)
    writer = ExcelWriter()
    writer.open(sys.argv[3])
    rule1 = ExcelMergeOrderState(main_reader, sec_reader, writer)
    rule1.run()
    print("Done")
