
# coding: utf-8


import xlwt
import os
import datetime as dt

def set_style(name, height, bold=False, pattern_switch=False, alignment_center=True,underline=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height
    font.italic = False
    if underline:
        font.underline = 0x01

    borders = xlwt.Borders()  # 边框
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    alignment = xlwt.Alignment()  # 居中

    if alignment_center:
        alignment.horz = 0x02
        alignment.vert = 0x01
    else:
        alignment.horz = 0x01
        alignment.vert = 0x01
    # alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行

    pattern = xlwt.Pattern()  # 模式
    pattern.pattern = pattern_switch
    # pattern.pattern = 0x01
    pattern.pattern_fore_colour = 52
    pattern.pattern_back_colour = 52

    style.font = font
    style.borders = borders
    style.alignment = alignment
    style.pattern = pattern
    return style

def create_book(SheetName="Sheet1"):
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet(SheetName)
    return book, sheet1

# def write_ih_book(data, sheet1):
#     pass

def check_folder(folders):
    for folder in folders:
        if not os.path.exists(folder):
            print "创建文件目录", folder
            os.mkdir(folder)

    today_str = dt.datetime.today().strftime("%Y-%m-%d")
    data_folder = "data/"+today_str
    if not os.path.exists(data_folder):
        print "创建本周数据目录", data_folder
        os.mkdir(data_folder)

    return today_str


def write_simple_xls(df,xls_name):
    folders = ["data/"]
    child_folder = check_folder(folders)
    df.to_excel(folders[0] + child_folder + "/" +xls_name + ".xls")
