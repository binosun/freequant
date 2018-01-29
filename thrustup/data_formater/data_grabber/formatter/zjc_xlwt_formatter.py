
# coding: utf-8


import copy
import pandas as pd

from  freequant.thrustup.data_formater.data_grabber.formatter.basic_xlwt_formatter import *


def write_format_xls(df,xls_name,sheet_name="Sheet1"):
    book, sheet1 = create_book(sheet_name)
    write_format_book(df, sheet1)
    folders = ["data/"]
    child_folder = check_folder(folders)
    book.save(folders[0] + child_folder + "/" +xls_name + ".xls")

def write_format_book(data, sheet1):
    if data.empty:
        print("数据为空")
        return

    sheet1.write(0, 0, "", set_style('Times New Roman', 220, True))
    row0 = data.columns.tolist()
    column0 = data.index.tolist()
    columns_count = len(row0)
    rows_count = len(column0)

    # 第一行
    tall_style = xlwt.easyxf('font:height 360;')
    first_row = sheet1.row(0)
    first_row.set_style(tall_style)
    for i in range(1, columns_count + 1):
        sheet1.write(0, i, row0[i - 1], set_style('Times New Roman', 220, True, True))  # 第一行

    # 第一列
    sheet1.col(0).width = 180 * 25
    for i in range(0, rows_count):
        sheet1.write(i + 1, 0, column0[i], set_style('Arial', 220, True))  # 第一列

    # 其他数据
    columns_format = pd.read_excel("config/__column_format.xlsx",header=0)
    columns_format.set_index(columns_format["column_name"], drop=True, inplace=True)
    format_column_list = columns_format.index.tolist()
    content_style = set_style('Arial', 200)
    for i in range(0, columns_count):
        for j in range(1, rows_count + 1):
            cell_data = data[row0[i]][j - 1]
            if u"日期" in row0[i]:
                if cell_data:
                    # print cell_data
                    cell_data = dt.datetime.strptime(cell_data, '%Y-%m-%d') if isinstance(cell_data, unicode) else cell_data
                    if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                        dateFormat = copy.deepcopy(content_style)
                        dateFormat.num_format_str = 'yyyy/mm/dd'
                        sheet1.write(j, i + 1, cell_data, dateFormat)
                    else:
                        sheet1.write(j, i + 1, "", content_style)
                else:
                    sheet1.write(j, i + 1, "", content_style)
                continue
            cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
            sheet1.write(j, i + 1, cell_data, content_style)  # 第一列

    for index, column_name in enumerate(row0):
        # print "index",index,column_name
        if column_name in format_column_list:
            column_width = columns_format["width"][column_name]
            # print "column_width",index,column_width
            sheet1.col(index+1).width = int(column_width) * 25
        else:
            sheet1.col(index+1).width = 210 * 25




