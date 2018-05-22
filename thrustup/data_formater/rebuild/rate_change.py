# coding: utf-8

import os
import xlwt
import json
import datetime as dt
import pandas as pd
from WindPy import *

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

w.start()
today_str = dt.datetime.today().strftime("%Y-%m-%d")
data_tag = "data_output//" + today_str + "//"

def get_north_rate_df(tag,this_week,last_week):
    tag_list = [u"沪港通持股.xlsx", u"深港通持股.xlsx"]

    if tag == "SH":
        this_week_file = data_tag + this_week + tag_list[0]
        last_week_file = data_tag+ last_week + tag_list[0]
    elif tag == "SZ":
        this_week_file = data_tag + this_week + tag_list[1]
        last_week_file = data_tag+ last_week + tag_list[1]

    if not os.path.exists(this_week_file):
        print(u"未找到数据文件"+ this_week_file)
    else:
        print (u"从文件中得到股票数据"+ this_week_file)
        this_df = pd.read_excel(this_week_file, header=0)
        this_df.set_index(this_df[u"证券代码"], drop=True, inplace=True)
        this_df_index = this_df.index.tolist()[0:-2]


    if not os.path.exists(last_week_file):
        print(u"未找到数据文件"+ last_week_file)
    else:
        print (u"从文件中得到股票数据"+ last_week_file)
        last_df = pd.read_excel(last_week_file, header=0)
        last_df.set_index(last_df[u"证券代码"], drop=True, inplace=True)

    last_df_index = last_df.index.tolist()[0:-2]
    intersect_codes = list(set(this_df_index) & set(last_df_index))  # 取交集
    result_df = pd.DataFrame(index=intersect_codes)

    name_dict = {}
    last_dict = {}
    this_dict = {}
    rate_dict = {}
    for stock in intersect_codes:
        name_dict[stock] = last_df[u"证券简称"][stock]
        last_dict[stock] = last_df[u"占流通A股(%)(计算)"][stock]
        this_dict[stock] = this_df[u"占流通A股(%)(计算)"][stock]
        rate_dict[stock] = (this_dict[stock]-last_dict[stock])

    result_df[u"证券简称"] = pd.Series(name_dict)
    result_df[u"上周占流通A股(%)"] = pd.Series(last_dict)
    result_df[u"本周占流通A股(%)"] = pd.Series(this_dict)
    result_df[u"持仓周增(%)"] = pd.Series(rate_dict)
    result_df = result_df[result_df[u"本周占流通A股(%)"] > 1]
    result_df.sort_values(by=u"持仓周增(%)", ascending=False, inplace=True)

    industry_result = w.wss(result_df.index.tolist(), "industry_sw", "industryType=1")
    last_codes = industry_result.Codes
    last_fields = industry_result.Fields
    last_data = industry_result.Data
    last_data_dt = pd.DataFrame(last_data, index=last_fields, columns=last_codes).T

    industry = last_data_dt["INDUSTRY_SW"]
    result_df.insert(1, "申万一级行业", industry)
    return result_df

def get_south_rate_df(this_week, last_week):
    south_tag = u"沪深港通持股.xlsx"
    this_week_file = data_tag+this_week + south_tag
    last_week_file = data_tag+last_week + south_tag

    if not os.path.exists(this_week_file):
        print(u"未找到数据文件"+ this_week_file)
    else:
        print (u"从文件中得到股票数据"+ this_week_file)
        this_df = pd.read_excel(this_week_file, header=0)
        this_df.set_index(this_df[u"证券代码"], drop=True, inplace=True)
        this_df_index = this_df.index.tolist()[0:-2]


    if not os.path.exists(last_week_file):
        print(u"未找到数据文件"+ last_week_file)
    else:
        print (u"从文件中得到股票数据"+ last_week_file)
        last_df = pd.read_excel(last_week_file, header=0)
        last_df.set_index(last_df[u"证券代码"], drop=True, inplace=True)

    last_df_index = last_df.index.tolist()[0:-2]
    intersect_codes = list(set(this_df_index) & set(last_df_index))  # 取交集
    result_df = pd.DataFrame(index=intersect_codes)

    name_dict = {}
    last_dict = {}
    this_dict = {}
    rate_dict = {}
    for stock in intersect_codes:
        name_dict[stock] = last_df[u"证券简称"][stock]
        last_dict[stock] = last_df[u"占港股总股数(%)(计算)"][stock]
        this_dict[stock] = this_df[u"占港股总股数(%)(计算)"][stock]
        rate_dict[stock] = (this_dict[stock]-last_dict[stock])

    result_df[u"证券简称"] = pd.Series(name_dict)
    result_df[u"上周占港股总股数(%)"] = pd.Series(last_dict)
    result_df[u"本周占港股总股数(%)"] = pd.Series(this_dict)
    result_df[u"持仓周增(%)"] = pd.Series(rate_dict)
    result_df = result_df[result_df[u"本周占港股总股数(%)"] > 1]
    result_df.sort_values(by=u"持仓周增(%)", ascending=False, inplace=True)


    industry_result = w.wss(result_df.index.tolist(), "industry_HS","category=1")
    last_codes = industry_result.Codes
    last_fields = industry_result.Fields
    last_data = industry_result.Data
    last_data_dt = pd.DataFrame(last_data, index=last_fields, columns=last_codes).T

    industry = last_data_dt["INDUSTRY_HS"]
    result_df.insert(1, u"恒生一级行业", industry)
    return result_df


def set_style(name, height, bold=False, pattern_switch=False, alignment_center=True):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 20
    font.height = height
    font.italic = False

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

def create_book():
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("陆股通持股变动")
    return book,sheet1

def write_ih_book(data,sheet1):
    # if not data:
    #     print("数据为空")
    #     return
    if data.empty:
        print("数据为空")
        return
    data.sort_values(by=u"持仓周增(%)", ascending=False, inplace=True)
    sheet1.write(0, 0, "", set_style('Times New Roman', 220, True))
    row0 = data.columns.tolist()
    column0 = data.index.tolist()
    columns_count = len(row0)
    rows_count = len(column0)


    for i in range(1, columns_count + 1):
        sheet1.write(0, i, row0[i - 1], set_style('Times New Roman', 220, True, True))  # 第一行
    for i in range(0, rows_count):
        sheet1.write(i + 1, 0, column0[i], set_style('Arial', 220, True))  # 第一列
    for i in range(0, columns_count):
        for j in range(1, rows_count + 1):
            cell_data = data[row0[i]][j - 1]
            cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
            if (i==2 or i==3) and (cell_data>10):
                # sheet1.write(j, i + 1, cell_data, set_style('Arial', 200, True, True, True))
                sheet1.write(j, i + 1, cell_data)
            else:
                sheet1.write(j, i + 1, cell_data)

        sheet1.col(0).width = 160 * 25
        sheet1.col(1).width = 160 * 25
        sheet1.col(2).width = 250 * 25
        sheet1.col(3).width = 250 * 25
        sheet1.col(4).width = 250 * 25
        sheet1.col(5).width = 250 * 25
        
        tall_style = xlwt.easyxf('font:height 360;')
        first_row = sheet1.row(0)
        first_row.set_style(tall_style)


def north_direction(this_week, last_week):

    df_list = []
    for tag in ["SH","SZ"]:
        rate_df = get_north_rate_df(tag, this_week, last_week)
        df_list.append(rate_df)

    rate_df = pd.concat(df_list, axis=0)
    book,sheet1 = create_book()
    write_ih_book(rate_df, sheet1)
    book.save(data_tag + "north_direction_rate_change" + today_str + ".xls")


def south_direction(this_week, last_week):

    rate_df = get_south_rate_df(this_week, last_week)
    # rate_df.to_excel(data_tag + "south_direction_rate_change" + today_str + ".xls", float_format = '%.2f')
    book,sheet1 = create_book()
    write_ih_book(rate_df, sheet1)
    book.save(data_tag + "south_direction_rate_change" + today_str + ".xls")

def rate_change_main():
    this_week = raw_input(u"请输入本周日期：")
    last_week = raw_input(u"请输入上周日期：")
    north_direction(this_week, last_week)
    south_direction(this_week, last_week)


if __name__ == "__main__":
    rate_change_main()

# 准备工作：
# 1.每周，万得-我的-收藏夹-我收藏的专题统计报表，导出为“20171013沪港通持股.xlsx”、“20171013深港通持股.xlsx”格式，
    # 放在目录F:\binger\freequant\thrustup\data_formater\rebuild\data_output\下
# 2.运行本脚本，自动生成新数据文件，例如“rate_change2017-10-17.xls”