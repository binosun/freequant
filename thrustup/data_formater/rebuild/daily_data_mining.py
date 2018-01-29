# coding: utf-8

import os
import xlwt
import json
import numpy as np
import datetime as dt
import pandas as pd
from WindPy import *

w.start()
today_str = dt.datetime.today().strftime("%Y%m%d")

data_tag = "daily_data_output/" + today_str
today_file = data_tag + u"业绩披露.xlsx"
print today_file

def get_raw_df():
    # print today_file
    if not os.path.exists(today_file):
        print(u"未找到数据文件"+ today_file)
    else:
        print (u"从文件中得到股票数据"+ today_file)
        raw_df = pd.read_excel(today_file, header=0)
        raw_df.set_index(raw_df[u"证券代码"], drop=True, inplace=True)
        del raw_df[u"序号"]
        del raw_df[u"证券类型"]
        del raw_df[u"事件大类"]
        del raw_df[u"事件类型"]
        del raw_df[u"披露日期"]
        del raw_df[u"证券代码"]

        raw_df = raw_df.dropna(axis=0, how='any')
        return raw_df

def make_output_df(df):

    date_dict = calculate_date()

    stocks = list(set(df.index.tolist()))
    industry_result = w.wss(stocks, "industry_sw", "industryType=3")

    # industry_result = w.wss(stocks, "sec_name,industry_sw,tot_oper_rev,wgsd_yoy_or,wgsd_net_inc,yoyprofit",
    #                         "industryType=3;unit=1;rptDate="+date_dict["rpt_date_str"]+";rptType=1;currencyType=")
    # # print "industry_result",industry_result
    industry_data_dt = pd.DataFrame(industry_result.Data, index=industry_result.Fields, columns=industry_result.Codes).T
    print industry_data_dt

    # industry_data_dt["TOT_OPER_REV"] = industry_data_dt["TOT_OPER_REV"]/1e8
    # industry_data_dt["WGSD_NET_INC"] = industry_data_dt["WGSD_NET_INC"] / 1e8
    #
    # industry_data_dt.rename(columns={"SEC_NAME": u"证券名称", "INDUSTRY_SW": u"行业","TOT_OPER_REV":u"营业总收入(亿元)",
    #                                  "WGSD_YOY_OR":u"营收同比(%)","WGSD_NET_INC":u"净利润(亿元)","YOYPROFIT":u"净利润同比(%)"}, inplace=True)

    industry = industry_data_dt["INDUSTRY_SW"]
    df.insert(1, "申万三级行业名称", industry)
    df.sort_values(by="申万三级行业名称", ascending=False, inplace=True)

    new_df = split_df_column(df)

    # new_df = industry_data_dt


    qfa_result = w.wss(stocks, "qfa_cgrgr,qfa_cgrprofit", "rptDate="+date_dict["rpt_date_str"])
    qfa_data_dt = pd.DataFrame(qfa_result.Data, index=qfa_result.Fields, columns=qfa_result.Codes).T
    # qfa_yoygr_data = qfa_data_dt["QFA_YOYGR"]
    # new_df[u"单季度营收同比(%)"]=qfa_yoygr_data
    qfa_cgrgrr_data = qfa_data_dt["QFA_CGRGR"]
    new_df[u"单季度营收环比(%)"]=qfa_cgrgrr_data
    # qfa_yoyprofitr_data = qfa_data_dt["QFA_YOYPROFIT"]
    # new_df[u"单季度净利润同比(%)"]=qfa_yoyprofitr_data
    qfa_cgrprofitr_data = qfa_data_dt["QFA_CGRPROFIT"]
    new_df[u"单季度净利润环比(%)"]=qfa_cgrprofitr_data

    last_qfa_result = w.wss(stocks, "wgsd_yoy_or,yoyprofit", "rptDate="+date_dict["last_rpt_date_str"])
    last_qfa_data_dt = pd.DataFrame(last_qfa_result.Data, index=last_qfa_result.Fields, columns=last_qfa_result.Codes).T
    last_qfa_yoygr_data = last_qfa_data_dt["WGSD_YOY_OR"]
    new_df[u"上期营收同比(%)"]=last_qfa_yoygr_data

    last_qfa_yoyprofitr_data = last_qfa_data_dt["YOYPROFIT"]
    new_df[u"上期净利润同比(%)"]=last_qfa_yoyprofitr_data

    new_df[u"营收同比增长率变化(%)"] = (new_df[u"上期营收同比(%)"]/new_df[u"营收同比(%)"]-1)*100
    new_df[u"净利润同比增长率变化(%)"] = (new_df[u"上期净利润同比(%)"] / new_df[u"净利润同比(%)"]-1)*100

    del new_df[u"上期营收同比(%)"]
    del new_df[u"上期净利润同比(%)"]

    new_df = new_df.replace(np.nan, "")
    return new_df

def split_df_column(split_df):
    op_dict = {}
    yoy_op_dict = {}
    net_profit_dict = {}
    yoy_net_profit_dict = {}
    eps_dict = {}
    roe_dict = {}
    
    events = split_df[u"事件摘要"]
    # print "events",events
    for stock in split_df.index.tolist():
        single_str = events[stock]
        # print "single_str",type(single_str),single_str
        split_result = single_str.split(u"，")
        split_result.pop(0)
        for index,part in enumerate(split_result):
            # print "split_result",split_result
            bk_part = part
            if index == 0:
                digit = part[7:-1]
            elif index == 1:
                print stock,part
                digit = float(part[4:-2])
            elif index == 2:
                digit = part[4:-2]
            elif index == 3:
                digit = float(part[4:-1])
            elif index == 4:
                digit = float(part[8:-1])
            elif index == 5:
                digit = float(part[8:-1])
            # digit = int(filter(str.isdigit, part.encode('gbk')))
            # if "-" in bk_part:
            #     new_digit = 0-digit
            #     print "digit",digit,type(digit)
            # print bk_part,"-" in bk_part,"index",index,"digit",digit
            if index == 0:
                op_dict[stock] = digit
            elif index == 1:
                yoy_op_dict[stock] = digit
            elif index == 2:
                net_profit_dict[stock] = digit
            elif index == 3:
                yoy_net_profit_dict[stock] = digit
            elif index == 4:
                eps_dict[stock] = digit
            elif index == 5:
                roe_dict[stock] = digit
    split_df[u"营业总收入(元)"] = pd.Series(op_dict)
    split_df[u"营收同比(%)"] = pd.Series(yoy_op_dict)
    split_df[u"净利润(元)"] = pd.Series(net_profit_dict)
    split_df[u"净利润同比(%)"] = pd.Series(yoy_net_profit_dict)
    split_df[u"基本EPS(元)"] = pd.Series(eps_dict)
    split_df[u"加权平均ROE(%)"] = pd.Series(roe_dict)

    del split_df[u"事件摘要"]
    return split_df

def calculate_date():
    today = dt.datetime.today()
    today_str = today.strftime("%Y-%m-%d")
    year_str = today.strftime("%Y")
    month_str = today.strftime("%m")
    day_str = today.strftime("%d")
    date_list = ["1231", "0331", "0630", "0930"]
    flag = "A"

    trade_date = w.tdaysoffset(0, today_str, "").Data[0][0]

    if flag == "A":
        if 3 - int(month_str)>= 0:
            if  day_str != "31":
                rpt_date_str = str(int(year_str) - 1) + date_list[0]
            else:
                rpt_date_str = year_str + date_list[1]
        elif 6 - int(month_str)>= 0:
            if day_str != "30":
                rpt_date_str = year_str + date_list[1]
            else:
                rpt_date_str = year_str + date_list[2]
        elif 9 - int(month_str)>= 0:
            if day_str != "30":
                rpt_date_str = year_str + date_list[2]
            else:
                rpt_date_str = year_str + date_list[3]
        elif 12 - int(month_str)>= 0:
            if day_str != "31":
                rpt_date_str = year_str + date_list[3]
            else:
                rpt_date_str = year_str + date_list[3]

        index = date_list.index(rpt_date_str[4:])
        if index == 0:
            last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + date_list[3]
        else:
            last_rpt_date_str = rpt_date_str[0:4] + date_list[index - 1]

    else:
        if 6 - int(month_str) >= 0:
            if day_str != "30":
                rpt_date_str = str(int(year_str) - 1) + date_list[0]
            else:
                rpt_date_str = year_str + date_list[2]
        else:
            rpt_date_str = year_str + date_list[2]

        if rpt_date_str[4:] == date_list[2]:
            last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + date_list[0]
        else:
            last_rpt_date_str = rpt_date_str[0:4] + date_list[0]

    date = {}
    date["last_year_last_rpt_date_str"] = str(int(year_str)-1) + date_list[0]
    date["trade_date"] = trade_date
    date["rpt_date_str"] = rpt_date_str
    date["year_str"] = year_str
    date["today_str"] = today_str
    date["last_rpt_date_str"] = last_rpt_date_str
    # print date
    return date

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
    sheet1 = book.add_sheet("sheet1")
    return book, sheet1


def write_ih_book(data, sheet1):
    if data.empty:
        print("数据为空")
        return

    sheet1.write(0, 0, "", set_style('Times New Roman', 220, True))
    row0 = data.columns.tolist()
    column0 = data.index.tolist()
    columns_count = len(row0)
    rows_count = len(column0)

    for i in range(1, columns_count + 1):
        sheet1.write(0, i, row0[i - 1], set_style('Times New Roman', 220, True, True))  # 第一行
    for i in range(0, rows_count):
        sheet1.write(i + 1, 0, column0[i], set_style('Arial', 220, True))  # 第一列

    content_style = set_style('Arial', 200)
    for i in range(0, columns_count):
        for j in range(1, rows_count + 1):
            cell_data = data[row0[i]][j - 1]
            cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
            sheet1.write(j, i + 1, cell_data, content_style)
        sheet1.col(0).width = 150 * 20
        sheet1.col(1).width = 150 * 20
        sheet1.col(2).width = 200 * 20
        # sheet1.col(3).width = 150 * 20

        sheet1.col(3).width = 230 * 20
        sheet1.col(4).width = 200 * 20
        sheet1.col(5).width = 200 * 20
        sheet1.col(6).width = 200 * 20
        sheet1.col(7).width = 250 * 20
        sheet1.col(8).width = 300 * 20

        sheet1.col(9).width = 320 * 20
        sheet1.col(10).width = 360 * 20


        tall_style = xlwt.easyxf('font:height 360;')
        first_row = sheet1.row(0)
        first_row.set_style(tall_style)

def main():
    print "main"
    raw_df = get_raw_df()
    output_df = make_output_df(raw_df)
    book, sheet1 = create_book()
    write_ih_book(output_df, sheet1)
    book.save(data_tag + u"最新业绩披露.xls")

if __name__ == "__main__":
    main()

# 准备工作：
# 1.每天早晨，万得公司行动事件汇总，导出为“20171017业绩披露.xlsx”格式，
    # 放在目录F:\binger\freequant\thrustup\data_formater\rebuild\daily_data_output下
# 2.运行本脚本，自动生成新数据文件，例如“20171017最新业绩披露.xls”