# -*- coding: utf-8 -*-

import json
import copy
from collections import OrderedDict
import os
import pandas as pd
import datetime as dt
from pandas import DataFrame
import xlwt
import numpy as np


import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)

from WindPy import *

class DataFormatter(object):

    def __init__(self,flag, industry, group):
        self.date_list = ["1231", "0331", "0630", "0930"]
        self.dir_dict = {"data_folder": "data_output/"}
        self.flag = flag
        self.industry = industry
        self.group = group
        self.group_in_chinese = ""
        self.book = None

        self.init()
        self.xls_cleanData_filename = self.group + "_" + self.flag +  self.date["today_str"]+ ".xls"
        self.date_mark = self.dir_dict["data_folder"] + self.date["today_str"] + "/"
        self.mk_dir()

    def init(self):
        self.start_engine()
        self.get_config()
        self.calculate_date()


    def start_engine(self):
        w.start()
        print "start engine finished"

    def get_config(self):
        with open("config.json", "r") as f:
            configuration = json.load(f, object_pairs_hook=OrderedDict)
        self.config = configuration["H_config"] if self.flag == "H" else configuration["A_config"]
        self.email_config = configuration["email_config"]

    def set_style(self, name, height, bold=False, pattern_switch=False, alignment_switch=True):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name  # 'Times New Roman'
        font.bold = bold
        font.color_index = 20
        font.height = height
        font.italic = False
        # font.struck_out = True
        # font.outline = True
        # font.shadow = True

        borders = xlwt.Borders()  # 边框
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1

        alignment = xlwt.Alignment()  # 居中
        alignment.horz = 0x02
        alignment.vert = 0x01
        # alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行

        pattern = xlwt.Pattern()  # 模式
        pattern.pattern = pattern_switch
        # pattern.pattern = 0x01
        pattern.pattern_fore_colour = 52
        pattern.pattern_back_colour = 52

        style.font = font
        style.borders = borders
        if alignment_switch:
            style.alignment = alignment
        style.pattern = pattern
        return style

    def mk_dir(self):

        for folder in self.dir_dict:
            if not os.path.exists(self.dir_dict[folder]):
                print "创建文件目录", self.dir_dict[folder]
                os.mkdir(self.dir_dict[folder])

        if not os.path.exists(self.date_mark):
            print "创建文件目录", self.date_mark
            os.mkdir(self.date_mark)

    def get_all_stocks(self, stock_file_name):
        # 获取所有股票
        print("stock_file_name", stock_file_name)
        if os.path.exists(stock_file_name):
            print (u"从文件中得到股票数据")
            with open(stock_file_name, "r") as f:
                stocks = json.load(f, object_pairs_hook=OrderedDict)
                stock_codes = stocks["codes"]
        else:
            today_str = dt.datetime.today().strftime("%Y-%m-%d")
            if self.flag == "H":
                query_key = "date=" + today_str + ";sectorid=a002010100000000"  # 港股
            else:
                query_key = "date=" + today_str + ";sectorid=a001010100000000"
            stocks = w.wset("sectorconstituent", query_key)
            stock_codes = stocks.Data[1]
            # stock_names = stocks.Data[2]
            stocks_json = OrderedDict()
            stocks_json["codes"] = stock_codes
            fh = open(stock_file_name, 'w')
            fh.write(json.dumps(stocks_json))
            fh.close()
        print(u"全部共 %s 只股票" % len(stock_codes))
        return stock_codes

    def calculate_date(self):
        today = dt.datetime.today()
        import time
        now_time = time.strptime(today.strftime("%H:%M"), "%H:%M")
        if now_time < time.strptime("16:20", "%H:%M"):
            today = today - dt.timedelta(days=1)

        today_str = today.strftime("%Y-%m-%d")
        year_str = today.strftime("%Y")
        month_str = today.strftime("%m")
        day_str = today.strftime("%d")

        trade_date = w.tdaysoffset(0, today_str, "").Data[0][0]

        if self.flag == "A":
            if 3 - int(month_str)>= 0:
                if  day_str != "31":
                    rpt_date_str = str(int(year_str) - 1) + self.date_list[3]
                else:
                    rpt_date_str = year_str + self.date_list[1]
            elif 6 - int(month_str)>= 0:
                if day_str != "30":
                    rpt_date_str = year_str + self.date_list[1]
                else:
                    rpt_date_str = year_str + self.date_list[2]
            elif 9 - int(month_str)>= 0:
                if day_str != "30":
                    rpt_date_str = year_str + self.date_list[2]
                else:
                    rpt_date_str = year_str + self.date_list[3]
            elif 12 - int(month_str)>= 0:
                if day_str != "31":
                    rpt_date_str = year_str + self.date_list[3]
                else:
                    rpt_date_str = year_str + self.date_list[3]

            index = self.date_list.index(rpt_date_str[4:])
            if index == 0:
                last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + self.date_list[3]
            else:
                last_rpt_date_str = rpt_date_str[0:4] + self.date_list[index - 1]

        else:
            if 6 - int(month_str) >= 0:
                if day_str != "30":
                    rpt_date_str = str(int(year_str) - 1) + self.date_list[0]
                else:
                    rpt_date_str = year_str + self.date_list[2]
            else:
                rpt_date_str = year_str + self.date_list[2]

            if rpt_date_str[4:] == self.date_list[2]:
                last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + self.date_list[0]
            else:
                last_rpt_date_str = rpt_date_str[0:4] + self.date_list[0]

        self.date = {}
        self.date["last_year_last_rpt_date_str"] = str(int(year_str)-1) + self.date_list[0]
        self.date["trade_date"] = trade_date
        self.date["rpt_date_str"] = rpt_date_str
        self.date["year_str"] = year_str
        self.date["today_str"] = today_str
        self.date["last_rpt_date_str"] = last_rpt_date_str
        return self.date

    def get_query_stocks(self, stocks, main_config):
        print "in function: get_query_stocks"
        if self.group == "history_high":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且最高价最近" + self.config["param_n"] + u"天创历史新高的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "历史新高"
            return self.get_history_high_stocks(stocks, main_config)

        elif self.group == "stage_high":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且最高价最近" + self.config["param_n"] + u"天创"+self.config["param_m"]+u"天新高的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "阶段新高"
            return self.get_stage_high_stocks(stocks, main_config)
        elif self.group == "stage_low":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且最高价最近" + self.config["param_n"] + u"天创"+self.config["param_m_low"]+u"天新低的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "阶段新低"
            return self.get_stage_low_stocks(stocks, main_config)
        elif self.group == "quarter_increase":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且营收、利润、毛利率环比加速的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "季度加速"
            return self.get_quarter_increase_stocks(stocks)
        elif self.group == "increase_holding":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且本周发生股东增持事件的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "增持"
            return self.get_increase_holding_stocks(main_config)
        elif self.group == "share_pledged":
            self.remarks = u"备注：" \
                           u"1.该表列出上市满1年并且质押比例超过50%的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "质押比"
            return self.get_share_pledged_stocks(stocks)
        elif self.group == "thrust_up_plate":
            self.remarks = u"备注：" \
                           u"1.该表列出上市超过三年且最近三个月发生平台突破的股票及相关数据；" \
                           u"2.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "底部突破"
            return self.get_thrust_up_plate_stocks(stocks)
        elif self.group == "peg_pick":
            self.remarks = u"备注：" \
                           u"1.表中财务数据均为当前报告期数据，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"2.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "peg选股"
            return self.get_peg_pick_stocks(stocks)
        elif self.group == "small_cap_undervalue":
            self.remarks = u"备注：" \
                           u"1.本表从全部A股中选择低估小市值股票，筛选标准为(市值<50亿，PE(TTM)<30，PB<4，营收同比>0，净利润同比>0)；" \
                           u"2.表中财务数据均为当前报告期数据(预测数据除外)，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.group_in_chinese = "低估小盘股"
            return self.get_small_cap_undervalued_stocks(stocks)
        elif self.group == "MA60":
            self.remarks = u"备注：" \
                           u"统计时剔除停牌和上市不足60天的股票。"
            self.group_in_chinese = "行业MA60"
            return self.get_MA60_pic(stocks)
        else:
            self.remarks = ""
            self.group_in_chinese = ""
            return stocks


    def get_data(self, stocks, main_config):
        date = self.date
        trade_date = date["trade_date"]
        rpt_date_str = date["rpt_date_str"]
        today_str = date["today_str"]
        year_str = date["year_str"]

        param_n = main_config["param_n"]
        delta_days = main_config["delta_days"]
        if self.flag == "H":
            params = "tradeDate=" + trade_date.strftime("%Y%m%d") + ";priceAdj=F;cycle=D;n=" + param_n + \
                     ";ruleType=9;rptDate=" + rpt_date_str + ";category=1;year=" + year_str
        else:
            params = "tradeDate="+trade_date.strftime("%Y%m%d")+";priceAdj=F;cycle=D;n="+param_n+\
                     ";ruleType=9;rptDate="+rpt_date_str+";industryType=1;year="+year_str
        result = w.wss(stocks, main_config["field_map"].keys(), params)
        codes = result.Codes
        fields = result.Fields
        data = result.Data
        limit_date = trade_date - dt.timedelta(days=delta_days)
        df_data = DataFrame(data, index=fields, columns=codes).T

        if self.flag == "A":
            xq_fields = ["xq_wow_focus", "xq_accmfocus"]
            xq_trade_date = w.tdaysoffset(-2, today_str, "").Data[0][0]
            xq_params = "tradeDate=" + xq_trade_date.strftime("%Y%m%d")
            xq_result = w.wss(codes, xq_fields, xq_params)
            xq_codes = xq_result.Codes
            xq_fields = xq_result.Fields
            xq_data = np.array(xq_result.Data)
            where_are_nan = np.isnan(xq_data)
            where_are_inf = np.isinf(xq_data)
            xq_data[where_are_nan] = 0
            xq_data[where_are_inf] = 0
            xq_df_data = DataFrame(xq_data, index=xq_fields, columns=xq_codes).T

            for column in xq_df_data.columns.tolist():
                for row in xq_df_data.index.tolist():
                    if xq_df_data[column][row] is not np.nan:
                        df_data[column][row] = xq_df_data[column][row]

        # df_data = df_data[df_data["IPO_DATE"] < limit_date]
        df_data = df_data.sort_values(by=self.industry)
        del df_data["HISTORY_HIGH"]
        del df_data["IPO_DATE"]
        return self.__completed_last_rpt_data(df_data)

    def __completed_last_rpt_data(self, df_data):

        ROE = df_data["ROE"]
        lose_list = []
        values = ROE.values.tolist()
        index = ROE.index.tolist()
        for key, value in enumerate(values):
            if not isinstance(value, float):
                lose_list.append(index[key])
        if lose_list:
            last_data_result = w.wss(lose_list,"roe,roa,debttoassets,current", "rptDate="+self.date["last_rpt_date_str"])
            last_codes = last_data_result.Codes
            last_fields = last_data_result.Fields
            last_data = last_data_result.Data

            last_data_dt = DataFrame(last_data,index=last_fields,columns=last_codes).T
            # last_data_dt = last_data_dt.replace("None", np.nan)
            # if not last_data_dt.empty:
            #     where_are_nan = np.isnan(last_data_dt)
            #     where_are_inf = np.isinf(last_data_dt)
            #     last_data_dt[where_are_nan] = 0
            #     last_data_dt[where_are_inf] = 0

            for column in last_data_dt.columns.tolist():
                for row in last_data_dt.index.tolist():
                    if (not df_data[column][row]) and (last_data_dt[column][row]):
                        df_data[column][row] = last_data_dt[column][row]

        field_map = {key.upper(): value for key, value in self.config["field_map"].items()}
        df_data.rename(columns=field_map, inplace=True)
        return df_data.replace(np.nan, "")

    def main(self):
        stocks = self.get_all_stocks(self.config["stocks_file"])
        query_stocks = self.get_query_stocks(stocks, self.config)
        if not query_stocks:
            print("待选股票池为空")
            return
        df_data = self.get_data(query_stocks, self.config)
        print "df_data",df_data
        self.create_and_write_book(df_data)
        self.save_book_as_xls()


    def create_and_write_book(self, data):
        if not self.book:
            self.book = xlwt.Workbook(encoding="utf-8")
        self.sheet1 = self.book.add_sheet(self.group_in_chinese+"-"+self.flag + "股")

        self.__write_into_book(data)

    def __write_into_book(self, data):
        self.sheet1.write(0, 0, "", self.set_style('Times New Roman', 220, True))
        row0 = data.columns
        column0 = data.index
        columns_count = len(row0)
        rows_count = len(column0)

        if self.flag == "H":
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)
            for j in range(1, rows_count + 1):
                for i in range(0, columns_count):
                    if i == 14:
                        cell_data = data[row0[i]][j - 1]
                        if cell_data:
                            if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                                dateFormat = copy.deepcopy(content_style)
                                dateFormat.num_format_str = 'yyyy/mm/dd'
                                self.sheet1.write(j, i + 1, cell_data, dateFormat)
                            else:
                                self.sheet1.write(j, i + 1, "", content_style)
                        else:
                            self.sheet1.write(j, i + 1, "", content_style)
                        continue
                    if i == 16:
                        cell_data = data[row0[i]][j - 1]
                        style = self.set_style('Arial', 200, False, False, False)
                        self.sheet1.write(j, i + 1, cell_data, style)
                        continue

                    cell_data = data[row0[i]][j - 1]
                    cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    self.sheet1.write(j, i + 1, cell_data, content_style)  # 第一列

            self.sheet1.col(0).width = 160 * 25
            self.sheet1.col(1).width = 180 * 25
            self.sheet1.col(2).width = 90 * 25
            self.sheet1.col(3).width = 360 * 25
            self.sheet1.col(4).width = 250 * 25
            self.sheet1.col(5).width = 150 * 25
            self.sheet1.col(6).width = 100 * 25
            self.sheet1.col(7).width = 180 * 25
            self.sheet1.col(8).width = 180 * 25
            self.sheet1.col(9).width = 150 * 25
            self.sheet1.col(10).width = 120 * 25
            self.sheet1.col(11).width = 300 * 25
            self.sheet1.col(12).width = 300 * 25
            self.sheet1.col(13).width = 360 * 25
            self.sheet1.col(14).width = 350 * 25
            self.sheet1.col(15).width = 200 * 25
            self.sheet1.col(16).width = 200 * 25
            self.sheet1.col(17).width = 2500 * 25
        else:
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)

            dateFormat = copy.deepcopy(content_style)
            dateFormat.num_format_str = 'yyyy/mm/dd'
            for j in range(1, rows_count + 1):
                for i in range(0, columns_count):
                    if i == 18:
                        cell_data = data[row0[i]][j - 1]
                        if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                            self.sheet1.write(j, i + 1, cell_data, dateFormat)
                        else:
                            self.sheet1.write(j, i + 1, "", content_style)
                        continue

                    cell_data = data[row0[i]][j - 1]
                    cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    self.sheet1.write(j, i + 1, cell_data, content_style)  # 第一列

            self.sheet1.col(0).width = 160 * 25
            self.sheet1.col(1).width = 120 * 25
            self.sheet1.col(2).width = 90 * 25
            self.sheet1.col(3).width = 290 * 25
            self.sheet1.col(4).width = 250 * 25
            self.sheet1.col(5).width = 360 * 25
            self.sheet1.col(6).width = 250 * 25
            self.sheet1.col(7).width = 320 * 25
            self.sheet1.col(8).width = 150 * 25
            self.sheet1.col(9).width = 320 * 25
            self.sheet1.col(10).width = 100 * 25
            self.sheet1.col(11).width = 180 * 25
            self.sheet1.col(12).width = 180 * 25
            self.sheet1.col(13).width = 150 * 25
            self.sheet1.col(14).width = 120 * 25
            self.sheet1.col(15).width = 320 * 25
            self.sheet1.col(16).width = 300 * 25
            self.sheet1.col(17).width = 360 * 25
            self.sheet1.col(18).width = 350 * 25
            self.sheet1.col(19).width = 200 * 25
            self.sheet1.col(20).width = 200 * 25
            self.sheet1.col(21).width = 700 * 25

        # self.sheet1.write_merge(rows_count + 1, rows_count + 1, 0, columns_count, self.remarks)
        self.sheet1.write(rows_count + 1, 0, self.remarks)
        tall_style = xlwt.easyxf('font:height 360;')
        first_row = self.sheet1.row(0)
        first_row.set_style(tall_style)

    def save_book_as_xls(self):
        file_name = self.date_mark + self.xls_cleanData_filename
        if not os.path.exists(file_name):
            self.book.save(file_name)
        else:
            for i in range(1, 10001):
                new_name = self.date_mark + "(" + str(i) + ")" + self.xls_cleanData_filename
                if not os.path.exists(new_name):
                    self.book.save(new_name)
                    break

    def save_df_to_excel(self,data_df):
        self.file_name = self.date_mark + self.xls_cleanData_filename
        if not os.path.exists(self.file_name):
            xlsx_output = pd.ExcelWriter(self.file_name)
            data_df.to_excel(xlsx_output, float_format = '%.2f')
            xlsx_output.save()
        else:
            for i in range(1, 10001):
                new_name = self.date_mark + "(" + str(i) + ")" + self.xls_cleanData_filename
                if not os.path.exists(new_name):
                    self.file_name = new_name
                    xlsx_output = pd.ExcelWriter(self.file_name)
                    data_df.to_excel(xlsx_output, float_format = '%.2f')
                    xlsx_output.save()
                    break


#####################################################################
    def get_history_high_stocks(self, stocks, main_config):

        trade_date = self.date["trade_date"]
        delta_days = main_config["delta_days"]
        limit_date = trade_date - dt.timedelta(days=delta_days)

        param_n = main_config["param_n"]
        history_high_params = "tradeDate="+trade_date.strftime("%Y%m%d")+";priceAdj=F;n="+param_n
        history_high_result = w.wss(stocks, "history_high,ipo_date", history_high_params)
        history_high_codes = history_high_result.Codes
        history_high_fields = history_high_result.Fields
        history_high_data = history_high_result.Data
        df_history_high_data = DataFrame(history_high_data, index=history_high_fields, columns=history_high_codes).T
        df_history_high_data = df_history_high_data[(df_history_high_data["HISTORY_HIGH"] == "TRUE")&(df_history_high_data["IPO_DATE"] < limit_date)]
        history_high_stocks = df_history_high_data.index.tolist()
        del df_history_high_data["IPO_DATE"]

        df_history_high_data = df_history_high_data.replace("TRUE", 1)
        last_history_high_stocks = history_high_stocks
        str_trade_date = trade_date.strftime("%Y%m%d")

        while last_history_high_stocks:

            str_trade_date = w.tdaysoffset(-1, str_trade_date, "Days=Alldays;Period=W").Data[0][0].strftime("%Y%m%d")
            last_history_high_params = "tradeDate=" + str_trade_date + ";priceAdj=F;n=" + param_n
            last_history_high_result = w.wss(last_history_high_stocks, "history_high", last_history_high_params)
            last_history_high_codes = last_history_high_result.Codes
            last_history_high_fields = last_history_high_result.Fields
            last_history_high_data = last_history_high_result.Data
            last_df_history_high_data = DataFrame(last_history_high_data, index=last_history_high_fields, columns=last_history_high_codes).T
            last_df_history_high_data = last_df_history_high_data[(last_df_history_high_data["HISTORY_HIGH"] == "TRUE")]
            last_history_high_stocks = last_df_history_high_data.index.tolist()
            last_df_history_high_data.rename(columns={"HISTORY_HIGH": str_trade_date}, inplace=True)
            last_df_history_high_data = last_df_history_high_data.replace("TRUE", 1)
            last_df_history_high_data = last_df_history_high_data.replace("FALSE", 0)

            df_history_high_data = pd.concat([df_history_high_data, last_df_history_high_data], axis=1)

        columns = df_history_high_data.columns.tolist()
        df_history_high_data = df_history_high_data.replace(np.nan, 0)
        df_history_high_data["count"] = 0
        for column in columns[0:-1]:
            df_history_high_data["count"] += df_history_high_data[column]
            print df_history_high_data

        xlsx_output = pd.ExcelWriter(self.date_mark + self.group + "_" + self.flag + "_" +"count" + self.date["today_str"] + ".xls")
        df_history_high_data.to_excel(xlsx_output)
        xlsx_output.save()
        return history_high_stocks

    def get_stage_high_stocks(self, stocks, main_config):
        trade_date = self.date["trade_date"]
        param_n = main_config["param_n"]
        param_m = main_config["param_m"]
        stage_high_params = "tradeDate="+trade_date.strftime("%Y%m%d")+";priceAdj=F;n="+param_n+";m="+param_m
        stage_high_result = w.wss(stocks, "stage_high", stage_high_params)
        stage_high_codes = stage_high_result.Codes
        stage_high_fields = stage_high_result.Fields
        stage_high_data = stage_high_result.Data
        df_stage_high_data = DataFrame(stage_high_data, index=stage_high_fields, columns=stage_high_codes).T
        stage_high_stocks = df_stage_high_data[(df_stage_high_data["STAGE_HIGH"] == "TRUE")].index.tolist()
        return stage_high_stocks

    def get_stage_low_stocks(self, stocks, main_config):
        trade_date = self.date["trade_date"]
        param_n = main_config["param_n"]
        param_m_low = main_config["param_m_low"]
        stage_low_params = "tradeDate="+trade_date.strftime("%Y%m%d")+";priceAdj=F;n="+param_n+";m="+param_m_low
        stage_low_result = w.wss(stocks, "stage_low", stage_low_params)
        stage_low_codes = stage_low_result.Codes
        stage_low_fields = stage_low_result.Fields
        stage_low_data = stage_low_result.Data
        df_stage_low_data = DataFrame(stage_low_data, index=stage_low_fields, columns=stage_low_codes).T
        stage_low_stocks = df_stage_low_data[(df_stage_low_data["STAGE_LOW"] == "TRUE")].index.tolist()
        return stage_low_stocks

    def get_quarter_increase_stocks(self, stocks):
        if self.flag == "A":
            quarter_increase_result = w.wss(stocks, "qfa_cgrsales, qfa_cgrgr", "rptDate="+self.date["rpt_date_str"])
            print("quarter_increase_result", quarter_increase_result)
            quarter_increase_codes = quarter_increase_result.Codes
            quarter_increase_fields = quarter_increase_result.Fields
            quarter_increase_data = quarter_increase_result.Data
            df_quarter_increase_data = DataFrame(quarter_increase_data, index=quarter_increase_fields, columns=quarter_increase_codes).T
            step1_stocks = df_quarter_increase_data[(df_quarter_increase_data["QFA_CGRSALES"] > 0) & (df_quarter_increase_data["QFA_CGRGR"] > 0)].index.tolist()
            print("step1_stocks", step1_stocks)
            qfa_gpm_result = w.wss(step1_stocks, "qfa_grossprofitmargin", "rptDate="+self.date["rpt_date_str"])
            last_qfa_gpm_result = w.wss(step1_stocks, "qfa_grossprofitmargin", "rptDate="+self.date["last_rpt_date_str"])
            print("qfa_gpm_result", qfa_gpm_result)
            print("last_qfa_gpm_result", last_qfa_gpm_result)
            qfa_gpm_data = DataFrame([qfa_gpm_result.Data[0], last_qfa_gpm_result.Data[0]], index=["qfa_gpm", "last_qfa_gpm"], columns=qfa_gpm_result.Codes).T
            print("qfa_gpm_data", qfa_gpm_data)
            quarter_increase_stocks = qfa_gpm_data[(qfa_gpm_data["qfa_gpm"] != np.nan) &
                                                   (qfa_gpm_data["last_qfa_gpm"] != np.nan) &
                                                   (qfa_gpm_data["qfa_gpm"] > qfa_gpm_data["last_qfa_gpm"])].index.tolist()

            print("quarter_increase_stocks", quarter_increase_stocks, len(quarter_increase_stocks))
            return quarter_increase_stocks
        else:
            quarter_increase_result = w.wss(stocks, "wgsd_grossprofitmargin,tot_oper_rev,wgsd_ebit_oper", "rptDate="+self.date["rpt_date_str"]+";unit=1;rptType=1;currencyType=")
            print("quarter_increase_result", quarter_increase_result)
            quarter_increase_codes = quarter_increase_result.Codes
            quarter_increase_fields = quarter_increase_result.Fields
            quarter_increase_data = quarter_increase_result.Data
            df_quarter_increase_data = DataFrame(quarter_increase_data, index=quarter_increase_fields, columns=quarter_increase_codes).T
            clean_df = df_quarter_increase_data[(df_quarter_increase_data["WGSD_GROSSPROFITMARGIN"] != np.nan) &
                                                   (df_quarter_increase_data["TOT_OPER_REV"] != np.nan) &
                                                    (df_quarter_increase_data["WGSD_EBIT_OPER"] != np.nan)]
            step1_stocks = clean_df.index.tolist()
            last_quarter_increase_result = w.wss(step1_stocks, "wgsd_grossprofitmargin,tot_oper_rev,wgsd_ebit_oper",
                                            "rptDate=" + self.date["last_rpt_date_str"]+";unit=1;rptType=1;currencyType=")
            print("last_quarter_increase_result", last_quarter_increase_result)
            last_quarter_increase_codes = last_quarter_increase_result.Codes
            # last_quarter_increase_fields = last_quarter_increase_result.Fields
            last_quarter_increase_data = last_quarter_increase_result.Data
            last_df_quarter_increase_data = DataFrame(last_quarter_increase_data, index=["last_WGSD_GROSSPROFITMARGIN", "last_TOT_OPER_REV", "last_WGSD_EBIT_OPER"], columns=last_quarter_increase_codes).T
            # clean_df = pd.concat([clean_df, last_df_quarter_increase_data], axis=1)

            cl_df = pd.DataFrame()
            cl_df["item1"] = (clean_df["WGSD_GROSSPROFITMARGIN"]-last_df_quarter_increase_data["last_WGSD_GROSSPROFITMARGIN"])
            cl_df["item2"] = (clean_df["TOT_OPER_REV"] - last_df_quarter_increase_data["last_TOT_OPER_REV"])
            cl_df["item3"] = (clean_df["WGSD_EBIT_OPER"] - last_df_quarter_increase_data["last_WGSD_EBIT_OPER"])
            cl_df = cl_df[(cl_df["item1"]> 0) &
                          (cl_df["item3"] > 0) &
                          (cl_df["item2"] > 0)]
            clean_stocks = cl_df.index.tolist()
            return clean_stocks

    def get_increase_holding_stocks(self, main_config):
        increase_holding_file = self.date_mark + main_config["increase_holding_file"]
        if not os.path.exists(increase_holding_file):
            print "未找到文件",increase_holding_file
            return
        else:
            ih_df = pd.read_excel(increase_holding_file,
                                                   header=0)
            del ih_df[u"披露日期"]
            stocks = ih_df.dropna(axis=0,how='any')[u"证券代码"].tolist()
            stocks_result = list(set(stocks))
            # print "stocks", stocks_result
            return stocks_result

    def get_share_pledged_stocks(self, stocks):
        wss_result = w.wss(stocks, "share_pledgeda_pct","tradeDate=" + self.date["today_str"])
        share_pledged_data = DataFrame(wss_result.Data, columns=wss_result.Codes, index= wss_result.Fields).T
        final_stocks = share_pledged_data[share_pledged_data["SHARE_PLEDGEDA_PCT"] > 50].index.tolist()

    def get_thrust_up_plate_stocks(self, stocks):

        # 上市超过三年
        # 最近三个月的最低价和三年内最高价相比，跌幅超过70%
        # 最近三个月的最低价日期与当前日期相比，超过45天
        # 当前价突破最近三个月最高价
        # 近三个月股价标准差小于0.45

        day_1 = 3*365
        day_2 = 91
        day_3 = 45
        percent_1 = 0.7

        limit_date = self.date["trade_date"] - dt.timedelta(days=day_1)
        request_result = w.wss(stocks,"high_per,low_per,ipo_date","priceAdj=F;startDate="+limit_date.strftime("%Y-%m-%d")+";endDate="+self.date["today_str"])
        ipo_data = DataFrame(request_result.Data, columns=request_result.Codes, index=request_result.Fields).T

        ipo_data.rename(columns={"HIGH_PER":"HISTORY_HIGH", "LOW_PER":"HISTORY_LOW"}, inplace=True)
        ipo_data = ipo_data[ipo_data["IPO_DATE"] < limit_date]

        susp_result = w.wss(ipo_data.index.tolist(),
                              "susp_days","tradeDate="+(self.date["trade_date"]- dt.timedelta(days=1)).strftime("%Y%m%d"))
        susp_data = DataFrame(susp_result.Data, columns=susp_result.Codes, index=susp_result.Fields).T
        ipo_data = pd.concat([ipo_data, susp_data], axis=1)
        ipo_data = ipo_data[ipo_data["SUSP_DAYS"] < 1]

        start_date = (self.date["trade_date"] - dt.timedelta(days=day_2)).strftime("%Y-%m-%d")
        prriod_result = w.wss(ipo_data.index.tolist(),
                              "high_per,low_per,low_date_per,close,vhf",
                              "priceAdj=F;startDate="+start_date+";endDate="+self.date["today_str"]+
                              ";cycle=D;tradeDate="+self.date["trade_date"].strftime("%Y-%m-%d")+";VHF_N="+ str(day_2))
        period_data = DataFrame(prriod_result.Data, columns=prriod_result.Codes, index=prriod_result.Fields).T

        clean_df = pd.concat([ipo_data, period_data], axis=1)
        clean_df["high_percent"] =clean_df["LOW_PER"]/clean_df["HISTORY_HIGH"]
        clean_df["low_percent"] = clean_df["LOW_PER"] / clean_df["HISTORY_LOW"]

        clean_df = clean_df[(clean_df["VHF"] > 0)&(clean_df["VHF"] < 0.3)&(clean_df["high_percent"] < percent_1)&
                            (clean_df["low_percent"]<1.5)&
                            (clean_df["LOW_DATE_PER"] < (self.date["trade_date"] - dt.timedelta(days=day_3)))&
                            (clean_df["CLOSE"] >= clean_df["HIGH_PER"])]

        stocks_result = clean_df.index.tolist()
        return stocks_result

    def get_peg_pick_stocks(self, stocks):
        text_trade_date = (self.date["trade_date"] - dt.timedelta(days=1)).strftime("%Y-%m-%d")

        str_after_two_year = str(int(self.date["year_str"])+2)
        str_after_one_year = str(int(self.date["year_str"]) + 1)

        after_two_year_data = w.wss(stocks,
                                 "pe_est_last,est_eps",
                                 "year="+str_after_two_year +";tradeDate=" + text_trade_date)
        after_two_year_df = DataFrame(after_two_year_data.Data, columns=after_two_year_data.Codes,
                                      index=["pe_est_last_"+str_after_two_year, "est_eps_"+str_after_two_year]).T
        after_two_year_df = after_two_year_df.dropna(axis=0, how='any')

        after_one_year_data = w.wss(after_two_year_df.index.tolist(),
                                 "pe_est_last,est_eps",
                                 "year="+str_after_one_year +";tradeDate=" + text_trade_date)
        after_one_year_df = DataFrame(after_one_year_data.Data, columns=after_one_year_data.Codes,
                                      index=["pe_est_last_"+str_after_one_year, "est_eps_"+str_after_one_year]).T
        this_year_data = w.wss(after_two_year_df.index.tolist(),
                                 "pe_est_last,est_eps",
                                 "year="+self.date["year_str"] +";tradeDate=" + text_trade_date)
        this_year_df = DataFrame(this_year_data.Data, columns=this_year_data.Codes,
                                      index=["pe_est_last_"+self.date["year_str"], "est_eps_"+self.date["year_str"]]).T

        peg_df = pd.concat([after_two_year_df, after_one_year_df, this_year_df], axis=1)

        peg_df["eps_acc"] = 0
        peg_df["peg"] = 0
        peg_df["eps_acc"] = ((peg_df["est_eps_"+str_after_two_year] / peg_df["est_eps_"+self.date["year_str"]]) ** (0.5) - 1) * 100
        peg_df["peg"] = peg_df["pe_est_last_"+self.date["year_str"]]/peg_df["eps_acc"]

        peg_df = peg_df[(peg_df["peg"] > 0) & (peg_df["peg"] < 2)]
        peg_df = peg_df[(peg_df["pe_est_last_"+self.date["year_str"]] > 0) & (peg_df["pe_est_last_"+self.date["year_str"]] < 60)]

        if self.flag == "A":
            gross_pick_result = w.wss(peg_df.index.tolist(),
                                      "eps_ttm,pb,yoy_or,yoyop,qfa_cgrsales,qfa_cgrop,qfa_yoysales",
                                      "rptDate=" + self.date["rpt_date_str"]+
                                      ";rptType=1;tradeDate="+text_trade_date+";ruleType=9")

            gross_pick_df = DataFrame(gross_pick_result.Data, columns=gross_pick_result.Codes, index=gross_pick_result.Fields).T

            gross_pick_df = gross_pick_df[(gross_pick_df["PB"]>0) & (gross_pick_df["PB"]<12)]
            gross_pick_df = gross_pick_df[(gross_pick_df["YOY_OR"]>0) & (gross_pick_df["QFA_CGRSALES"]>0)]
            gross_pick_df = gross_pick_df[(gross_pick_df["YOYOP"] > 0) & (gross_pick_df["QFA_CGROP"] > 0)]

            well_pick_result = w.wss(gross_pick_df.index.tolist(),
                                     "pe_ttm,sec_name,roa,roe,debttoassets,current,ocftosales,profittogr,industry_sw,"
                                     "tot_oper_rev,cashflow_ttm,yoynetprofit_deducted,"
                                     "qfa_yoyprofit,qfa_grossprofitmargin,qfa_netprofitmargin",
                                     "rptDate=" + self.date["rpt_date_str"] + ";unit=1;rptType=1;industryType=1;tradeDate=" + text_trade_date)

            well_pick_df = DataFrame(well_pick_result.Data, columns=well_pick_result.Codes,
                                      index=well_pick_result.Fields).T

            last_report_result = w.wss(gross_pick_df.index.tolist(),
                                     "qfa_yoysales,qfa_yoyprofit,qfa_grossprofitmargin,qfa_netprofitmargin",
                                     "rptDate=" + self.date["last_rpt_date_str"])
            last_report_df = DataFrame(last_report_result.Data, columns=last_report_result.Codes,
                                      index=last_report_result.Fields).T
            last_report_df.rename(columns={"QFA_YOYSALES": "LAST_QFA_YOYSALES", "QFA_YOYPROFIT": "LAST_QFA_YOYPROFIT",
                                           "QFA_GROSSPROFITMARGIN":"LAST_QFA_GROSSPROFITMARGIN","QFA_NETPROFITMARGIN":"LAST_QFA_NETPROFITMARGIN"}, inplace=True)

            # 获取上年年报的营收同比
            last_year_report_result = w.wss(gross_pick_df.index.tolist(),
                                       "yoy_or","rptDate=" + self.date["last_year_last_rpt_date_str"])
            last_year_report_result = DataFrame(last_year_report_result.Data, columns=last_year_report_result.Codes,
                                      index=last_year_report_result.Fields).T
            last_year_report_result.rename(columns={"YOY_OR":"LAST_YEAR_YOY_OR"}, inplace=True)

            score_df = pd.concat([peg_df, gross_pick_df, well_pick_df,last_report_df,last_year_report_result], axis=1)
            score_df = score_df.dropna(axis=0, how='any')

            score_df["OCFTOPROFIT"] = score_df["OCFTOSALES"] / score_df["PROFITTOGR"]

            score_df["pe_score"] = 0
            score_df["pb_score"] = 0

            score_df["roa_score"] = 0
            score_df["roe_score"] = 0

            score_df["debttoassets_score"] = 0
            score_df["current_score"] = 0

            score_df["ocf_score"] = 0
            score_df["cashflow_score"] = 0
            score_df["ocftoprofit_score"] = 0

            score_df["yoy_or_score"] = 0
            score_df["yoynetprofit_deducted_score"] = 0

            score_df["qfa_yoysales_score"] = 0
            score_df["qfa_yoyprofit_score"] = 0
            score_df["qfa_grossprofitmargin_score"] = 0
            score_df["qfa_netprofitmargin_score"] = 0

            score_df["total_score"] = 0

            # 估值指标
            score_df["pe_score"] = np.where(score_df["PE_TTM"] < 40, np.where(score_df["PE_TTM"] < 20, [2], [1]), [0]) # 市盈率PE（TTM）
            score_df["pb_score"] = np.where(score_df["PB"] < 6, np.where(score_df["PB"] < 1, [2], [1]), [0]) # 市净率PB
            # 盈利能力指标
            score_df["roa_score"] = np.where(score_df["ROA"] > 5, np.where(score_df["ROA"] > 10, [2], [1]), [0]) # 总资产收益率
            score_df["roe_score"] = np.where(score_df["ROE"] > 10, np.where(score_df["ROE"] > 20, [2], [1]), [0]) # 净资产收益率
            # 资产负债指标
            score_df["debttoassets_score"] = np.where(score_df["DEBTTOASSETS"] < 50, np.where(score_df["DEBTTOASSETS"] < 25, [2], [1]), [0]) # 资产负债率
            score_df["current_score"] = np.where(score_df["CURRENT"] > 2,np.where(score_df["CURRENT"] > 5, [2], [1]), [0]) # 流动比率
            # 现金流指标
            score_df["ocf_score"] = np.where((score_df["OCFTOSALES"] > 0) & (score_df["TOT_OPER_REV"] > 0), [1], [0]) # 经营性现金净流量>0
            score_df["cashflow_score"] = np.where(score_df["CASHFLOW_TTM"] > 0, [1], [0]) # 现金净流量（TTM）
            score_df["ocftoprofit_score"] = np.where(score_df["OCFTOPROFIT"] > 1, [2], [0]) # 经营性现金流/净利润
            # 成长性指标
            score_df["yoy_or_score"] = np.where(score_df["YOY_OR"] > 15, np.where(score_df["YOY_OR"] > 30, [2], [1]), [0]) # 营业收入同比增长率
            score_df["yoynetprofit_deducted_score"] = np.where(score_df["YOYNETPROFIT_DEDUCTED"] > 15, np.where(score_df["YOYNETPROFIT_DEDUCTED"] > 30, [2], [1]), [0]) # 扣非后净利润同比增长率
            # 环比指标
            score_df["qfa_yoysales_score"] = np.where(score_df["QFA_YOYSALES"] > score_df["LAST_QFA_YOYSALES"], [1], [0]) # 单季度营收同比值加速
            score_df["qfa_yoyprofit_score"] = np.where(score_df["QFA_YOYPROFIT"] > score_df["LAST_QFA_YOYPROFIT"], [1], [0]) # 单季度净利润同比值加速
            score_df["qfa_grossprofitmargin_score"] = np.where(score_df["QFA_GROSSPROFITMARGIN"] > score_df["LAST_QFA_GROSSPROFITMARGIN"], [1], [0]) # 毛利率环比提升
            score_df["qfa_netprofitmargin_score"] = np.where(score_df["QFA_NETPROFITMARGIN"] > score_df["LAST_QFA_NETPROFITMARGIN"], [1], [0]) # 净利率环比提升

            score_df["total_score"] = score_df["pe_score"] + score_df["pb_score"] + score_df["roa_score"]+score_df["roe_score"] +\
                                      score_df["debttoassets_score"] + score_df["current_score"] + score_df["ocf_score"] + \
                                      score_df["cashflow_score"]+score_df["ocftoprofit_score"]+score_df["yoy_or_score"]+score_df["yoynetprofit_deducted_score"]+ \
                                      score_df["qfa_yoysales_score"]+score_df["qfa_yoyprofit_score"]+score_df["qfa_grossprofitmargin_score"]+score_df["qfa_netprofitmargin_score"]

            score_df.sort_values(by="peg", ascending=True, inplace=True)
            # xlsx_output = pd.ExcelWriter("full_date.xls")
            # score_df.to_excel(xlsx_output)
            # xlsx_output.save()

            score_df = score_df[(score_df["peg"]>0)&(score_df["peg"]<2) & (score_df["total_score"] >= 10)]


            name_map = [("SEC_NAME",u"股票简称"),("INDUSTRY_SW",u"申万一级行业"),("peg",u"计算PEG"),
                        ("pe_est_last_"+str_after_two_year,u"预测PE_"+str_after_two_year),
                        ("pe_est_last_"+str_after_one_year,u"预测PE_"+str_after_one_year),
                        ("pe_est_last_"+self.date["year_str"],u"预测PE_"+self.date["year_str"]),
                        ("est_eps_" + str_after_two_year, u"预测EPS_" + str_after_two_year),
                        ("est_eps_" + str_after_one_year, u"预测EPS_" + str_after_one_year),
                        ("est_eps_" + self.date["year_str"], u"预测EPS_" + self.date["year_str"]),
                        ("eps_acc",u"CAGR"),("total_score",u"总分"),("PB",u"PB"),
                        ("PE_TTM",u"PE（TTM）"),("YOY_OR",u"营收同比增长率"),("LAST_YEAR_YOY_OR",u"上年度营收同比增长率"),
                        ("YOYOP",u"营业利润同比增长率"),("ROA",u"ROA"),("ROE",u"ROE"),("DEBTTOASSETS",u"资产负债率"),
                        ("YOYNETPROFIT_DEDUCTED",u"扣非后净利润同比增长率"),("QFA_YOYSALES",u"单季度营收同比增长率"),
                        ("LAST_QFA_YOYSALES",u"上季度营收同比增长率"),("QFA_YOYPROFIT",u"单季度净利润同比增长率"),
                        ("LAST_QFA_YOYPROFIT",u"上季度净利润同比增长率"),("QFA_GROSSPROFITMARGIN",u"单季度毛利率"),
                        ("LAST_QFA_GROSSPROFITMARGIN",u"上季度毛利率"),("QFA_NETPROFITMARGIN",u"单季度净利率"),
                        ("LAST_QFA_NETPROFITMARGIN",u"上季度净利率"),("OCFTOPROFIT",u"经营性现金净流量/净利润")]
            field_map = OrderedDict(name_map)
            rename_map = dict(name_map)

            clean_df = score_df[field_map.keys()]
            clean_df.rename(columns=rename_map, inplace=True)

            self.save_df_to_excel(clean_df)

        elif self.flag == "H":

            gross_pick_result = w.wss(peg_df.index.tolist(),
                                      "eps_ttm,pb,yoy_or,yoyop",
                                      "rptDate=" + self.date["rpt_date_str"] + ";rptType=1;tradeDate=" +
                                      text_trade_date + ";ruleType=9")
            gross_pick_df = DataFrame(gross_pick_result.Data, columns=gross_pick_result.Codes,
                                      index=gross_pick_result.Fields).T

            gross_pick_df = gross_pick_df[
                (gross_pick_df["PB"] > 0) & (gross_pick_df["PB"] < 12)]
            gross_pick_df = gross_pick_df[(gross_pick_df["YOY_OR"] > 0)]
            gross_pick_df = gross_pick_df[(gross_pick_df["YOYOP"] > 0)]

            # print "gross_pick_df", gross_pick_df
            print "gross_pick_df.index.tolist()", gross_pick_df.index.tolist()
            well_pick_result = w.wss(gross_pick_df.index.tolist(),
                                     "pe_ttm,sec_name,roa,roe,debttoassets,current,profittogr,industry_hs,"
                                     "tot_oper_rev,cashflow_ttm,yoynetprofit_deducted",
                                     "rptDate=" + self.date[
                                         "rpt_date_str"] + ";unit=1;rptType=1;category=1;tradeDate=" + text_trade_date)
            # print "well_pick_result", well_pick_result
            well_pick_df = DataFrame(well_pick_result.Data, columns=well_pick_result.Codes,
                                     index=well_pick_result.Fields).T
            print "well_pick_df", well_pick_df


            # 获取上年年报的营收同比
            last_year_report_result = w.wss(gross_pick_df.index.tolist(),
                                            "yoy_or", "rptDate=" + self.date["last_year_last_rpt_date_str"])
            last_year_report_df = DataFrame(last_year_report_result.Data, columns=last_year_report_result.Codes,
                                                index=last_year_report_result.Fields).T
            last_year_report_df.rename(columns={"YOY_OR": "LAST_YEAR_YOY_OR"}, inplace=True)

            score_df = pd.concat([peg_df, gross_pick_df, well_pick_df, last_year_report_df], axis=1)
            score_df = score_df.dropna(axis=0, how='any')


            score_df["pe_score"] = 0
            score_df["pb_score"] = 0

            score_df["roa_score"] = 0
            score_df["roe_score"] = 0

            score_df["debttoassets_score"] = 0
            score_df["current_score"] = 0

            score_df["cashflow_score"] = 0

            score_df["yoy_or_score"] = 0
            score_df["yoynetprofit_deducted_score"] = 0


            score_df["total_score"] = 0

            # 估值指标
            score_df["pe_score"] = np.where(score_df["PE_TTM"] < 40, np.where(score_df["PE_TTM"] < 20, [2], [1]),
                                            [0])  # 市盈率PE（TTM）
            score_df["pb_score"] = np.where(score_df["PB"] < 6, np.where(score_df["PB"] < 1, [2], [1]),
                                            [0])  # 市净率PB
            # 盈利能力指标
            score_df["roa_score"] = np.where(score_df["ROA"] > 5, np.where(score_df["ROA"] > 10, [2], [1]),
                                             [0])  # 总资产收益率
            score_df["roe_score"] = np.where(score_df["ROE"] > 10, np.where(score_df["ROE"] > 20, [2], [1]),
                                             [0])  # 净资产收益率
            # 资产负债指标
            score_df["debttoassets_score"] = np.where(score_df["DEBTTOASSETS"] < 50,
                                                      np.where(score_df["DEBTTOASSETS"] < 25, [2], [1]),
                                                      [0])  # 资产负债率
            score_df["current_score"] = np.where(score_df["CURRENT"] > 2,
                                                 np.where(score_df["CURRENT"] > 5, [2], [1]), [0])  # 流动比率
            # 现金流指标
            # score_df["ocf_score"] = np.where((score_df["OCFTOSALES"] > 0) & (score_df["TOT_OPER_REV"] > 0), [1],
            #                                  [0])  # 经营性现金净流量>0
            score_df["cashflow_score"] = np.where(score_df["CASHFLOW_TTM"] > 0, [1], [0])  # 现金净流量（TTM）
            # score_df["ocftoprofit_score"] = np.where(score_df["OCFTOPROFIT"] > 1, [2], [0])  # 经营性现金流/净利润
            # 成长性指标
            score_df["yoy_or_score"] = np.where(score_df["YOY_OR"] > 15,
                                                np.where(score_df["YOY_OR"] > 30, [2], [1]), [0])  # 营业收入同比增长率
            score_df["yoynetprofit_deducted_score"] = np.where(score_df["YOYNETPROFIT_DEDUCTED"] > 15,
                                                               np.where(score_df["YOYNETPROFIT_DEDUCTED"] > 30, [2],
                                                                        [1]), [0])  # 扣非后净利润同比增长率


            score_df["total_score"] = score_df["pe_score"] + score_df["pb_score"] + score_df["roa_score"] + \
                                      score_df["roe_score"] + \
                                      score_df["debttoassets_score"] + score_df["current_score"] + \
                                      score_df["cashflow_score"] + score_df[
                                          "yoy_or_score"] + score_df["yoynetprofit_deducted_score"]

            score_df.sort_values(by="peg", ascending=True, inplace=True)
            xlsx_output = pd.ExcelWriter("full_date.xls")
            score_df.to_excel(xlsx_output)
            xlsx_output.save()

            score_df = score_df[(score_df["total_score"] >= 6)]

            name_map = [("SEC_NAME", u"股票简称"), ("INDUSTRY_HS", u"恒生一级行业"),("peg",u"计算PEG"),
                        ("pe_est_last_"+str_after_two_year,u"预测PE_"+str_after_two_year),
                        ("pe_est_last_"+str_after_one_year,u"预测PE_"+str_after_one_year),
                        ("pe_est_last_"+self.date["year_str"],u"预测PE_"+self.date["year_str"]),
                        ("est_eps_" + str_after_two_year, u"预测EPS_" + str_after_two_year),
                        ("est_eps_" + str_after_one_year, u"预测EPS_" + str_after_one_year),
                        ("est_eps_" + self.date["year_str"], u"预测EPS_" + self.date["year_str"]),
                        ("eps_acc",u"eps增速"),("total_score",u"总分"),("PB",u"PB"),
                        ("PE_TTM", u"PE（TTM）"), ("PB", u"PB"), ("YOY_OR", u"营收同比增长率"),
                        ("LAST_YEAR_YOY_OR", u"上年度营收同比增长率"),
                        ("YOYOP", u"营业利润同比增长率"), ("ROA", u"ROA"), ("ROE", u"ROE"), ("DEBTTOASSETS", u"资产负债率"),
                        ("YOYNETPROFIT_DEDUCTED", u"扣非后净利润同比增长率")]
            field_map = OrderedDict(name_map)
            rename_map = dict(name_map)

            clean_df = score_df[field_map.keys()]
            clean_df.rename(columns=rename_map, inplace=True)

            self.save_df_to_excel(clean_df)

    def get_small_cap_undervalued_stocks(self, stocks):
        # stocks = stocks[0:100]
        # print stocks

        text_trade_date = (self.date["trade_date"] - dt.timedelta(days=1)).strftime("%Y-%m-%d")
        str_after_two_year = str(int(self.date["year_str"])+2)
        str_after_one_year = str(int(self.date["year_str"]) + 1)


        print "unit=1;ruleType=10;tradeDate="+text_trade_date+";rptDate=" + self.date["rpt_date_str"]
        cap_query_result = w.wss(stocks, "mkt_cap_ard,pe_ttm,yoy_or,yoyprofit,pb",
                                 "unit=1;ruleType=10;tradeDate="+text_trade_date+";rptDate=" + self.date["rpt_date_str"])
        cap_df = DataFrame(cap_query_result.Data, columns=cap_query_result.Codes, index=cap_query_result.Fields).T
        cap_df = cap_df[(cap_df["MKT_CAP_ARD"]<5*10e9)& (cap_df["PE_TTM"]<30)& (cap_df["PB"]<4)&
                        (cap_df["PE_TTM"]>0)& (cap_df["YOY_OR"]>0)& (cap_df["YOYPROFIT"]>0)]

        # self.get_peg_pick_stocks(cap_df.index.tolist())
        return cap_df.index.tolist()

    def get_MA60_pic(self,all_stocks):
        # all_stocks = all_stocks[0:10]
        # import matplotlib.pyplot as plt
        # # 中文乱码的处理
        # plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        # plt.rcParams['axes.unicode_minus'] = False

        text_trade_date = (self.date["trade_date"]).strftime("%Y-%m-%d")

        request_result = w.wss(
            all_stocks,
            "MA,industry_citic,close", "tradeDate=" + text_trade_date + ";MA_N=60;priceAdj=F;cycle=D;industryType=1")

        result_df = DataFrame(request_result.Data, columns=request_result.Codes, index=request_result.Fields).T
        # result_df.replace(None, np.nan)
        # result_df.to_csv("raw_request.csv", index=True, encoding="gb2312")
        result_df.dropna(axis=1,how="any")
        # result_df.to_csv("request.csv",index=True,encoding="gb2312")
        result_df["MA-CLOSE"] = result_df["MA"] < result_df["CLOSE"]
        result_df.replace(True, 1)
        result_df.replace(False, 0)

        pivot = result_df.pivot_table(index = ["INDUSTRY_CITIC"],values=["CLOSE","MA-CLOSE"],aggfunc=[np.count_nonzero])
        # pivot.to_csv("pivot.csv",index=True,encoding="gb2312")
        # result = pivot["CLOSE"]/pivot["MA-CLOSE"]
        result = pivot.reset_index()
        rate = result["count_nonzero"]["MA-CLOSE"]/result["count_nonzero"]["CLOSE"]
        name = result["INDUSTRY_CITIC"]

        final_df = pd.DataFrame({u"行业":name, u"比例":rate})
        # final_df.sort_values(by=u"行业", ascending=True, inplace=True)
        final_df.to_csv("MA60.csv", index=False, encoding="gb2312")

        self.save_df_to_excel(final_df)

        # print "final_df",final_df
        # # 绘图
        # plt.bar(range(len(rate)),rate, align='center', color='steelblue', alpha=0.8)
        # # 添加轴标签
        # plt.ylabel('Rate')
        # # 添加标题
        # plt.title('MA60强势股占比')
        # # 添加刻度标签
        # plt.xticks(range(len(name)),name)
        #
        # mngr = plt.get_current_fig_manager()
        # mngr.window.setGeometry(50, 50, 960, 640)
        # plt.tight_layout()
        #
        #
        # # 设置Y轴的刻度范围
        # plt.ylim([0, 1.2])
        # plt.show()



#############################################################

class AShareData(DataFormatter):
    def __init__(self, flag, industry, group):
        super(AShareData, self).__init__(flag, industry, group)


class HShareData(DataFormatter):
    def __init__(self, flag, industry, group):
        super(HShareData, self).__init__(flag, industry, group)


def thrust_up_main():
    # # 创历史新高
    A_stocks_data = AShareData("A", "INDUSTRY_SW", "history_high")
    A_stocks_data.main()
    H_stocks_data = HShareData("H", "INDUSTRY_HS", "history_high")
    H_stocks_data.main()
    #
    # # 创阶段新高
    A_stocks_data = AShareData("A", "INDUSTRY_SW", "stage_high")
    A_stocks_data.main()
    H_stocks_data = HShareData("H", "INDUSTRY_HS", "stage_high")
    H_stocks_data.main()

    ## 创阶段新低
    # A_stocks_data = AShareData("A", "INDUSTRY_SW", "stage_low")
    # A_stocks_data.main()
    # H_stocks_data = HShareData("H", "INDUSTRY_HS", "stage_low")
    # H_stocks_data.main()

    # 环比增长
    # A_stocks_data = AShareData("A", "INDUSTRY_SW", "quarter_increase")
    # A_stocks_data.main()
    # H_stocks_data = HShareData("H", "INDUSTRY_HS", "quarter_increase")
    # H_stocks_data.main()

    # # 股东增持
    # # 公司行动事件汇总——全部A股、本周——公司资料变更
    # # 将股东增持文件命名为ih_A.xlsx
    # A_stocks_data = AShareData("A", "INDUSTRY_SW", "increase_holding")
    # A_stocks_data.main()
    # H_stocks_data = HShareData("H", "INDUSTRY_HS", "increase_holding")
    # H_stocks_data.main()

    # 质押比例
    # A_stocks_data = AShareData("A", "INDUSTRY_SW", "share_pledged")
    # A_stocks_data.main()

    # 平台突破
    A_stocks_data = AShareData("A", "INDUSTRY_SW", "thrust_up_plate")
    A_stocks_data.main()
    H_stocks_data = HShareData("H", "INDUSTRY_HS", "thrust_up_plate")
    H_stocks_data.main()

    # peg选股
    A_stocks_data = AShareData("A", "INDUSTRY_SW", "peg_pick")
    A_stocks_data.main()
    H_stocks_data = AShareData("H", "INDUSTRY_HS", "peg_pick")
    H_stocks_data.main()

    # A股小市值低估值统计
    A_stocks_data = AShareData("A", "INDUSTRY_SW", "small_cap_undervalue")
    A_stocks_data.main()

    # 中信行业MA60统计
    A_stocks_data = AShareData("A", "INDUSTRY_CITIC", "MA60")
    A_stocks_data.main()

if __name__ == "__main__":
    thrust_up_main()


