
# coding:utf-8

import numpy as np
import pandas as pd
from pandas import DataFrame
import os
import datetime as dt
import xlwt
import copy
import json
from collections import OrderedDict

today = dt.datetime.today()
today_str = today.strftime("%Y-%m-%d")
# str_after_two_year = str(int(today_str[0:4]) + 2)
# str_after_one_year = str(int(today_str[0:4]) + 1)
date_mark = "data_output"+ "/" + today_str + "/"


class Concater(object):
    def __init__(self,flag,group):

        self.assist_df = DataFrame()
        self.main_df = DataFrame()
        self.flag = flag
        self.group = group
        self.remarks = ""
        self.book = None
        self.date_mark = "data_output"+ "/" + today_str + "/"

        if "peg_pick" in self.group:
            self.main_file = "peg_pick" + "_" + self.flag + today_str + ".xls"
        else:
            self.main_file = self.group + "_" + self.flag + today_str + ".xls"
        self.xls_cleanData_filename = "cl_" + self.group + "_" + self.flag + today_str + ".xls"

        self.get_file_name()
        self.get_df_from_xls()
        self.create_book()

    def get_file_name(self):
        if self.group == "increase_holding":
            self.assist_file = "ih_" + self.flag + ".xlsx"
        elif self.group == "history_high":
            self.assist_file = self.group + "_" + self.flag + "_count" + today_str + ".xls"
        elif self.group == "peg_pick":
            self.remarks = ""
            self.assist_file = None
        elif self.group == "peg_pick_normal_speed":
            self.remarks = u"备注：" \
                           u"1.本表数据筛选自表《peg精选-A股》，筛选标准为(PE<25,计算PEG<1,CAGR>20)；" \
                           u"2.表中财务数据均为当前报告期数据(预测数据除外)，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.assist_file = None
        elif self.group == "peg_pick_high_speed":
            self.remarks = u"备注：" \
                           u"1.本表数据筛选自表《peg精选-A股》，筛选标准为(计算PEG<0.85,CAGR>40)；" \
                           u"2.表中财务数据均为当前报告期数据(预测数据除外)，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.assist_file = None
        elif self.group == "small_cap_undervalue":
            self.remarks = u"备注：" \
                           u"1.本表数据筛选自表《peg精选-A股》，筛选标准为(市值<10亿,PE(TTM)<30)；" \
                           u"2.表中财务数据均为当前报告期数据(预测数据除外)，部分公司因尚未公布当前报告期数据，因此为空值；" \
                           u"3.净资产收益率、总资产净利率、资产负债率、流动比率等四指标数据在当前报告期未取到，则取其上一报告期数据。"
            self.assist_file = None


    def get_df_from_xls(self):
        if (not self.assist_file) or (not os.path.exists(self.date_mark+self.assist_file)):
            print "未找到文件", self.date_mark,self.assist_file
        else:
           self.assist_df = pd.read_excel(self.date_mark+self.assist_file,header=0)
        if not os.path.exists(self.date_mark+self.main_file):
            print "未找到文件", self.date_mark,self.main_file
        else:
            self.main_df = pd.read_excel(self.date_mark+self.main_file,header=0)
            # print self.main_df

    def set_style(self, name, height, bold=False, pattern_switch=False, alignment_center=True):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name  # 'Times New Roman'
        font.bold = bold
        font.color_index = 2
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
        # alignment.alignment = alignment_switch
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

    def create_book(self):
        if not self.book:
            self.book = xlwt.Workbook(encoding="utf-8")

        if self.group == "increase_holding":
            self.group_in_chinese = "增持"
        elif self.group == "history_high":
            self.group_in_chinese = "历史新高"
        elif self.group == "peg_pick":
            self.group_in_chinese = "peg精选"
        elif self.group == "peg_pick_normal_speed":
            self.group_in_chinese = "低PE有成长"
        elif self.group == "peg_pick_high_speed":
            self.group_in_chinese = "估值合理高成长"
        elif self.group == "small_cap_undervalue":
            self.group_in_chinese = "低估小市值"
        self.sheet1 = self.book.add_sheet(self.group_in_chinese + "-" +self.flag + "股")

    def concat_history_high_df(self):
        # 合并创新高股票数据
        if self.assist_df.empty or self.main_df.empty:
            print "待合并创新高股票df为空"
            return

        self.main_df[u"连续新高周数"] = 0
        main_code = self.main_df.index.tolist()
        print "main_code",main_code
        for stock in main_code[0:-1]:
            # print stock
            # print "assist",self.assist_df["count"][stock]
            # print "main",self.main_df[u"连续新高周数"][stock]
            self.main_df[u"连续新高周数"][stock] = self.assist_df["count"][stock]

        self.remarks = main_code.pop(-1)
        print self.remarks
        # print "main_code.pop(-1)",main_code.pop(-1)
        self.main_df.drop([self.remarks],inplace=True)

        # print self.main_df
        # print self.remarks
        add_column = self.main_df[u"连续新高周数"]
        self.main_df.drop(labels=[u"连续新高周数"], axis=1, inplace=True)
        self.main_df.insert(2, u"连续新高周数", add_column)

        # print self.main_df
        # xlsx_output = pd.ExcelWriter("1result.xls")
        # self.main_df.to_excel(xlsx_output, float_format = '%.2f')
        # xlsx_output.save()
        # self.main_df = self.main_df.replace(np.nan, "")
        return self.main_df

    def concat_ih_df(self):
        # 合并股东增持股票数据
        if self.assist_df.empty or self.main_df.empty:
            print "待合并股东增持股票数据df为空"
            return

        del self.assist_df[u"序号"]
        del self.assist_df[u"证券类型"]
        del self.assist_df[u"事件大类"]
        del self.assist_df[u"披露日期"]
        self.assist_df = self.assist_df.dropna(axis=0, how='any')

        assist_code = list(set(self.assist_df[u"证券代码"].tolist()))
        main_code = self.main_df.index.tolist()
        self.remarks = main_code.pop(-1)
        print self.remarks
        self.main_df.drop([self.remarks],inplace=True)

        del_codes = list(set(assist_code)^set(main_code))
        for code in del_codes:
            code_index = self.assist_df[self.assist_df[u"证券代码"] == code].index.tolist()
            self.assist_df.drop(code_index,inplace=True)

        self.assist_df.reset_index(drop = True, inplace = True)


        if self.flag == "H":
            self.assist_df[u"恒生行业代码（三级行业）"] = ""
        else:
            self.assist_df[u"申万行业代码（三级行业）"] = ""
            self.assist_df[u"雪球一周关注增长率%"] = ""
            self.assist_df[u" 雪球累计关注人数"] = ""
            self.assist_df[u"市盈率PE（TTM,扣非）"] = ""
            self.assist_df[u"预测PEG（未来12个月）"] = ""

        self.assist_df[u"收盘价"] = ""
        self.assist_df[u"市盈率PE（TTM）"] = ""
        self.assist_df[u"市净率PB"] = ""
        self.assist_df[u"预测PE"] = ""
        self.assist_df[u"净资产收益率"] = ""
        self.assist_df[u"总资产净利率"] = ""
        self.assist_df[u"资产负债率"] = ""
        self.assist_df[u"流动比率"] = ""
        self.assist_df[u"营业收入同比增长率"] = ""
        self.assist_df[u"净利润（同比增长率）"] = ""
        self.assist_df[u"单季度营业收入环比增长率"] = ""
        self.assist_df[u"单季度净利润环比增长率"] = ""
        self.assist_df[u"业绩预告日期"] = ""
        self.assist_df[u"业绩预告类型"] = ""
        self.assist_df[u"业绩预告摘要"] = ""

        # print "self.assist_df",self.assist_df

        for index in self.assist_df.index.tolist():
            stock = self.assist_df[u"证券代码"][index]
            print stock
            if self.flag == "H":
                self.assist_df[u"恒生行业代码（三级行业）"][index] = self.main_df[u"恒生行业代码（三级行业）"][stock]
            else:
                self.assist_df[u"雪球一周关注增长率%"][index] = self.main_df[u"雪球一周关注增长率%"][stock]
                self.assist_df[u" 雪球累计关注人数"][index] = self.main_df[u"雪球累计关注人数"][stock]
                self.assist_df[u"申万行业代码（三级行业）"][index] = self.main_df[u"申万行业代码（三级行业）"][stock]
                self.assist_df[u"市盈率PE（TTM,扣非）"][index] = self.main_df[u"市盈率PE（TTM,扣非）"][stock]
                self.assist_df[u"预测PEG（未来12个月）"][index] = self.main_df[u"预测PEG（未来12个月）"][stock]


            self.assist_df[u"收盘价"][index] = self.main_df[u"收盘价"][stock]
            self.assist_df[u"市盈率PE（TTM）"][index] = self.main_df[u"市盈率PE（TTM）"][stock]
            self.assist_df[u"市净率PB"][index] = self.main_df[u"市净率PB"][stock]
            self.assist_df[u"预测PE"][index] = self.main_df[u"预测PE"][stock]
            self.assist_df[u"净资产收益率"][index] = self.main_df[u"净资产收益率"][stock]
            self.assist_df[u"总资产净利率"][index] = self.main_df[u"总资产净利率"][stock]
            self.assist_df[u"资产负债率"][index] = self.main_df[u"资产负债率"][stock]
            self.assist_df[u"流动比率"][index] = self.main_df[u"流动比率"][stock]
            self.assist_df[u"营业收入同比增长率"][index] = self.main_df[u"营业收入同比增长率"][stock]
            self.assist_df[u"净利润（同比增长率）"][index] = self.main_df[u"净利润（同比增长率）"][stock]
            self.assist_df[u"单季度营业收入环比增长率"][index] = self.main_df[u"单季度营业收入环比增长率"][stock]
            self.assist_df[u"单季度净利润环比增长率"][index] = self.main_df[u"单季度净利润环比增长率"][stock]
            self.assist_df[u"业绩预告日期"][index] = self.main_df[u"业绩预告日期"][stock]
            self.assist_df[u"业绩预告类型"][index] = self.main_df[u"业绩预告类型"][stock]
            self.assist_df[u"业绩预告摘要"][index] = self.main_df[u"业绩预告摘要"][stock]
            self.assist_df = self.assist_df.replace(np.nan,"")

        print "self.assist_df",self.assist_df.head(5)
        return self.assist_df

    def write_ih_book(self, data):
        if data.empty:
            print "数据为空"
            return
        print self.remarks
        self.sheet1.write(0, 0, "", self.set_style('Times New Roman', 220, True))
        row0 = data.columns.tolist()
        column0 = data.index.tolist()
        columns_count = len(row0)
        rows_count = len(column0)

        # xlsx_output = pd.ExcelWriter("h_result.xls")
        # data.to_excel(xlsx_output, float_format = '%.2f')
        # xlsx_output.save()

        if self.flag == "H":
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)
            for i in range(0, columns_count):
                for j in range(1, rows_count + 1):
                    cell_data = data[row0[i]][j - 1]
                    if i == 18:
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
                    if i == 20 or i == 4:
                        style = self.set_style('Arial', 200, False, False, False)
                        self.sheet1.write(j, i + 1, cell_data, style)
                        continue

                    cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    self.sheet1.write(j, i + 1, cell_data, content_style)  # 第一列
            self.sheet1.col(0).width = 90 * 25
            self.sheet1.col(1).width = 160 * 25
            self.sheet1.col(2).width = 150 * 25
            self.sheet1.col(3).width = 250 * 25
            self.sheet1.col(4).width = 120 * 25
            self.sheet1.col(5).width = 550 * 25
            self.sheet1.col(6).width = 360 * 25
            self.sheet1.col(7).width = 90 * 25

            self.sheet1.col(8).width = 250 * 25
            self.sheet1.col(9).width = 150 * 25
            self.sheet1.col(10).width = 100 * 25
            self.sheet1.col(11).width = 180 * 25
            self.sheet1.col(12).width = 180 * 25
            self.sheet1.col(13).width = 150 * 25
            self.sheet1.col(14).width = 120 * 25
            self.sheet1.col(15).width = 300 * 25
            self.sheet1.col(16).width = 300 * 25
            self.sheet1.col(17).width = 360 * 25
            self.sheet1.col(18).width = 350 * 25
            self.sheet1.col(19).width = 200 * 25
            self.sheet1.col(20).width = 200 * 25
            self.sheet1.col(21).width = 2500 * 25
        else:
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)

            dateFormat = copy.deepcopy(content_style)
            dateFormat.num_format_str = 'yyyy-mm-dd'
            column3_style = self.set_style('Arial', 200, False, False, False)
            for i in range(0, columns_count):
                for j in range(1, rows_count + 1):
                    # print i,j,data[row0[i]][j - 1]
                    if i == 4:
                        cell_data = data[row0[i]][j - 1]

                        self.sheet1.write(j, i + 1, cell_data, column3_style)
                        continue
                    if i == 22:
                        cell_data = data[row0[i]][j - 1]
                        if cell_data:
                            if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                                self.sheet1.write(j, i + 1, cell_data, dateFormat)
                            else:
                                self.sheet1.write(j, i + 1, "", content_style)
                        else:
                            self.sheet1.write(j, i + 1, "", dateFormat)
                        continue

                    cell_data = data[row0[i]][j - 1]
                    # cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    self.sheet1.write(j, i + 1, cell_data, content_style)

            self.sheet1.col(0).width = 90 * 25
            self.sheet1.col(1).width = 160 * 25
            self.sheet1.col(2).width = 120 * 25
            self.sheet1.col(3).width = 290 * 25
            self.sheet1.col(4).width = 120 * 25
            self.sheet1.col(5).width = 800 * 25

            self.sheet1.col(6).width = 400 * 25
            self.sheet1.col(7).width = 290 * 25
            self.sheet1.col(8).width = 270 * 25
            self.sheet1.col(9).width = 300 * 25
            self.sheet1.col(10).width = 320 * 25
            self.sheet1.col(11).width = 90 * 25
            self.sheet1.col(12).width = 250 * 25
            self.sheet1.col(13).width = 150 * 25
            self.sheet1.col(14).width = 100 * 25
            self.sheet1.col(15).width = 180 * 25
            self.sheet1.col(16).width = 180 * 25
            self.sheet1.col(17).width = 150 * 25
            self.sheet1.col(18).width = 120 * 25
            self.sheet1.col(19).width = 320 * 25
            self.sheet1.col(20).width = 300 * 25
            self.sheet1.col(21).width = 360 * 25
            self.sheet1.col(22).width = 350 * 25
            self.sheet1.col(23).width = 200 * 25
            self.sheet1.col(24).width = 200 * 25
            self.sheet1.col(25).width = 700 * 25

        print self.remarks
        self.sheet1.write(rows_count + 1, 0, self.remarks)
        tall_style = xlwt.easyxf('font:height 360;')
        first_row = self.sheet1.row(0)
        first_row.set_style(tall_style)

    def write_history_high_book(self, data):
        if data.empty:
            print "数据为空"
            return
        data.replace(np.nan, "")
        self.sheet1.write(0, 0, "", self.set_style('Times New Roman', 220, True))
        row0 = data.columns.tolist()
        column0 = data.index.tolist()
        columns_count = len(row0)
        rows_count = len(column0)


        if self.flag == "H":
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)

            for i in range(0, columns_count):
                for j in range(1, rows_count + 1):
                    cell_data = data[row0[i]][j - 1]
                    cell_data = "" if cell_data is np.nan else cell_data
                    print i,j,cell_data,type(cell_data)
                    if i == 2:  # 连续新高周数
                        self.sheet1.write(j, i + 1, int(cell_data), content_style)
                        continue
                    if i == 15:
                        if cell_data:
                            # print cell_data
                            # if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                            #     dateFormat = copy.deepcopy(content_style)
                            #     dateFormat.num_format_str = 'yyy-mm-dd'
                            #     self.sheet1.write(j, i + 1, cell_data, dateFormat)
                            # else:
                                self.sheet1.write(j, i + 1, "", content_style)
                        else:
                            self.sheet1.write(j, i + 1, "", content_style)
                        continue
                    if i == 21 or i == 5:
                        style = self.set_style('Arial', 200, False, False, False)
                        self.sheet1.write(j, i + 1, cell_data, style)
                        continue

                    cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    self.sheet1.write(j, i + 1, cell_data, content_style)  # 第一列
            self.sheet1.col(0).width = 90 * 25
            self.sheet1.col(1).width = 160 * 25
            self.sheet1.col(2).width = 150 * 25
            self.sheet1.col(3).width = 180 * 25
            self.sheet1.col(4).width = 360 * 25
            self.sheet1.col(5).width = 250 * 25
            self.sheet1.col(6).width = 150 * 25
            self.sheet1.col(7).width = 100 * 25
            self.sheet1.col(8).width = 180 * 25
            self.sheet1.col(9).width = 180 * 25
            self.sheet1.col(10).width = 150 * 25
            self.sheet1.col(11).width = 120 * 25
            self.sheet1.col(12).width = 300 * 25
            self.sheet1.col(13).width = 300 * 25
            self.sheet1.col(14).width = 360 * 25
            self.sheet1.col(15).width = 350 * 25
            self.sheet1.col(16).width = 200 * 25
            self.sheet1.col(17).width = 200 * 25
            self.sheet1.col(18).width = 2500 * 25
        else:
            for i in range(1, columns_count + 1):
                self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
            for i in range(0, rows_count):
                self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
            content_style = self.set_style('Arial', 200, False, False)

            for i in range(0, columns_count):
                for j in range(1, rows_count + 1):
                    cell_data = data[row0[i]][j - 1]
                    if i == 2:  # 连续新高周数
                        self.sheet1.write(j, i + 1, int(cell_data), content_style)
                        continue
                    if i == 5:  # 行业代码
                        column3_style = self.set_style('Arial', 200, False, False, False)
                        self.sheet1.write(j, i + 1, cell_data, column3_style)
                        continue
                    if i == 19:  # 业绩预告日期
                        # print i,j,cell_data
                        if cell_data:
                            if cell_data > dt.datetime.strptime("1900-01-01", '%Y-%m-%d'):
                                dateFormat = copy.deepcopy(content_style)
                                dateFormat.num_format_str = 'yyyy-mm-dd'
                                self.sheet1.write(j, i + 1, cell_data, dateFormat)
                            else:
                                self.sheet1.write(j, i + 1, "", content_style)
                        else:
                            self.sheet1.write(j, i + 1, "-", content_style)
                        continue

                    cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                    # print cell_data,type(cell_data)
                    self.sheet1.write(j, i + 1, cell_data, content_style)

            self.sheet1.col(0).width = 160 * 25
            self.sheet1.col(1).width = 120 * 25 # 股票简称  4个字宽120
            self.sheet1.col(2).width = 90 * 25
            self.sheet1.col(3).width = 180 * 25
            self.sheet1.col(4).width = 290 * 25
            self.sheet1.col(5).width = 250 * 25
            self.sheet1.col(6).width = 360 * 25
            self.sheet1.col(7).width = 250 * 25
            self.sheet1.col(8).width = 320 * 25
            self.sheet1.col(9).width = 150 * 25
            self.sheet1.col(10).width = 320 * 25
            self.sheet1.col(11).width = 100 * 25
            self.sheet1.col(12).width = 180 * 25
            self.sheet1.col(13).width = 180 * 25
            self.sheet1.col(14).width = 150 * 25
            self.sheet1.col(15).width = 120 * 25
            self.sheet1.col(16).width = 320 * 25
            self.sheet1.col(17).width = 300 * 25
            self.sheet1.col(18).width = 360 * 25
            self.sheet1.col(19).width = 350 * 25
            self.sheet1.col(20).width = 200 * 25
            self.sheet1.col(21).width = 200 * 25
            self.sheet1.col(22).width = 700 * 25

        self.sheet1.write(rows_count + 1,0, self.remarks)
        tall_style = xlwt.easyxf('font:height 360;')
        first_row = self.sheet1.row(0)
        first_row.set_style(tall_style)

    def write_peg_pick_book(self):
        if self.main_df.empty:
            print "数据为空"
            return

        print self.flag,self.group
        if self.flag == "A" and self.group == "peg_pick_normal_speed":
            # 筛选出低PE有成长的股票
            self.main_df = self.main_df[(self.main_df[u"计算PEG"] < 1) & (self.main_df[u"CAGR"] > 20) & (
            self.main_df[u"预测PE_" + today_str[0:4]] < 25)]
        elif self.flag == "A" and  self.group == "peg_pick_high_speed":
            # 筛选出估值合理的高成长股
            self.main_df = self.main_df[(self.main_df[u"计算PEG"] < 0.85) & (self.main_df[u"CAGR"] > 40)]

        print self.group
        print self.main_df.head(5)

        self.sheet1.write(0, 0, "", self.set_style('Times New Roman', 220, True))
        row0 = self.main_df.columns.tolist()
        column0 = self.main_df.index.tolist()
        columns_count = len(row0)
        rows_count = len(column0)


        for i in range(1, columns_count + 1):
            self.sheet1.write(0, i, row0[i - 1], self.set_style('Times New Roman', 220, True, True))  # 第一行
        for i in range(0, rows_count):
            self.sheet1.write(i + 1, 0, column0[i], self.set_style('Arial', 220, True))  # 第一列
        content_style = self.set_style('Arial', 200, False, False) # 普通格式


        for j in range(1, rows_count + 1):
            for i in range(0, columns_count):
                cell_data = self.main_df[row0[i]][j - 1]
                cell_data = "" if cell_data is np.nan else cell_data

                cell_data = round(cell_data, 2) if isinstance(cell_data, float) else cell_data
                cell_data = round(cell_data, 0) if isinstance(cell_data, np.int64) else cell_data

                self.sheet1.write(j, i + 1, cell_data, content_style)

        if self.flag == "A" or self.flag == "H":
            # PEG数据模板，A股和H股前多少列格式相同，H股少了单季度数据，统一采用A股没有影响
            self.sheet1.col(0).width = 160 * 25
            self.sheet1.col(1).width = 150 * 25 # 股票简称  4个字宽120
            self.sheet1.col(2).width = 210 * 25
            self.sheet1.col(3).width = 120 * 25
            self.sheet1.col(4).width = 180 * 25
            self.sheet1.col(5).width = 180 * 25
            self.sheet1.col(6).width = 180 * 25
            self.sheet1.col(7).width = 210 * 25
            self.sheet1.col(8).width = 210 * 25
            self.sheet1.col(9).width = 210 * 25
            self.sheet1.col(10).width = 120 * 25
            self.sheet1.col(11).width = 90 * 25

            self.sheet1.col(12).width = 100 * 25
            self.sheet1.col(13).width = 180 * 25
            self.sheet1.col(14).width = 210 * 25
            self.sheet1.col(15).width = 320 * 25
            self.sheet1.col(16).width = 270 * 25
            self.sheet1.col(17).width = 120 * 25
            self.sheet1.col(18).width = 120 * 25
            self.sheet1.col(19).width = 180 * 25
            self.sheet1.col(20).width = 330 * 25
            self.sheet1.col(21).width = 320 * 25 #扣非后净利润同比增长率
            self.sheet1.col(22).width = 320 * 25
            self.sheet1.col(23).width = 340 * 25
            self.sheet1.col(24).width = 340 * 25
            self.sheet1.col(25).width = 200 * 25
            self.sheet1.col(26).width = 200 * 25
            self.sheet1.col(27).width = 200 * 25
            self.sheet1.col(28).width = 200 * 25
            self.sheet1.col(29).width = 370 * 25

        self.sheet1.write(rows_count + 1,0, self.remarks)
        tall_style = xlwt.easyxf('font:height 360;')
        first_row = self.sheet1.row(0)
        first_row.set_style(tall_style)

    def write_small_cap_undervalue_book(self):
        self.write_peg_pick_book()

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


def concat_ih_xls():
    A_stocks = Concater("A","increase_holding")
    data = A_stocks.concat_ih_df()
    print "data",data.head(5)
    A_stocks.write_ih_book(data)
    A_stocks.save_book_as_xls()

    H_stocks = Concater("H","increase_holding")
    data = H_stocks.concat_ih_df()
    H_stocks.write_ih_book(data)
    H_stocks.save_book_as_xls()



def concat_history_high_xls():
    A_stocks = Concater("A","history_high")
    data = A_stocks.concat_history_high_df()
    A_stocks.write_history_high_book(data)
    A_stocks.save_book_as_xls()
    H_stocks = Concater("H","history_high")
    data = H_stocks.concat_history_high_df()
    H_stocks.write_history_high_book(data)
    H_stocks.save_book_as_xls()

def write_peg_pick_xls():
    A_stocks = Concater("A","peg_pick")
    A_stocks.write_peg_pick_book()
    A_stocks.save_book_as_xls()

    H_stocks = Concater("H","peg_pick")
    H_stocks.write_peg_pick_book()
    H_stocks.save_book_as_xls()

    A_stocks = Concater("A","peg_pick_normal_speed")
    A_stocks.write_peg_pick_book()
    A_stocks.save_book_as_xls()
    A_stocks = Concater("A","peg_pick_high_speed")
    A_stocks.write_peg_pick_book()
    A_stocks.save_book_as_xls()


def write_small_cap_undervalue():
    A_stocks = Concater("A","small_cap_undervalue")
    A_stocks.write_small_cap_undervalue_book()
    A_stocks.save_book_as_xls()


def concat_excel_main():
    # concat_ih_xls()
    concat_history_high_xls()
    write_peg_pick_xls()
    # write_small_cap_undervalue()

if __name__ == "__main__":
    concat_excel_main()
