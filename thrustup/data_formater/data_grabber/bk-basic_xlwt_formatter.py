
# coding: utf-8

from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import json
import xlwt
import copy
import os
import datetime as dt




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

def create_book(SheetName="Sheet1"):
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet(SheetName)
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

def write_format_xls(df,xls_name,sheet_name="Sheet1"):
    book, sheet1 = create_book(sheet_name)
    write_ih_book(df, sheet1)
    folders = ["data/"]
    child_folder = check_folder(folders)
    book.save(folders[0] + child_folder + "/" +xls_name + ".xls")

def write_simple_xls(df,xls_name):
    folders = ["data/"]
    child_folder = check_folder(folders)
    df.to_excel(folders[0] + child_folder + "/" +xls_name + ".xls")

class EastMoneyCrawler():

    def __init__(self, crawler_tag):
        self.driver = webdriver.PhantomJS()
        self.tag = crawler_tag

    def parse_web(self,web):
        driver = self.driver
        driver.get(web)
        web_soup = BeautifulSoup(driver.page_source, "lxml")
        # print soup.prettify()
        self.get_df(web_soup)

    def get_df(self,soup):
        if self.tag == "jzc" or "jjc":
            column_names = [u"代码",u"名称",u"公告",u"最新价",u"涨跌幅",u"股东名称",u"增减",u"变动数量（万股）",
                            u"占总股本比例",u"占流通股比例",u"持股总数（万股）",u"占总股本比例",u"持流通股数",
                            u"占流通股比例",u"变动开始日",u"变动截止日",u"公告日"]

        data = self.get_data(soup)
        data_df = pd.DataFrame(data,columns=column_names)
        print data_df
        data_df.to_excel("jzcData.xls")

    def get_data(self, soup):
        data = []
        tables = soup.findAll('table')
        tab = tables[0]
        for tr in tab.findAll('tr'):
            rows = []
            for td in tr.findAll('td'):
                rows.append(td.getText())
            print rows
            data.append(rows)
        return data


if __name__ == "__main__":
    # jzc_html = 'http://data.eastmoney.com/executive/gdzjc-jzc.html'
    # jzc_crawer = EastMoneyCrawler("jzc") # 净增持
    # jzc_crawer.parse_web(jzc_html)



    # driver = webdriver.PhantomJS()
    # jzc_html = "http://datainterface3.eastmoney.com/EM_DataCenter_V3/api/GDZC/GetGDZC?tkn=eastmoney&cfg=gdzc&secucode=&fx=1&sharehdname=&pageSize=50&pageNum=1&sortFields=BDJZ&sortDirec=1&startDate=2017-11-29&endDate=2017-11-30"
    # driver.get(jzc_html)
    # web_soup = BeautifulSoup(driver.page_source, "lxml")
    # print web_soup.prettify()


    import urllib2

    start_data = "2017-11-24"  # 公告日的起始日期
    end_data = "2017-12-01"    # 公告日的截止日期
    jzc_html = "http://datainterface3.eastmoney.com/EM_DataCenter_V3/api/GDZC/GetGDZC?" \
               "tkn=eastmoney&cfg=gdzc&secucode=&sharehdname=&pageSize=200&pageNum=1&sortFields=NOTICEDATE&sortDirec=1" \
               "&fx=1"\
               "&startDate="+start_data+"&endDate="+end_data
    jjc_html = "http://datainterface3.eastmoney.com/EM_DataCenter_V3/api/GDZC/GetGDZC?" \
               "tkn=eastmoney&cfg=gdzc&secucode=&sharehdname=&pageSize=200&pageNum=1&sortFields=NOTICEDATE&sortDirec=1" \
               "&fx=2"\
               "&startDate="+start_data+"&endDate="+end_data
    request = urllib2.Request(jjc_html)
    response = urllib2.urlopen(request)
    body = json.loads(response.read())
    print type(body)
    raw_data = body["Data"][0]
    split_symbol = raw_data["SplitSymbol"]
    field_name = raw_data["FieldName"].split(",")
    inner_data = raw_data["Data"]
    df_data = []
    for str_data in inner_data:
        df_row_data = str_data.split(split_symbol)
        df_data.append(df_row_data)


    print "len(field_name)",len(field_name)
    data_df = pd.DataFrame(df_data, columns=field_name)
    print data_df
    del data_df["SHCode"]
    del data_df["CompanyCode"]
    del data_df["Close"]
    del data_df["ChangePercent"]
    del data_df["JYFS"]
    del data_df["BDKS"]
    del data_df["BDZGBBL"]
    del data_df["BDHCYLTGSL"]
    del data_df["BDJZ"]
    data_df.rename(columns={"SCode":u"代码","SName":u"名称","ShareHdName":u"股东名称","FX":u"增减",
                            "ChangeNum":u"变动数量(万股)","BDSLZLTB":u"占总流通股比例%",
                            "BDHCGZS":u"变动后持股总数(万股)","BDHCGBL":u"变动后持股比例%",
                            "BDHCYLTSLZLTGB":u"变动后占总流通股比例%","BDJZ":u"变动截止日","NOTICEDATE":u"公告日"},
                   inplace=True)
    # data_df.to_excel("jzcData.xlsx")
    book, sheet1 = create_book()
    write_ih_book(data_df, sheet1)

    book.save("jjcData.xls")