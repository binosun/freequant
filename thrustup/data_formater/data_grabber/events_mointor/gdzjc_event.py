
# coding: utf-8


import json
from collections import OrderedDict

import pandas as pd
from WindPy import *

from freequant.thrustup.data_formater.data_grabber.common.date_counter import calculate_date
from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester
from freequant.thrustup.data_formater.data_grabber.common.wind_handler import get_wind_basic_data
from freequant.thrustup.data_formater.data_grabber.formatter.zjc_xlwt_formatter import write_format_xls


def request_url(url, header):
    body = json.loads(urllib_requester(url, header))
    return body

def init_url():
    headers = {}

    dates = calculate_date()
    start_date_str = raw_input("股东增减持公告日的起始日期(回车默认上周五),例20171124:")
    if not start_date_str:
        start_date_str = dates["last_friday"]
        start_date = datetime.strptime(start_date_str, "%Y-%M-%d")
    else:
        start_date = datetime.strptime(start_date_str, "%Y%M%d")

    end_date_str = raw_input("股东增减持公告日的截止日期,例20171130:")
    end_date = datetime.strptime(end_date_str, "%Y%M%d")

    start_date = start_date.strftime("%Y-%M-%d")
    end_date = end_date.strftime("%Y-%M-%d")

    urls = {
                "jzc" : "http://datainterface3.eastmoney.com/EM_DataCenter_V3/api/GDZC/GetGDZC?" \
                           "tkn=eastmoney&cfg=gdzc&secucode=&sharehdname=&pageSize=1000&pageNum=1&sortFields=NOTICEDATE&sortDirec=1" \
                           "&fx=1&startDate=" + start_date + "&endDate=" + end_date,
                "jjc" : "http://datainterface3.eastmoney.com/EM_DataCenter_V3/api/GDZC/GetGDZC?" \
                           "tkn=eastmoney&cfg=gdzc&secucode=&sharehdname=&pageSize=1000&pageNum=1&sortFields=NOTICEDATE&sortDirec=1" \
                           "&fx=2&startDate=" + start_date + "&endDate=" + end_date
                }
    return headers, urls

def get_zjc_df(raw_dict):
    raw_data = raw_dict["Data"][0]
    split_symbol = raw_data["SplitSymbol"]
    field_name = raw_data["FieldName"].split(",")
    inner_data = raw_data["Data"]
    df_data = []
    for str_data in inner_data:
        df_row_data = str_data.split(split_symbol)
        df_data.append(df_row_data)

    data_df = pd.DataFrame(df_data, columns=field_name)

    data_df.drop(labels=["SHCode", "CompanyCode", "Close", "ChangePercent", "JYFS", "BDKS", "BDZGBBL",
                         "BDHCYLTGSL", "BDJZ"], axis=1, inplace=True)

    del data_df["BDHCGBL"]  # 变动后持股比例%
    del data_df["BDHCGZS"]  # 变动后持股总数(万股)
    del data_df["BDHCYLTSLZLTGB"] # 变动后持流通股占总流通股比
    del data_df["ShareHdName"] # 股东名称
    stocks = [stock + ".SH" if stock[0] == "6" else stock + ".SZ" for stock in data_df["SCode"]]
    data_df["SCode"] = pd.Series(stocks, index=data_df.index)

    data_df["ChangeNum"] = pd.Series([float(cell.encode("utf-8")) for cell in data_df["ChangeNum"]], index=data_df.index)
    list_BDSLZLTB = []
    for item in data_df["BDSLZLTB"]:
        cell = item.encode("utf-8")
        num = float(cell.encode("utf-8")) if cell else 0.0
        list_BDSLZLTB.append(num)
    data_df["BDSLZLTB"] = pd.Series(list_BDSLZLTB, index=data_df.index)
    data_df.rename(columns={"SCode"         : u"代码", "SName": u"名称", "FX": u"增减",
                            "ChangeNum"     : u"变动数量(万股)", "BDSLZLTB": u"占总流通股比例%",
                            "BDHCYLTSLZLTGB": u"变动后持流通股占总流通股比%",
                             "BDJZ"         : u"变动截止日期",
                            "NOTICEDATE"    : u"最新公告日期"},
                   inplace=True)
    return data_df

def group_zjc_df(zjc_df, key):
    if key == "jjz":
        grouped_df = zjc_df.groupby([u"代码"]).max()
    else:
        grouped_df = zjc_df.groupby([u"代码"]).min()
        grouped_df[u"最新公告日期"] = zjc_df[u"最新公告日期"].groupby(zjc_df[u"代码"]).max()
    grouped_df[u"变动数量(万股)"] = zjc_df[u"变动数量(万股)"].groupby(zjc_df[u"代码"]).sum()
    grouped_df[u"占总流通股比例%"] = zjc_df[u"占总流通股比例%"].groupby(zjc_df[u"代码"]).sum()
    grouped_df.sort_values(by=u"占总流通股比例%", ascending=False, inplace=True)
    return grouped_df

def purify_zjc_df(df):
    drop_list = [u"名称",u"近期创历史新高",u"上市日期",u"雪球一周关注增长率%",u"雪球累计关注人数"]
    columns = df.columns.tolist()
    for column in drop_list:
        if column in columns:
            df.drop(labels=[column], axis=1,inplace = True)

    columns_order = OrderedDict([(u'股票简称',0), (u"申万行业代码(三级行业)",1), (u"收盘价",2)])
    for column_name in columns_order:
        mid = df[column_name]
        df.drop(labels=[column_name], axis=1, inplace=True)
        df.insert(columns_order[column_name], column_name, mid)
    return df

def gdzjc_event_main():
    header, url_dict = init_url()
    df_list = []
    zjc_sheet_name = {"jzc": u"股东增持", "jjc": u"股东减持"}
    for key in url_dict:

        url = url_dict[key]
        response_dict = request_url(url, header)
        zjc_df = get_zjc_df(response_dict)
        # write_simple_xls(zjc_df, "raw_" + key)

        grouped_df = group_zjc_df(zjc_df,key)
        # write_format_xls(grouped_df, key)
        write_format_xls(grouped_df, zjc_sheet_name[key], zjc_sheet_name[key])

        # format_stocks = grouped_df.index.tolist()
        # wind_df = get_wind_basic_data(format_stocks)
        #
        # new_zjc_df = pd.concat((grouped_df,wind_df), axis=1)
        # pure_df = purify_zjc_df(new_zjc_df)
        # write_format_xls(pure_df, zjc_sheet_name[key], zjc_sheet_name[key])

    return df_list  # [jzc, jjc]



if __name__ == "__main__":
    w.start()
    zjc_df_list = gdzjc_event_main()




    