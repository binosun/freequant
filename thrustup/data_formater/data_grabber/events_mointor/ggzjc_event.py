# coding: utf-8


import json
from collections import OrderedDict

import pandas as pd
from WindPy import *
from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester

from freequant.thrustup.data_formater.data_grabber.common.wind_handler import get_wind_basic_data
from freequant.thrustup.data_formater.data_grabber.formatter.zjc_xlwt_formatter import write_format_xls


def request_url(url, header):
    url_response = urllib_requester(url,header)
    result = url_response[28:-1]
    print "result",result,type(result)
    return result


def init_url():
    headers = {}
    urls = OrderedDict([(
        "ggjzc", "http://datainterface.eastmoney.com/EM_DataCenter/JS.aspx?"
                "type=GG&sty=ZCPHB&p=1&ps=50"
                "&js=var%20XGGGPSwk={pages:(pc),data:[(x)]}&sr=true&stat=1&st=1&rt=50418311"),
        ("ggjjc", "http://datainterface.eastmoney.com/EM_DataCenter/JS.aspx?"
               "type=GG&sty=ZCPHB&p=1&ps=50"
               "&js=var%20DoyiJlWo={pages:(pc),data:[(x)]}&sr=false&stat=1&st=1&rt=50418322")
    ])

    return headers, urls


def get_ggzjc_df(raw_data):
    inner_data = json.loads(raw_data)
    print "inner_data",inner_data,type(inner_data)

    df_data = []
    field_name = [u"代码",u"名称",u"股份金额(万)",u"股份数额(万股)",u"增减持均价(元)",u"最新价(元)",u"涨跌幅"]
    for str_data in inner_data:
        df_row_data = str_data.split(",")
        df_data.append(df_row_data)

    # print "len(field_name)", len(field_name)
    data_df = pd.DataFrame(df_data, columns=field_name)
    print data_df.head(3)

    stocks = [stock + ".SH" if stock[0] == "6" else stock + ".SZ" for stock in data_df[u"代码"]]
    data_df[u"代码"] = pd.Series(stocks, index=data_df.index)
    data_df.set_index(data_df[u"代码"], drop=True, inplace=True)

    stocks_share = [float(share) for share in data_df[u"股份金额(万)"]]
    data_df[u"股份金额(万)"] = pd.Series(stocks_share, index=data_df.index)
    data_df[u"股份金额(万)"] = data_df[u"股份金额(万)"]/10000

    stocks_share = [float(share) for share in data_df[u"股份数额(万股)"]]
    data_df[u"股份数额(万股)"] = pd.Series(stocks_share, index=data_df.index)
    data_df[u"股份数额(万股)"] = data_df[u"股份数额(万股)"] / 10000

    stocks_share = [round(float(share), 2) for share in data_df[u"增减持均价(元)"]]
    data_df[u"增减持均价(元)"] = pd.Series(stocks_share, index=data_df.index)
    # data_df[u"增减持均价(元)"] = data_df[u"增减持均价(元)"] / 10000

    # data_df.to_excel("jzcData.xlsx")
    return data_df


# def group_zjc_df(zjc_df, key):
#     if key == "jjz":
#         grouped_df = zjc_df.groupby([u"代码"]).max()
#     else:
#         grouped_df = zjc_df.groupby([u"代码"]).min()
#         grouped_df[u"最新公告日期"] = zjc_df[u"最新公告日期"].groupby(zjc_df[u"代码"]).max()
#     grouped_df[u"变动数量(万股)"] = zjc_df[u"变动数量(万股)"].groupby(zjc_df[u"代码"]).sum()
#     grouped_df[u"占总流通股比例%"] = zjc_df[u"占总流通股比例%"].groupby(zjc_df[u"代码"]).sum()
#     # grouped_df[u"变动后持股比例%"] = zjc_df[u"变动后持股比例%"].groupby(zjc_df[u"代码"]).sum()
#     grouped_df.sort_values(by=u"变动数量(万股)", ascending=False, inplace=True)
#     return grouped_df


def purify_zjc_df(df):
    df.drop(labels=[u"代码",u"名称", u"最新价(元)",u"涨跌幅",u"近期创历史新高", u"上市日期",
                    u"雪球一周关注增长率%", u"雪球累计关注人数"], axis=1, inplace=True)

    columns_order = OrderedDict([(u'股票简称', 0), (u"申万行业代码(三级行业)", 1), (u"收盘价", 2)])
    for column_name in columns_order:
        mid = df[column_name]
        df.drop(labels=[column_name], axis=1, inplace=True)
        df.insert(columns_order[column_name], column_name, mid)
        # print column_name,df.head(5)

    return df


def ggzjc_event_main():
    header, url_dict = init_url()
    df_list = []
    ggzjc_sheet_name = {"ggjzc": u"高管净增持", "ggjjc": u"高管净减持"}
    for key in url_dict:
        url = url_dict[key]
        print ggzjc_sheet_name[key]
        response_data = request_url(url, header)
        ggzjc_df = get_ggzjc_df(response_data)
        # write_format_xls(ggzjc_df, key)

        # grouped_df = group_zjc_df(zjc_df, key)
        # write_format_xls(grouped_df, key)

        format_stocks = ggzjc_df.index.tolist()
        wind_df = get_wind_basic_data(format_stocks)

        new_zjc_df = pd.concat((ggzjc_df, wind_df), axis=1)
        pure_df = purify_zjc_df(new_zjc_df)
        write_format_xls(pure_df, ggzjc_sheet_name[key], ggzjc_sheet_name[key])
        # break

    return df_list  # [jzc, jjc]


if __name__ == "__main__":
    w.start()
    zjc_df_list = ggzjc_event_main()