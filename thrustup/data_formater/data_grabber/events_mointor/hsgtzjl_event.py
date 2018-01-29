
# coding: utf-8

# 沪深港通资金流量监控



import json
from collections import OrderedDict

import pandas as pd

from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester

from freequant.thrustup.data_formater.common.date_counter import calculate_date
from freequant.thrustup.data_formater.data_grabber.formatter.hsgtzjl_xlwt_formatter import write_format_xls


def request_url(url, header):
    body = urllib_requester(url, header)
    # print "body[8:-3]",body[8:-3]
    response = json.loads(body[8:-3])

    return response


def init_url():
    headers = {}
    date = calculate_date()
    half_year_ago_str = date["half_year_ago_str"]
    urls = OrderedDict([
        ("zljlr","http://dcfm.eastmoney.com/EM_MutiSvcExpandInterface/api/js/get?type=HSGTZJZS&"
                 "token=70f12f2f4f091e459a279469fe49eca5&js=({data:[(x)]})&"
                 "filter=(DateTime%3E^"+half_year_ago_str+"^)")
    ])

    return headers, urls

def purify_df(df):
    name_pairs = [("DateTime",u"日期"),("HSMoney",u"沪股通"),("SSMoney",u"深股通"),("NorthMoney",u"北向资金（百万元）"),
    ("GGHSMoney",u"港股通（沪）"),("GGSSMoney",u"港股通（深）"),("SouthSumMoney",u"南向资金（百万元）")]
    field_map = OrderedDict(name_pairs)
    clean_df = df[field_map.keys()]
    clean_df.rename(columns=field_map, inplace=True)
    print "clean_df.head",clean_df.head(5)
    return clean_df


def hsgtzjl_event_main():
    header, url_dict = init_url()
    df_list = []
    for key in url_dict:
        url = url_dict[key]
        print "url",url
        data_list = request_url(url, header)
        df_list.extend(data_list)

    hsgtzjl_df = pd.DataFrame(data=df_list)
    pure_df = purify_df(hsgtzjl_df)
    write_format_xls(pure_df, u"沪深股通资金流", u"沪深股通资金流")
    return pure_df


if __name__ == "__main__":
    hsgtzjl_df = hsgtzjl_event_main()

