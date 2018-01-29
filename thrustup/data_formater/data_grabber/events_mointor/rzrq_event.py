
# coding: utf-8

# 沪深港通资金流量监控



import json

import pandas as pd

from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester
from freequant.thrustup.data_formater.data_grabber.formatter.rzrq_xlwt_formatter import write_format_xls


def request_url(url, header):
    body = urllib_requester(url, header)
    # print "body",body
    response = json.loads(body)
    # print "response",response
    return response


def init_url():
    headers = {}
    urls = "http://dcfm.eastmoney.com/EM_MutiSvcExpandInterface/api/js/get?type=RZRQ_LSTOTAL_NJ&" \
           "token=70f12f2f4f091e459a279469fe49eca5&p=1&ps=240&st=tdate&sr=-1&js=(x)"
    return headers, urls

def purify_df(df):
    clean_df  = pd.DataFrame()
    clean_df[u"日期"] = df["tdate"]
    clean_df[u"融资余额（两市，亿元）"] = df["rzye"]/1e8
    clean_df[u"融券余额（两市，亿元）"] = df["rqye"]/1e8
    clean_df[u"融资融券余额（两市，亿元）"] = df["rzrqye"]/1e8

    # name_pairs = [("tdate",u"日期"),("rzye",u"融资余额（两市，元）"),
    #               ("rqye",u"融券余额（两市，元）"),("rzrqye",u"融资融券余额（两市，元）")]
    # field_map = OrderedDict(name_pairs)
    # clean_df = df[field_map.keys()]
    # clean_df.rename(columns=field_map, inplace=True)
    # print "clean_df.head",clean_df.head(5)

    print "clean_df",clean_df.head(50)
    return clean_df

def rzrq_event_main():
    header, url = init_url()

    print "url",url
    data = request_url(url, header)
    rzrq_df = pd.DataFrame(data=data)
    print "rzrq_df.head",rzrq_df.head(5)
    pure_df = purify_df(rzrq_df)

    write_format_xls(pure_df, u"两融余额", u"两融余额")
    return pure_df


if __name__ == "__main__":
    rzrq_df = rzrq_event_main()

