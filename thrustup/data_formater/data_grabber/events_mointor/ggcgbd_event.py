#!/usr/bin/env Python
# coding: utf-8

# 高管持股变动监控

import json
import sys

import pandas as pd
from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester

from freequant.thrustup.data_formater.data_grabber.formatter.ggcgbd_xlwt_formatter import write_format_xls

reload( sys )
sys.setdefaultencoding('gbk')


def request_url(url, header):
    print url
    body = urllib_requester(url, header)
    # response = body
    print body
    response = json.loads(body)
    columns = response["X"].split(",")
    money_income = response["Y"][0].split(",")
    money_income = [float(num) for num in money_income]
    money_outcome = response["Y"][1].split(",")
    money_outcome = [float(num) for num in money_outcome]
    data = {"日期":columns,"流入":money_income,"流出":money_outcome}

    response_df = pd.DataFrame(data = data)
    return response_df


def init_url():
    headers = {}
    urls = "http://data.eastmoney.com/DataCenter_V3/chart/GGCG.ashx?mkt=1&stat=2&isxml=false"

    return headers, urls


def ggcgbd_event_main():
    header, url = init_url()

    print "url",url
    data_df = request_url(url, header)

    # pure_df = purify_df(hsgtzjl_df)
    write_format_xls(data_df, u"高管持股变动", u"高管持股变动")
    return data_df


if __name__ == "__main__":
    ggcgbd_df = ggcgbd_event_main()

