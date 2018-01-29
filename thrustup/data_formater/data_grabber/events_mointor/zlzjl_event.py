
# coding: utf-8

# 主力资金流。
# 从东方财富抓取5日主力净流入前10名和主力净流出前10名。


import json
from collections import OrderedDict

import pandas as pd

from freequant.thrustup.data_formater.data_grabber.common.url_requester import urllib_requester
from freequant.thrustup.data_formater.data_grabber.formatter.zlzjl_xlwt_formatter import write_format_xls


def request_url(url, header):
    body = urllib_requester(url, header)
    response = json.loads(body[1:-1])
    response_list = []
    for string in response:
        single_industry = string.split(",")
        single_industry[-1] = float(single_industry[-1])
        response_list.append(single_industry)
    return response_list


def init_url():
    headers = {}

    urls = OrderedDict([
        ("zljlr","http://nufm.dfcfw.com/EM_Finance2014NumericApplication/JS.aspx?type=CT&cmd=C._BKHY&sty=DCFFPBFM5&"
                 "st=(BalFlowMainNet5)&sr=-1&p=1&ps=10&js=&token=894050c76af8597a853f5b408b759f5d&"),

        ("zljlc", "http://nufm.dfcfw.com/EM_Finance2014NumericApplication/JS.aspx?type=CT&cmd=C._BKHY&sty=DCFFPBFM5&"
               "st=(BalFlowMainNet5)&sr=1&p=1&ps=10&js=&token=894050c76af8597a853f5b408b759f5d&")
    ])
    "http://nufm.dfcfw.com/EM_Finance2014NumericApplication/JS.aspx?type=CT&cmd=C._BKHY&sty=DCFFPBFM5&" \
    "st=(BalFlowMainNet5)&sr=-1&p=1&ps=10000&js=&token=894050c76af8597a853f5b408b759f5d&"
    return headers, urls

def purify_df(df):
    drop_list = ["column1","column2"]
    columns = df.columns.tolist()
    for column in drop_list:
        if column in columns:
            df.drop(labels=[column], axis=1, inplace=True)

    df[u"金额(亿元)"] = df[u"金额(亿元)"]/10e4
    df.sort_values(by=u"金额(亿元)", ascending=True, inplace=True)

    return df


def zlzjl_event_main():
    header, url_dict = init_url()
    df_list = []
    for key in url_dict:
        url = url_dict[key]
        data_list = request_url(url, header)
        df_list.extend(data_list)

    zlzjl_df = pd.DataFrame(data=df_list,columns=["column1","column2",u"行业",u"金额(亿元)"])
    print "zlzjl_df", zlzjl_df
    pure_df = purify_df(zlzjl_df)
    write_format_xls(pure_df, u"5日主力动向", u"5日主力动向")

    return pure_df


if __name__ == "__main__":
    zlzjl_df = zlzjl_event_main()

    # import numpy as np
    # import seaborn as sns
    # import matplotlib.pyplot as plt
    #
    # sns.set(style="white")
    #
    # # Load the example planets dataset
    # planets = sns.load_dataset("planets")
    #
    # # Make a range of years to show categories with no observations
    # years = np.arange(2000, 2015)
    # print "planets",type(planets),planets
    #
    # # Draw a count plot to show the number of planets discovered each year
    # g = sns.factorplot(x="year", data=planets, kind="count",
    #                    palette="BuPu", size=6, aspect=1.5, order=years)
    # g.set_xticklabels(step=2)
    #
    # plt.savefig("1.png")
    # plt.show()

