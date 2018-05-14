
#coding: utf-8

import datetime as dt
import json
import time

import pandas as pd
from bs4 import BeautifulSoup

from freequant.thrustup.data_formater.data_grabber.common.toolkit import Toolkit
from freequant.thrustup.data_formater.data_grabber.formatter.announcement_xlwt_formatter import write_format_xls
from snowball_login import make_session

session, header = make_session()

timestamp = Toolkit.getUserData('config/timestamp.data')
last_timestamp = time.strptime(timestamp["last_timestamp"], "%m-%d %H:%M")
today = dt.datetime.today()
today_str = today.strftime("%Y%m%d%H%M")
now_str = today.strftime("%m-%d %H:%M")

def get_announcement(symbol_id):
    fav_temp = "https://xueqiu.com/statuses/stock_timeline.json?symbol_id=" + symbol_id + "&count=30&source=公告&page=1"
    # fav_temp = "https://xueqiu.com/statuses/stock_timeline.json?symbol_id=" + symbol_id + "&count=30&source=研报&page=1"
    # print "fav_temp",fav_temp
    collection = session.get(fav_temp, headers=header)
    context = collection.text
    # print collection.status_code, collection.encoding
    notice = json.loads(context)
    # print "notice",notice
    notice_list = notice["list"]

    announcements = []
    # print notice_list
    for notice in notice_list:
        # print notice
        if "2017" in notice["timeBefore"]:
            break
        if ("今天") in notice["timeBefore"].encode("utf-8") or \
                        ("前") in notice["timeBefore"].encode("utf-8") or \
                        time.strptime(notice["timeBefore"], "%m-%d %H:%M") > last_timestamp:
            if ".pdf" in notice["description"] or ".PDF" in notice["description"]:
                if "SH" in symbol_id or "SZ" in symbol_id: # A股
                    description = notice["description"].split(": ")
                    title = description[0].split(u"：")[-1].split(" ")[0] if "SZ" in symbol_id \
                        else description[0].split(u"： ")[-1].split(" ")[0]
                    pdf = BeautifulSoup(description[1], 'html.parser')
                    announcements.append([notice["user"]["screen_name"],notice["timeBefore"],
                                          title, pdf.find_all('a')[0].get("href")])
                elif symbol_id[0] == "0": # H股
                    description = notice["description"].split(", PDF)")
                    # 用screen_name的最后一个字对title进行split
                    title = description[0]+")"
                    # title = description[0].split(notice["user"]["screen_name"].split("(")[0][-1])[1]
                    pdf = BeautifulSoup(description[1], 'html.parser')
                    announcements.append([notice["user"]["screen_name"],notice["timeBefore"],
                                          title, pdf.find_all('a')[0].get("href")])
            else:
                if "SH" in symbol_id or "SZ" in symbol_id:  # A股
                    pass
                elif symbol_id[0] == "0":  # H股
                    pass
                else: # 美股
                    description = BeautifulSoup(notice["description"], 'html.parser')
                    # print "description", description
                    descp = description.get_text()
                    print descp
                    if "EPS " not in descp and " Filed:" in descp:
                        title = descp.split(" Filed:")[0].split(" - ")[1]
                    elif "$ " in descp and u" 网页链接" in descp:
                        title = descp.split("$ ")[1].split(u" 网页链接")[0]
                    else:
                        title = descp.split("$ ")[1]
                    print title

                    if len(description.find_all('a')) > 1:
                        href = description.find_all('a')[1].get("href")
                    else:
                        href = ""
                    announcements.append([notice["user"]["screen_name"],notice["timeBefore"],
                                          title,href])
        else:
            break
    return announcements


def get_research(symbol_id):
    fav_temp = "https://xueqiu.com/statuses/stock_timeline.json?symbol_id=" + symbol_id + "&count=30&source=研报&page=1"
    collection = session.get(fav_temp, headers=header)
    context = collection.text
    notice = json.loads(context)
    # print "notice",notice
    notice_list = notice["list"]
    # print "notice_list",notice_list

    for notice in notice_list:
        print notice["user"]["screen_name"]
        print "title",notice["title"]
        if ("今天") in notice["timeBefore"].encode("utf-8") or \
                        ("前") in notice["timeBefore"].encode("utf-8") or \
                        time.strptime(notice["timeBefore"], "%m-%d %H:%M") > last_timestamp:
            print notice["timeBefore"], notice["description"]
            print "target","https://xueqiu.com" + notice["target"]
        else:
            break

def get_self_selection_stocks():
    stocks_df = pd.read_excel("config/__self_selection.xlsx", header=0)
    stocks_code = stocks_df["codes"].tolist()
    return stocks_code

def get_snowball_code(code):
    splited = code.split(".")
    if splited[1] == "SZ" or splited[1] == "SH":
        return splited[1] + splited[0]
    elif splited[1] == "HK":
        return splited[0].zfill(5)
    elif splited[1] == "O" or splited[1] == "N":
        return splited[0]
    else:
        print "wrong snowball code: ",code
        return

def announcement_main():

    symbol_ids = get_self_selection_stocks()
    print symbol_ids
    symbols_announcements = []
    # try:
    for key, symbol_id in enumerate(symbol_ids):
        snowball_code = get_snowball_code(symbol_id)
        if not snowball_code:
            continue
        code = snowball_code.encode("utf-8")
        single_announcements = get_announcement(code) # 公告

        print symbol_id, single_announcements
        symbols_announcements.extend(single_announcements)

        # if "SH" in symbol_id or "SZ" in symbol_id:
        #     get_research(code) # 研报
        # break
        time.sleep(4)
        if key %10 == 0:
            time.sleep(5)
    # except Exception as e:
    #     print e

    df = pd.DataFrame(data=symbols_announcements, columns=[u"名称",u"公告时间",u"公告标题",u"网页链接"])
    # df.set_index(df[u"公告时间"], drop=True, inplace=True)
    # del df[u"公告时间"]
    # df.to_excel("1.xls")
    write_format_xls(df,u"自选股公告采集("+ today_str+u")")
    Toolkit.saveTimestamp("config/timestamp", "last_timestamp=" + now_str)
    print timestamp["last_timestamp"] + "============>" + now_str


if __name__ == "__main__":
    announcement_main()