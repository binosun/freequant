
# coding: utf-8

import datetime as dt
import calendar

from WindPy import *

def get_last_friday():
    last_friday = dt.datetime.today()
    one_day = dt.timedelta(days=1)

    last_friday = last_friday - dt.timedelta(days=3)
    while last_friday.weekday() != calendar.FRIDAY:
        last_friday -= one_day
    return last_friday.strftime("%Y-%m-%d")

def calculate_date(flag="A"):
    w.start()
    date_list = ["1231", "0331", "0630", "0930"]
    now_time = dt.datetime.now()
    today = dt.datetime.today()
    if dt.datetime.strptime("19 00 00", "%H %M %S") < dt.datetime.strptime("19 00 00", "%H %M %S"):
        print "<"
        pass
    today = dt.datetime.today()
    today_str = today.strftime("%Y-%m-%d")
    year_str = today.strftime("%Y")
    month_str = today.strftime("%m")
    day_str = today.strftime("%d")

    trade_date = w.tdaysoffset(0, today_str, "").Data[0][0]

    if flag == "A":
        if 3 - int(month_str) >= 0:
            if day_str != "31":
                rpt_date_str = str(int(year_str) - 1) + date_list[0]
            else:
                rpt_date_str = year_str + date_list[1]
        elif 6 - int(month_str) >= 0:
            if day_str != "30":
                rpt_date_str = year_str + date_list[1]
            else:
                rpt_date_str = year_str + date_list[2]
        elif 9 - int(month_str) >= 0:
            if day_str != "30":
                rpt_date_str = year_str + date_list[2]
            else:
                rpt_date_str = year_str + date_list[3]
        elif 12 - int(month_str) >= 0:
            if day_str != "31":
                rpt_date_str = year_str + date_list[3]
            else:
                rpt_date_str = year_str + date_list[3]

        index = date_list.index(rpt_date_str[4:])
        if index == 0:
            last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + date_list[3]
        else:
            last_rpt_date_str = rpt_date_str[0:4] + date_list[index - 1]

    else:
        if 6 - int(month_str) >= 0:
            if day_str != "30":
                rpt_date_str = str(int(year_str) - 1) + date_list[0]
            else:
                rpt_date_str = year_str + date_list[2]
        else:
            rpt_date_str = year_str + date_list[2]

        if rpt_date_str[4:] == date_list[2]:
            last_rpt_date_str = str(int(rpt_date_str[0:4]) - 1) + date_list[0]
        else:
            last_rpt_date_str = rpt_date_str[0:4] + date_list[0]

    date = {}
    date["last_year_last_rpt_date_str"] = str(int(year_str) - 1) + date_list[0]
    date["trade_date"] = trade_date
    date["rpt_date_str"] = rpt_date_str
    date["year_str"] = year_str
    date["today_str"] = today_str
    date["last_rpt_date_str"] = last_rpt_date_str
    date["last_friday"] = get_last_friday()

    date["half_year_ago"] =  today - dt.timedelta(days=183)
    date["half_year_ago_str"] = (date["half_year_ago"]).strftime("%Y-%m-%d")
    # print date
    return date