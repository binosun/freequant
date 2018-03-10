
# coding: utf-8
import datetime as dt
import os
import shutil

def purify_file():
    today = dt.datetime.today()
    today_str = today.strftime("%Y-%m-%d")

    date_mark = "data_output"+ "/" + today_str + "/"
    del_folder = date_mark+"del/"


    file_name_list = ["ih_A.xlsx","ih_H.xlsx"]

    last_name_list = ["history_high_A_count","history_high_H_count","increase_holding_A","increase_holding_H"]
    for last_name in last_name_list:
        file_name = last_name + today_str + ".xls"
        file_name_list.append(file_name)

    file_name_list.append((today - dt.timedelta(days=1)).strftime("%Y%m%d") + u"沪港通持股.xlsx")
    file_name_list.append((today - dt.timedelta(days=8)).strftime("%Y%m%d") + u"沪港通持股.xlsx")
    file_name_list.append((today - dt.timedelta(days=1)).strftime("%Y%m%d") + u"深港通持股.xlsx")
    file_name_list.append((today - dt.timedelta(days=8)).strftime("%Y%m%d") + u"深港通持股.xlsx")
    file_name_list.append((today - dt.timedelta(days=1)).strftime("%Y%m%d") + u"沪深港通持股.xlsx")
    file_name_list.append((today - dt.timedelta(days=8)).strftime("%Y%m%d") + u"沪深港通持股.xlsx")

    if not os.path.exists(del_folder):
        print "创建文件目录", del_folder
        os.mkdir(del_folder)

    for file in file_name_list:
        abs_file_path = date_mark + file
        if not os.path.exists(abs_file_path):
            print "文件丢失",abs_file_path
        else:
            shutil.move(abs_file_path, del_folder)


if __name__ == "__main__":
    purify_file()

