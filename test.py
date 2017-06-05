# -*- coding: utf-8 -*-
"""
Created on Tue Feb 07 19:41:56 2017

@author: sunbin6
"""

import sys
import datetime
import numpy as np
import pandas as pd

reload(sys)  #重新加载sys
sys.setdefaultencoding("utf8")  ##调用setdefaultencoding函数

def get_record_from_account(file_name, sheet_name):
    accountant_sheet_to_df =  pd.read_excel(file_name,
                                      sheetname=sheet_name,
                                      header=0,
                                      encoding="utf-8")
    if sheet_name == u"微信":
        column_to_del = [u"时间", u"公众账号ID", u"商户号",
                        u"子商户号", u"设备号", u"微信订单号",
                        u"用户标识", u"交易类型", u"交易状态", u"付款银行", u"货币种类",
                        u"企业红包金额", u"微信退款单号", u"商户退款单号",
                        u"退款金额", u"企业红包退款金额", u"退款类型", u"退款状态",
                        u"商品名称", u"商户数据包", u"手续费", u"费率"]
        df = del_record_column(accountant_sheet_to_df, column_to_del)

    elif sheet_name == u"支付宝":
        column_to_del =[u"序号",u"时间",u"支付宝交易号",u"支付宝流水号",u"账务类型",
                       u"支出（-元）",u"账户余额（元）",u"服务费（元）",u"支付渠道",
                       u"签约产品",u"对方账户",u"对方名称",u"银行订单号",u"商品名称"]
        df =  del_record_column(accountant_sheet_to_df, column_to_del)
    df[u"订单类型"] = sheet_name
    return df
def get_product_wechat_record(product, wechat_record):
    if product == "教育":
        edu_wechat_record = wechat_record[wechat_record[u"备注"] == "教育"]
        edu_wechat_column_to_format = [u"交易时间", u"商户订单号"]
        return format_edu_wechat_record_data(edu_wechat_record, edu_wechat_column_to_format)
    else:
        pass


def format_edu_wechat_record_data(df, column_name_list):
    #将交易时间和商户订单号前的`去掉
    for column_name in column_name_list:
        series = df[column_name]
        record_list = []
        for record in series:
            record_list.append(str(record)[1:])
        del df[column_name]
        new_series = pd.Series(record_list,index = series.index)
        column_df = pd.DataFrame({column_name: new_series})
        df = pd.concat([df, column_df], axis=1)
    df["index"] = range(len(df))
    df = df.set_index(["index"])

    #将交易时间改为入账时间,将总金额改为收入（+元）
    name_map = {u"交易时间":u"入账时间",u"商户订单号":u"商户订单号",u"总金额":u"收入（+元）"}
    df.rename(columns=name_map, inplace=True)
    df = format_df_date(df)
    return df

def get_product_alipay_record(product, alipay_record):
    edu_alipay_record = alipay_record[alipay_record[u"备注"] == product]

    # print (edu_alipay_record)
    series = edu_alipay_record[u"收入（+元）"]
    del edu_alipay_record[u"收入（+元）"]
    record_list = []
    for record in series:
        if isinstance(record, int):
            record = float(record)
        elif isinstance(record, float):
            pass
        else:
            record = np.nan
        record_list.append(record)
    new_series = pd.Series(record_list, index=series.index)
    column_df = pd.DataFrame({u"收入（+元）": new_series})
    edu_alipay_record = pd.concat([edu_alipay_record, column_df], axis=1)
    edu_alipay_record = edu_alipay_record.dropna(axis=0, how="any")
    # save_record(edu_alipay_record, "sheet", "edu_alipay_record.xlsx")
    # print edu_alipay_record
    return edu_alipay_record

def del_record_column(df, edu_wechat_column_to_del):
    for column in edu_wechat_column_to_del:
        del df[column]
    return df

def format_df_date(df):
    record_list = []

    for date in df[u"入账时间"]:
        new_date = datetime.datetime.strptime(date, "%Y-%m-%d")
        record_list.append(new_date)
        # print(new_date,type(new_date))
    new_series = pd.Series(record_list,index = df[u"入账时间"].index)
    del df[u"入账时间"]
    column_df = pd.DataFrame({u"入账时间": new_series})
    df = pd.concat([df, column_df], axis=1)
    return df

def save_record(record, product, save_file_name):
    xlsx_output = pd.ExcelWriter(save_file_name)
    record.to_excel(xlsx_output, sheet_name = product)
    xlsx_output.save()


def clean_account_record(product, file_name, save_file_name):
    wechat_record = get_record_from_account(file_name, u"微信")
    alipay_record = get_record_from_account(file_name, u"支付宝")
    wechat_format_record = get_product_wechat_record(product, wechat_record)
    alipay_format_record = get_product_alipay_record(product, alipay_record)
    # save_record(alipay_format_record, "sheet", "alipay_format_record.xlsx")
    # print(alipay_format_record)
    record = pd.concat([wechat_format_record, alipay_format_record], axis = 0)
    save_record(record, product, save_file_name)
    return record

def clean_operator_record(product, file_name, save_file_name):
    if product == "教育":
        edu_record = get_record_from_operator(file_name, file_name)
        save_record(edu_record, product, save_file_name)
        return edu_record
def get_record_from_operator(file_name,sheet_name):
    if sheet_name == file_name:
        # 聚好学运营导出的Excel的Sheet名称和文件名称相同
        # 例如文件名为18-2120170207175121.xlsx，
        # 则sheet名为18-2120170207175121
        operator_sheet_to_df =  pd.read_excel(file_name,
                                      sheetname=sheet_name[:-5],
                                      header=0,
                                      encoding="utf-8")

        column_to_del = [u"tradeId",u"pay_platform",u"seller_email",u"customerName",
                         u"phone",u"userId",u"resourceId",u"resourceType",
                         u"status",u"fee",u"coupon_value",u"priceDesc",u"priceDays",
                         u"startTime",u"endTime",u"priceId",u"discountId",u"created_time",
                         u"modified_time",u"customerId",u"attach_data",u"md5_userId",u"figureNames",
                         u"firstSubjectNames",u"ageNames",u"ageStageNames",u"gradeNames",u"genderName",u"orderType"]
        column_to_save = [u"order_id",u"venderName",u"resourceName",u"pay_value",u"gmt_payment"]
        df = del_record_column(operator_sheet_to_df, column_to_del)
        name_map = {u"交易时间": u"入账时间", "order_id": u"商户订单号", "pay_value": u"收入（+元）"}
        df.rename(columns=name_map, inplace=True)
        return df

def exclude_useless_pay_value(df):
    pay_value_series = df[u"收入（+元）"]
    del df[u"收入（+元）"]
    new_pay_value_list = []
    for pay_value in pay_value_series:
        if pay_value == 0:
            pay_value = np.nan
        new_pay_value_list.append(pay_value)


def format_operator_record(product,df):
    if product == "教育":
        pass

if __name__ == "__main__":
    account_file_name = "accountant_date"
    account_record = clean_account_record("教育",account_file_name  + ".xlsx" , account_file_name  + "_output" + ".xlsx")
    new_account_record = account_record.set_index(u"商户订单号")
    # print new_account_record
    operator_file_name = "18-2120170207175121"
    edu_operator_record = clean_operator_record("教育",operator_file_name  + ".xlsx" , operator_file_name  + "_output" + ".xlsx")
    new_edu_operator_record = edu_operator_record.set_index(u"商户订单号")
    # print new_edu_operator_record
    merged_record = pd.merge(account_record, edu_operator_record, how="outer")
    save_record(merged_record, "教育", "merge_record.xlsx")
    # print (merged_record)
    print ("--------finish------")