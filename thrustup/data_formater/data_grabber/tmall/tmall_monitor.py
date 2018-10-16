
# coding: utf-8

from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import datetime as dt
import numpy as np
import time

class ProductCrawler():
    # 获得单品的月销量和累计评价
    def __init__(self,df):
        service_args = ['--load-images=false', '--proxy-type=None']
        self.driver = webdriver.PhantomJS(service_args=service_args)
        self.base_url = "https://detail.tmall.com/item.htm?id="
        self.product_msg = df

    def get_sell_count(self,id):
        driver = self.driver
        web_url = self.base_url + str(id)
        driver.get(web_url)
        web_soup = BeautifulSoup(driver.page_source, "lxml")
        # print web_soup.prettify()

        # 销量数据
        sell_count = web_soup.find_all("li", class_="tm-ind-item tm-ind-sellCount ")
        txt_list = sell_count[0].find_all("span")
        print "txt_list", txt_list
        return txt_list

    def parse_web(self):

        ids = self.product_msg["id"].tolist()

        now_time = dt.datetime.now()
        self.product_msg[now_time.strftime("%Y%m%d")] = np.nan

        today_data = []

        for index ,id in enumerate(ids):
            txt_list = self.get_sell_count(id)
            if txt_list:
                today_data.append(int(txt_list[1].string))
            else:
                txt_list = self.get_sell_count(id)
                if txt_list:
                    today_data.append(int(txt_list[1].string))
                else:
                    today_data.append(np.nan)

        today_series = pd.Series(today_data,index=self.product_msg.index)
        self.product_msg[now_time.strftime("%Y%m%d")] = today_series
        # self.product_msg.iloc[now_time.strftime("%Y%m%d"), id] = int(txt_list[1].string)

        print self.product_msg
        self.product_msg.to_excel("result.xls")

def get_products():
    import pandas as pd
    product_df = pd.read_excel("product_list.xlsx",sheetname="lining")
    return product_df


if __name__ == "__main__":
    # products = get_products()
    # print products
    # tmall_crawler = ProductCrawler(products)
    # tmall_crawler.parse_web()


    driver = webdriver.PhantomJS()
    # driver = webdriver.Chrome()

    url = "https://anta.m.tmall.com/shop/shop_auction_search.htm?sort=hotsell"

    # url = "https://skechers.tmall.com/category.htm?orderType=hotsell_desc"
    # url = "https://anta.tmall.com/category.htm?orderType=hotsell_desc"

    driver.get(url)
    # time.sleep(10)
    web_soup = BeautifulSoup(driver.page_source, "lxml")
    driver.quit()
    # # print "web_soup.prettify()",web_soup.prettify()
    # print "---------------"
    # item_lines = web_soup.find_all("div", class_="item4line1")
    # print "item_lines",item_lines
    #
    # if item_lines:
    #     print "item_lines[0]",item_lines[0]
    #     for items in item_lines[0:2]:
    #         sale_num = items.find_all("span", class_="sale-num")
    #         for num in sale_num:
    #             print num.string

    print "web_soup.prettify()", web_soup.prettify()
    item_lines = web_soup.find_all("span", class_="tii_sold")
    print "item_lines",item_lines
    for item in item_lines:
        print item.string


