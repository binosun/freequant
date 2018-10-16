
# coding: utf-8

import json
import time
import requests
import pandas as pd
import datetime as dt
from sqlalchemy import create_engine


now_time = dt.datetime.now()
today_str = now_time.strftime("%Y%m%d")

session = requests.session()


def get_today_yesterday():
    now = time.gmtime()
    today = dt.datetime(now[0], now[1], now[2])
    yesterday = today - dt.timedelta(days=1)
    return today, yesterday

class TmallMonitor():

    def __init__(self, brand):
        self.brand = brand
        self.count_limit = 10
        self.url = "https://" + brand + ".m.tmall.com/shop/shop_auction_search.do?sort=d&p="

    def get_all_items(self):
        all_items = []
        for page in range(1,4+1):
            page_url = self.url + str(page)
            print page_url

            page_collection = session.get(page_url)
            page_context = page_collection.text
            page_data = json.loads(page_context)
            all_items.extend(page_data["items"])
        return all_items

    def purify_items(self, raw_items):
        shoes_category = []
        others_category = []
        if all_items:
            for item in all_items:
                purify_item = {}

                if u"鞋" in item["title"]:
                    purify_item["category"] = "shoes"
                    shoes_category.append(purify_item)
                else:
                    purify_item["category"] = "others"
                    others_category.append(purify_item)

                purify_item["sold"] = int(item[u"sold"])
                purify_item["item_id"] = item[u"item_id"]
        shoes_category = shoes_category[0:10]
        others_category = others_category[0:10]
        new_items = {}
        for index ,item in enumerate(shoes_category):
            new_items["shoes" + str(index+1)] = item["sold"]
        if self.brand != "skechers":
            for index ,item in enumerate(others_category):
                new_items["others" + str(index+1)] = item["sold"]

        return new_items


    def insert_data(self,items):
        today, yesterday = get_today_yesterday()
        df = pd.DataFrame(items,index=[today])

        engine = create_engine('mysql://root:root@localhost:3306/tmall')
        df.to_sql(self.brand, engine, schema='tmall', if_exists='append')

        raw_df = pd.read_excel(self.brand+".xls")
        result_df = pd.concat([raw_df,df])
        print "df", df
        print "raw_df", raw_df
        print "result_df", result_df
        result_df.to_excel(self.brand+".xls",index=True)


def getJsonData(cfg_file):
    with open(cfg_file, "r") as f:
        str_data = f.read()
        configuration = json.loads(str_data)
        return configuration


def writeJsonData(cfg_file,context):
    with open(cfg_file, "w") as f:
        f.write(str(context))
        f.close()

if __name__ == "__main__":
    timestamp = getJsonData("timestamp.json")
    if timestamp["stamp"] == today_str:
        print "today`data had been downloaded"
    else:
        time_data = {'stamp': today_str}
        json_data = json.dumps(time_data)

        brands = ["anta", "lining", "skechers", "nike","adidas"]

        for brand in brands:
            inner_brand = TmallMonitor(brand)
            all_items = inner_brand.get_all_items()
            print "all_items",all_items
            items = inner_brand.purify_items(all_items)
            print "items", items
            inner_brand.insert_data(items)

        writeJsonData("timestamp.json", json_data)
        print "finish loading data"