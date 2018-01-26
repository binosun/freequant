# coding: utf-8


import json

import pandas as pd




with open("PTA_price.json") as price_json:
    price_json = json.loads(price_json.read())
    # Date,Open,High,Low,Close,Vol
prices = price_json["PTA_price"]

pta_df = pd.DataFrame(prices,columns=["Data", "Open","High","Low","Close","Vol"])
# pta_df.to_excel("pta_prices.xls")



with open("601233.SH.xls") as price_json:
    price_json = json.loads(price_json.read())
    # Date,Open,High,Low,Close,Vol
prices = price_json["PTA_price"]

pta_df = pd.DataFrame(prices,columns=["Data", "Open","High","Low","Close","Vol"])







