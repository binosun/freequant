
# coding: utf-8
from WindPy import *
from pandas import DataFrame

from freequant.thrustup.data_formater.data_grabber.common.config_reader import get_config
from freequant.thrustup.data_formater.data_grabber.common.date_counter import calculate_date


def get_wind_basic_data(stocks,flag="A"):
    config = get_config()

    date = calculate_date()
    if flag == "A":
        main_config = config["A_config"]
    else:
        main_config = config["H_config"]

    w.start()
    param_n = main_config["param_n"]
    delta_days = main_config["delta_days"]
    year_str = date["year_str"]
    params = "tradeDate=" + date["trade_date"].strftime("%Y%m%d") + ";priceAdj=F;cycle=D;n=" + param_n + \
             ";ruleType=9;rptDate=" + date["rpt_date_str"] + ";industryType=3;year=" + year_str


    result = w.wss(stocks, main_config["field_map"].keys(), params)
    codes = result.Codes
    fields = result.Fields
    data = result.Data
    field_map = {key.upper(): value for key, value in main_config["field_map"].items()}
    print "field_map",field_map
    raw_wind_df = DataFrame(data, index=fields, columns=codes).T
    raw_wind_df.rename(columns=field_map, inplace=True)
    return raw_wind_df