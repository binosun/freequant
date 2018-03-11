
# coding: utf-8

import json
from collections import OrderedDict

def get_config():
    with open("config/config.json", "r") as f:
        configuration = json.load(f, object_pairs_hook=OrderedDict)
        return configuration
