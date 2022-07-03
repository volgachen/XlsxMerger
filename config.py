from glob import glob
import os
import yaml
import copy
from glob import glob
from collections import OrderedDict



def get_config(cfg_file):
    with open(cfg_file, 'r', encoding='utf-8') as f:
        yaml_cfg = yaml.load(f, Loader=yaml.FullLoader)
    for item in yaml_cfg["Sheets"]:
        if isinstance(item["posfix_string"], str):
            item["posfix_string"] = item["posfix_string"].replace(" ", "")

    if "Path" not in yaml_cfg or yaml_cfg["Path"] == "None":
        assert "Files" in yaml_cfg, "Path项与Files项至少存在一个！"
        return yaml_cfg
    
    yaml_cfg["Files"] = glob(yaml_cfg["Path"])
    return yaml_cfg
