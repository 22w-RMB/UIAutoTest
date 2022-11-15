import os.path

import yaml
from common.filepath_helper import FilePathHelper

config_path = os.path.join(FilePathHelper.get_project_path(),"conf","config.yaml")

# print(config_path)

file_config_dict : dict

with open(config_path,'r',encoding="utf-8") as f:
    file_config_dict = yaml.safe_load(f)
    # print(config_dict)

if __name__ == '__main__':
    print(file_config_dict)