import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from common.filepath_helper import FilePathHelper
from common.yaml_helper import file_config_dict
from common.excel_helper import data_config_dict as data_dict

class AutoTest01:


    def __init__(self , driver):

        self.driver = driver


        pass

    def login(self):
        print(data_dict)

        self.driver.get(data_dict['url'])
        self.waitElementDisplay(by=By.XPATH,value=data_dict['usernameXpath']).send_keys(data_dict['username'])
        self.waitElementDisplay(by=By.XPATH,value=data_dict['passwordXpath']).send_keys(data_dict['password'])
        self.waitElementDisplay(by=By.XPATH,value=data_dict['loginXpath']).click()



    def waitElementDisplay(self,  by , value ,isElements = False) :


        WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((by,value)))

        ele = self.driver
        if isElements:
            ele = self.driver.find_elements(by=by, value=value)
        else:
            ele = self.driver.find_element(by=by, value=value)
        return ele

if __name__ == '__main__':


    driverPath = os.path.join(FilePathHelper.get_project_path(), file_config_dict['input_driver_path'],"chromedriver.exe")
    service = Service(driverPath)
    driver = webdriver.Chrome(service=service)
    AutoTest01(driver).login()
    # print(driverPath)




    pass