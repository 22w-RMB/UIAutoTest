import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains, Keys
from common.filepath_helper import FilePathHelper
from common.yaml_helper import file_config_dict
from common.excel_helper import data_config_dict as data_dict
from common.excel_helper import data_list

class AutoTest01:


    def __init__(self , driver):

        self.driver = driver


        pass

    # def login(self):
    #     print(data_dict)
    #
    #     self.driver.get(data_dict['url'])
    #     self.waitElementDisplay(by=By.XPATH,value=data_dict['usernameXpath']).send_keys(data_dict['username'])
    #     self.waitElementDisplay(by=By.XPATH,value=data_dict['passwordXpath']).send_keys(data_dict['password'])
    #     self.waitElementDisplay(by=By.XPATH,value=data_dict['loginXpath']).click()
    #     # self.waitElementDisplay(by=By.,value=data_dict['loginXpath']).click()

    def login(self):
        print(data_dict)

        self.driver.get(data_dict['url'])

        usernameXpath = r'//*[@id="username"]'
        passwordXpath = r'//*[@id="password"]'
        loginXpath = r'//*[@id="root"]/section/main/div/div[2]/div/form/div/div[3]/div/div/div/button'


        self.waitElementDisplay(by=By.XPATH,value=usernameXpath).send_keys(data_dict['username'])
        self.waitElementDisplay(by=By.XPATH,value=passwordXpath).send_keys(data_dict['password'])
        self.waitElementDisplay(by=By.XPATH,value=loginXpath).click()

        selectXpath = r'//*[@id="root"]/section/header/div/div[2]/div/div[1]/a'
        # 非select 下拉框
        menuXpath = r'/html/body/div[3]/div/div/ul/li[text()="'+data_dict['enterprise'] + '"]'
        # 应用
        appXpath = r'//*[@id="row_item_0"]/div[@title="'+data_dict['application'] + '"]'

        selectEle = self.waitElementDisplay(by=By.XPATH, value=selectXpath)
        ActionChains(self.driver).move_to_element(selectEle).perform()
        self.waitElementDisplay(by=By.XPATH, value=menuXpath).click()
        self.waitElementDisplay(by=By.XPATH, value=appXpath).click()


    def elementOperator(self, isElements, operator, by, value ,key):

        if operator == 'click':
            self.waitElementDisplay(by,value,isElements).click()
        elif operator == 'input':
            e = self.waitElementDisplay(by, value, isElements)
            e.send_keys(Keys.CONTROL+"a")
            e.send_keys(key)
        elif operator == 'table':
            trs = self.waitElementDisplay(by, value, isElements)
            for tr in trs:
                # print(t.text)
                # print("====")
                tds = tr.find_elements(by=By.TAG_NAME, value='td')
                for td in tds:
                    print(td.text)
        else:
            exit('类型错误')

    def waitElementDisplay(self,  by , value ,isElements = False) :



        ele = self.driver
        if isElements:
            WebDriverWait(self.driver,10).until(EC.presence_of_element_located((by, value)))
            ele = self.driver.find_elements(by=by, value=value)
        else:
            WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((by,value)))
            ele = self.driver.find_element(by=by, value=value)
        return ele

if __name__ == '__main__':

    options = webdriver.ChromeOptions()
    options.add_experimental_option('detach', True)  # 不自动关闭浏览器
    driverPath = os.path.join(FilePathHelper.get_project_path(), file_config_dict['input_driver_path'],"chromedriver.exe")
    service = Service(driverPath)
    driver = webdriver.Chrome(service=service, options=options)
    # driver.get("https://www.hao123.com/")
    a = AutoTest01(driver)
    a.login()

    # for i in range(0,len(data_list)):
    #     a.elementOperator(data_list[i]['isElements'],data_list[i]['operator']
    #                       ,data_list[i]['by'],data_list[i]['value'],data_list[i]['key'])

    a.elementOperator(isElements=True,operator='table',by=By.XPATH,
                      value='//*[@id="root"]/section/section/section/main/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr'
                      , key = None)

    # print(driverPath)




    pass