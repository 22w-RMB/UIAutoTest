import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains, Keys
from common.filepath_helper import FilePathHelper
from common.yaml_helper import file_config_dict
from common.excel_helper import ExcelHepler

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


    def elementOperator(self, isElements, operator, by, value ,key=None,fileName=None,sheetName=None):

        if operator == 'click':
            e = self.waitElementDisplay(by,value,isElements)
            print(e.text)
            e.click()
        elif operator == 'input':
            e = self.waitElementDisplay(by, value, isElements)
            e.send_keys(Keys.CONTROL+"a")
            e.send_keys(key)
        elif operator == 'table':
            self.tableAssert(isElements, by, value,fileName,sheetName)
        elif operator == 'date':
            e = self.waitElementDisplay(by, value, isElements)
            e.send_keys(Keys.CONTROL+"a")
            e.send_keys(key)
            e.send_keys(Keys.ENTER)
        else:
            exit('类型错误')


    def tableAssert(self,isElements, by, value,fileName=None,sheetName=None):
        tableData = self.getTableData(isElements, by, value)
        fileData = self.getFileData(fileName,sheetName)
        self.listCompare(tableData, fileData)

    def getTableData(self,isElements, by, value):

        trs = self.waitElementDisplay(by, value, isElements)
        tableData = []
        i = 0
        for tr in trs:
            if i == 0:
                i += 1
                continue

            t = []
            tds = tr.find_elements(by=By.TAG_NAME, value='td')
            for td in tds:
                t.append(td.text)
                # print(td.text)
            tableData.append(t)

        # print("表格数据",tableData)
        return tableData

    def listCompare(self,tableData,fileData):



        if type(tableData) == str or type(tableData) == int:

            isFloat = False
            if "." in tableData:
                s = tableData.split(".")
                if len(s) == 2:
                    if s[0].isdigit() and s[1].isdigit() :
                        isFloat = True

            if isFloat:
                tableData = float(tableData)

            if tableData != fileData:
                print("数据不一样")
            else:
                print("数据一样")
            return

        if type(tableData) != type(fileData):
            print(tableData,fileData)
            print(type(tableData) ,type(fileData))
            print("类型不一样")
            return

        if type(tableData) != list:
            print("类型不是list")

        tl = len(tableData)
        fl = len(fileData)
        if tl!=fl:
            print("长度不一样")
            return

        for i in range(0,tl):
            self.listCompare(tableData[i],fileData[i])



    def getFileData(self,fileName,sheetName):

        # print(self.getFilePath(fileName))
        eh = ExcelHepler(self.getFilePath(fileName))
        d : list
        try:
            d = eh.getAllData(sheetName)
        finally:
            eh.close()
        # print("文件数据",d)
        return d

    def getFilePath(self,fileName):
        path = os.path.join(FilePathHelper.get_project_path(), file_config_dict['input_excel_path'])

        for root,dirs,files in os.walk(path):
            for file in files:
                # print(file)
                if fileName in file:
                    # print(root, file)
                    return os.path.join(root,file)



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

    # 获取配置信息
    excelPath = os.path.join(FilePathHelper.get_project_path(), file_config_dict['input_excel_path'], "配置信息.xlsx")
    excel_helper = ExcelHepler(excelPath)
    data_dict: dict
    data_list: list
    try:
        data_dict = excel_helper.getConfigInfo('login')
        data_list = excel_helper.getConfigData('公有数据管理')
    finally:
        excel_helper.close()

    options = webdriver.ChromeOptions()

    options.add_experimental_option('detach', True)  # 不自动关闭浏览器

    downloadPath = os.path.join(FilePathHelper.get_project_path(),file_config_dict["output_file_path"])
    print(downloadPath)
    prefs = {
        'download.default_directory': downloadPath,     # 设置下载路径，路径不存在会自动创建
        'download.prompt_for_download': False,           # 是否弹窗询问
        'safebrowsing.enabled': False,                   # 是否提示安全警告
        # Boolean that records if the download directory was changed by an
        # upgrade a unsafe location to a safe location.
    }

    options.add_experimental_option("prefs",prefs)


    driverPath = os.path.join(FilePathHelper.get_project_path(), file_config_dict['input_driver_path'],"chromedriver.exe")
    service = Service(driverPath)
    driver = webdriver.Chrome(service=service, options=options)
    # driver.get("https://www.hao123.com/")
    a = AutoTest01(driver)
    a.login()
    for i in range(0,len(data_list)):
        a.elementOperator(data_list[i]['isElements'],data_list[i]['operator']
                          ,data_list[i]['by'],data_list[i]['value'],data_list[i]['key']
                          ,data_list[i]['fileName'],data_list[i]['sheetName'])

    # a.elementOperator(isElements=True,operator='table',by=By.XPATH,
    #                   value='//*[@id="root"]/section/section/section/main/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr'
    #                   , key = None)

    # a.getExcelData('全网统一出清价格','2022-02-03')




    pass