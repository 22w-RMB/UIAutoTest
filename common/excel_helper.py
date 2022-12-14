import os.path

import xlwings
from common.filepath_helper import FilePathHelper
from common.yaml_helper import file_config_dict


class ExcelHepler:

    header = file_config_dict['execl_field']

    def __init__(self , filePath):

        self.app = xlwings.App(visible=False,add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        self.wb = self.app.books.open(filePath)

        pass

    def getAllData(self,sheetName, includeHeader=False):

        if sheetName not in self.getAllSheetNames():
            exit("没有" + sheetName + "表，请检查")

        ws = self.wb.sheets[sheetName]

        beginRow = 1 if includeHeader else 2
        data = []
        lastCell = ws.used_range.last_cell
        for r in range(beginRow,lastCell.row+1):
            rowList = []
            for c in range(1,lastCell.column+1):
                rowList.append(ws.range(r,c).value)
            data.append(rowList)

        return data

    def getConfigData(self,sheetName, includeHeader=False):

        if sheetName not in self.getAllSheetNames():
            exit("没有" + sheetName + "表，请检查")

        ws = self.wb.sheets[sheetName]

        beginRow = 1 if includeHeader else 2
        data = []
        lastCell = ws.used_range.last_cell
        for r in range(beginRow,lastCell.row+1):
            rowList = []
            for c in range(1,lastCell.column+1):
                rowList.append(ws.range(r,c).value)
            data.append(dict(zip(self.header,rowList)))

        return data


    def getConfigInfo(self , sheetName) -> dict:

        if sheetName not in self.getAllSheetNames():

            exit("没有"+sheetName+"表，请检查")

        d ={}

        sht = self.wb.sheets[sheetName]

        cell = sht.used_range.last_cell
        max_row = cell.row
        max_col = cell.column

        for i in range(1,max_row+1):

            d[sht.range(i,1).value] = sht.range(i,2).value

        return d




    def getAllSheetNames(self) -> list:

        l = []

        for sht in self.wb.sheets:

            l.append(sht.name)

        return l


    def close(self):
        self.wb.close()
        self.app.kill()




if __name__ == '__main__':


    excelPath = os.path.join(FilePathHelper.get_project_path(),file_config_dict['input_excel_path'],"配置信息.xlsx")
    e = ExcelHepler(excelPath)
    print(e.getConfigInfo('login'))
    print(e.getConfigData('公有数据管理'))

    e.close()