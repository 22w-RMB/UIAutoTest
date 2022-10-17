

import pandas as pd
import xlwings
import xlwings as xw
xlwings.Sheet

class excelHandler:


    def __init__(self, filePath):


        self.filePath = filePath
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        self.wb = self.app.books.open(filePath)
        print(self.wb.sheets)




    def getCellValue(self, sheetName, row, col):
        ws = self.getSheet(sheetName)
        return ws.range(row,col).value

    def readCell(self):
        ws = self.getSheet("配置信息")

        # 通过行、列获取单元格的值，数据默认是浮点数
        print(ws.range(2, 1).value)  # 登录页面
        # 通过 单元格引用读取值 ，大小写都可以
        print(ws.range("a2").value)  # 登录页面
        print(ws.range("A2").value)  # 登录页面
        # 读取一段区间内的值
        print(ws.range("a1:b3").options(ndim=2).value)
        # 结果：[['key', 1231.0], ['登录页面', 'https://adssx-test-gzdevops3.tsintergy.com/usercenter/#/login'], ['账号', 'zhanzw']]

        print(ws.range((1, 1), (3, 2)).options(ndim=2).value)
        # 结果：[['key', 1231.0], ['登录页面', 'https://adssx-test-gzdevops3.tsintergy.com/usercenter/#/login'], ['账号', 'zhanzw']]

        # 赋值
        ws.range(1, 2).value = "!213"

        # 删除、插入行
        ws.range('a3').api.EntireRow.Delete()   # 会删除 ’a3‘ 单元格所在行
        ws.api.Rows(4).Insert() # 会在第 4 行插入一行，原来的第4行下移

        # 删除、插入列
        ws.range('c2').api.EntireColumn.Delete() # 会删除 ’c2‘ 单元格所在列
        ws.api.Columns(3).Insert()  # 会在第 3 列插入一列，原来的第 3 列右移

        # 选择 Sheet 页面最右下角的单元格（无论该单元格有无数据），即获取整个sheet的最大行数、列数
        cell = ws.used_range.last_cell
        rows = cell.row
        columns = cell.column
        print(rows, "  ", columns)

        # 这种方法只计算连续单元格，如果遇到空单元格则停止，即 excel 的 ctrl+shift+down 按钮的逻辑
        cell1 = ws.range('a1').expand('down')
        max_row = cell1.rows.count
        print(max_row)

        # 合并单元格
        ws.range('c5:d6').api.Merge()   # 合并单元格
        # ws.range('c5:d6').api.UnMerge()  # 拆分单元格
        self.saveFile()


        # shape = ws.used_range.shape
        # max_row = shape[0]
        # max_col = shape[1]
        # print(ws.used_range.shape)  # (6, 2)
        pass


    def getSheet(self, sheetName) :

        # 获取 sheet
        # 通过表名获取 sheet
        ws = self.wb.sheets['Sheet1']
        # 通过索引索取 sheet
        ws = self.wb.sheets[0]
        # 获取当前活动的 sheet
        ws = self.wb.sheets.active


        for i in range(0,len(self.wb.sheets)):

            if sheetName == self.wb.sheets[i].name:
                return  self.wb.sheets[i]



        exit("找不到表名，程序终止")


    def saveFile(self, savePath = None):

        if savePath is None:
            savePath = self.filePath

        self.wb.save(savePath)


    def close(self):
        self.wb.close()
        self.app.kill()




if __name__ == '__main__':


    path = r'D:\code\python\UIAutoTest\input\excel\配置信息.xlsx'


    e = excelHandler(path)
    try:
        e.getAllDatas('配置信息')
    finally:

        e.close()

    pass


