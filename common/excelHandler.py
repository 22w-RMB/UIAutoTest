

import pandas as pd
import xlwings
import xlwings as xw


class excelHandler:


    def __init__(self, filePath):


        self.filePath = filePath
        # visible：表示处理过程是否可见 ， add_book：表示是否打开新的Excel程序
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False   # 关闭一些提示信息，可以加快运行信息。默认为True
        self.app.screen_updating = False  # 更新显示工作表的内容。默认为True。关闭它也可以提高运行速度
        # 打开工作簿
        self.wb = self.app.books.open(filePath)
        # 打开工作簿
        # self.wb = xw.Book(filePath)
        # 创建新的工作簿
        # self.wb = self.app.books.add()
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

        # 将列表[1,2,3]储存在A1：C1中
        ws.range('A1').value = [1, 2, 3]
        # 将列表[1,2,3]储存在A1:A3中
        ws.range('A1').options(transpose=True).value = [1, 2, 3]
        # 将2x2表格，即二维数组，储存在A1:B2中，如第一行1，2，第二行3，4
        ws.range('A1').options(expand='table').value = [[1, 2], [3, 4]]

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
        ws = self.wb.sheets(1)
        # 获取表名
        print(ws.name ,"==================")
        # # 修改表名
        # ws.name = "1234"
        # # 获取当前活动的 sheet
        # ws = self.wb.sheets.active
        # # 追加新的sheet
        # sht1 = self.wb.sheets.add()
        # sht2 = self.wb.sheets.add()
        # # 获取sheet的个数
        # print(self.wb.sheets.count)
        # # 清除工作表所有格式
        # ws.clear()
        # # 清除工作表的所有内容但是保留原有格式
        # ws.clear_contents()
        # # 删除工作表
        # # ws.delete()
        # # 自动调整行高列宽
        # ws.autofit('c')

        '''
        将指定工作表复制到工作簿的另一位置。
        expression.Copy(Before, After)
        expression      必需。该表达式返回上面的对象之一。
        Before      Variant 类型，可选。指定某工作表，复制的工作表将置于此工作表之前。如果已经指定了 After，则不能指定 Before。
        After      Variant 类型，可选。指定某工作表，复制的工作表将置于此工作表之后。如果已经指定了 Before，则不能指定 After。
        说明
        如果既未指定 Before 参数也未指定 After 参数，则 Microsoft Excel 将新建一个工作簿，其中将包含复制的工作表。
        本示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。
        Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")
        '''

        # # 复制工作表到同一工作表后面
        # ws.api.Copy(After=ws.api)
        # # 将ws 工作表复制到最后一个工作表后面
        # ws2 = self.wb.sheets[-1]
        # ws.api.Copy(After=ws2.api)

        # # 将 ws 工作表复制到另外一个工作簿的某个工作表后面
        # wb_other = xw.Book('1.xlsx')
        # wx_other = wb_other.sheets[-1]
        # ws.api.Copy(After=wx_other.api)
        #
        for i in range(0,len(self.wb.sheets)):

            if sheetName == self.wb.sheets[i].name:
                return  self.wb.sheets[i]



        exit("找不到表名，程序终止")


    def saveFile(self, savePath = None):

        if savePath is None:
            savePath = self.filePath

        # 保存
        self.wb.save(savePath)


    def close(self):
        # 关闭 & 杀死进程
        self.wb.close()
        self.app.kill()
        # 退出excel程序,不保存任何工作簿
        # self.app.quit()



if __name__ == '__main__':


    path = r'/input/excel/配置信息.xlsx'


    e = excelHandler(path)
    try:
        e.getSheet('配置信息')
        e.saveFile(r'D:\code\python\UIAutoTest\input\excel\配置信息1.xlsx')
    finally:

        e.close()

    pass


