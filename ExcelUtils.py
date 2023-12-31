import win32com.client as win32
import os
import traceback
import pywintypes
import pandas as pd
from win32com.client import constants
import xlwings as xw


class ExcelUtils:

    def __init__(self, filePath):
        self.app = win32.DispatchEx('Excel.Application')
        # self.app = win32.gencache.EnsureDispatch('Excel.Application')
        self.app.Application.DisplayAlerts = False
        self.app.Application.ScreenUpdating = False
        self.app.Application.Visible = False
        # self.app.DefaultWebOptions.Encoding = win32.constants.xlUTF8

        self.filePath = filePath
        self.fileName, self.extension = os.path.splitext(os.path.basename(filePath))

    def ConvertXlsToXlsx(self):
        try:
            fileName = '%s.xlsx' % self.fileName
            dirName = os.path.dirname(self.filePath)
            newFilePath = os.path.join(dirName, fileName)
            wk = self.app.Workbooks.Open(self.filePath, False, False)
            wk.SaveAs(newFilePath, 51, ConflictResolution=2)
            # print(wk.Sheets(sheetName).Range("A1").Value)
            wk.Close()
            self.app.Quit()
        except Exception as e:
            self.app.Quit()
            strE = traceback.format_exc()
            raise Exception(strE)

    def ReadXlsToDf(self, sheetName):
        try:
            # fileName, extension = os.path.splitext(os.path.basename(filePath))
            fileName = '%s.xlsx' % self.fileName
            dirName = os.path.dirname(self.filePath)
            newFilePath = os.path.join(dirName, fileName)
            wk = self.app.Workbooks.Open(filePath, False, False)
            wk.Sheets(sheetName).UsedRange.Copy()
            wk.Close()
            self.app.Quit()
            df = pd.read_clipboard()
            return df
        except Exception as e:
            self.app.Quit()
            strE = traceback.format_exc()
            raise Exception(strE)

    def SaveXlsx(self, sheetName, dailyReportName):
        try:
            # 构造新file路径
            dirName = os.path.dirname(self.filePath)
            newFilePath = os.path.join(dirName, dailyReportName)

            # 打开excel
            wk = self.app.Workbooks.Open(self.filePath, False, False)
            wk.Sheets(sheetName).Activate()

            # 将剪切板数据paste到excel中
            # xlPasteValues = pywintypes.UnicodeType(-4163)
            # xlPasteSpecialOperationNone = pywintypes.UnicodeType(-4142)

            # xlPasteValues = constants.xlPasteValues
            # xlPasteSpecialOperationNone = constants.xlPasteSpecialOperationNone

            wk.Sheets(sheetName).Range("A2").Select()
            # wk.Sheets(sheetName).Range("A2").PasteSpecial('-4163', '-4142', False, False)
            # wk.Sheets(sheetName).Range("A2").PasteSpecial(xlPasteValues, xlPasteSpecialOperationNone, False, False)
            # wk.Sheets(sheetName).Range("A2").PasteSpecial(Format="Text", Link=False, DisplayAsIcon=False)
            wk.Sheets(sheetName).PasteSpecial(Format="Text", Link=False, DisplayAsIcon=False)

            # 删除header
            wk.Sheets(sheetName).Rows(2).Delete()

            # 触发pivot
            wk.RefreshAll()

            # 保存文件到daily report
            wk.SaveAs(newFilePath)

            wk.Close()
            self.app.Quit()
        except Exception as e:
            self.app.Quit()
            strE = traceback.format_exc()
            raise Exception(strE)

    def SaveXlsxWings(self, sheetName, dailyReportName, df):
        try:
            # 构造新file路径
            dirName = os.path.dirname(self.filePath)
            newFilePath = os.path.join(dirName, dailyReportName)

            app = xw.App(visible=False)
            wk = app.books.open(self.filePath)
            # wk.sheets[sheetName].range('A2').api.PasteSpecial()
            st = wk.sheets[sheetName]

            columnIndexList = ['K', 'O', 'AG']
            for i in columnIndexList:
                columnRange = st.range('%s:%s' % (i, i))
                columnRange.number_format = '@'
            # # 获取AG列的范围
            # columnRange = st.range('AG:AG')

            # # 设置列的单元格格式为文本（字符串）
            # columnRange.number_format = '@'

            # # 获取E、F、G列的范围
            # dateColumnRange = st.range('E:E,F:F,G:G')

            # # 设置列的单元格格式为日期格式
            # dateColumnRange.number_format = 'yyyy/MM/dd'


            st.range('A2').value = df.values
            
            wk.api.RefreshAll()
            wk.save(newFilePath)
            wk.close()
            app.quit()
        except Exception as e:
            app.quit()
            strE = traceback.format_exc()
            raise Exception(strE)


if __name__ == '__main__':

    filePath = r'C:\Users\tang.k.5\OneDrive - Procter and Gamble\Desktop\Code Projects\SAMBCOptimization\Input Files\ZOCR.xls'
    obj = ExcelUtils(filePath)
    obj.ConvertXlsToXlsx()
    # ExcelUtils.ConvertXlsToXlsx(filePath)
