# import win32com.client as win32
import os
import traceback

import pandas as pd


class ExcelUtils:

    def __init__(self):
        self.app = win32.DispatchEx('Excel.Application')
        self.app.Application.DisplayAlerts = False
        self.app.Application.ScreenUpdating = False
        self.app.Application.Visible = False

    def ConvertXlsToXlsx(self, filePath):
        try:
            fileName, extension = os.path.splitext(os.path.basename(filePath))
            fileName = '%s.xlsx' % fileName
            dirName = os.path.dirname(filePath)
            newFilePath = os.path.join(dirName, fileName)
            wk = self.app.Workbooks.Open(filePath, False, False)
            wk.SaveAs(newFilePath, 51, ConflictResolution=2)
            # print(wk.Sheets(sheetName).Range("A1").Value)
            wk.Close()
            self.app.Quit()
        except Exception as e:
            self.app.Quit()
            strE = traceback.format_exc()
            raise Exception(strE)

    def ReadXlsToDf(self, filePath, sheetName):
        try:
            fileName, extension = os.path.splitext(os.path.basename(filePath))
            fileName = '%s.xlsx' % fileName
            dirName = os.path.dirname(filePath)
            newFilePath = os.path.join(dirName, fileName)
            wk = self.app.Workbooks.Open(filePath, False, False)
            wk.Sheet(sheetName).UsedRange.Copy()
            wk.Close()
            self.app.Quit()
            df = pd.read_clipboard()
            return df
        except Exception as e:
            self.app.Quit()
            strE = traceback.format_exc()
            raise Exception(strE)


if __name__ == '__main__':
    obj = ExcelUtils()
    filePath = r'C:\Users\tang.k.5\OneDrive - Procter and Gamble\Desktop\Code Projects\SAMBCOptimization\Input Files\ZOCR.xls'
    obj.ConvertXlsToXlsx(filePath)
    # ExcelUtils.ConvertXlsToXlsx(filePath)
