import pandas as pd
import time
import os
from YamlHandler import YamlHandler
from PandasUtils import PandasUtils
from ExcelUtils import ExcelUtils





class SAMBCOptimization:
    def __init__(self) -> None:
        self.Initialize()

        self.mainFolder = '%s\\Input Files'%os.getcwd() if self.isTestMode else os.getcwd()
       
    def Initialize(self):
        self.settings = YamlHandler(os.path.join(
            os.getcwd(), 'config.yaml')).ReadYaml()
        
        self.isTestMode = self.settings['isTestMode']
        self.zderRevisedColumns = self.settings['zderRevisedColumns']
        self.zderzderReservedColumnList = self.zderRevisedColumns.keys()
        self.finalReportFieldList = self.settings['finalReportFieldList']
        pass
        
        

    def ReadFilesToDataFrame(self):

        # 构造dfMain
        self.dfMain = pd.DataFrame(columns=self.finalReportFieldList)

        # 读取dfZDER
        self.dfZDER = PandasUtils.GetDataFrame(self.mainFolder,'ZDER.xlsx', 'Sheet1')        

        # 将ZOCR.xls转成ZOCR.xlsx，并且读取dfZOCR
        ExcelUtils().ConvertXlsToXlsx(os.path.join(self.mainFolder, 'ZOCR.xls'))
        self.dfZOCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZOCR.xlsx', 'ZOCR')
        
        # 读取dfZCCR
        self.dfZCCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZCCR.xlsx', 'Sheet1')

    #------------------------------ZDER的操作------------------------------#

    def DeleteUselessColumnsInDfZder(self):
        self.dfZDER = PandasUtils.DeleteColumns(self.dfZDER, self.zderzderReservedColumnList)

    def RenameFieldNameForDfZder(self):
        self.dfZDER.rename(columns=self.zderRevisedColumns, inplace=True)

    def AppendColumnsToDfZder(self):
        appendColumnList = []

        # 如果columnName不在finalReportFieldList中，则需要将columnName加到dfZDER中
        for columnName in self.finalReportFieldList:
            if not columnName in self.zderzderReservedColumnList:
                appendColumnList.append(columnName)

        self.dfZDER[appendColumnList] = ""

    def SortDfZderFieldsByFinalReportFields(self):
        # 按主表将field进行排序
        self.dfZDER.reindex(self.finalReportFieldList)
        
    def AppendDfZderToDfMain(self):
        self.DeleteUselessColumnsInDfZder()
        self.RenameFieldNameForDfZder()
        self.SortDfZderFieldsByFinalReportFields()
        # self.dfMain = self.dfMain.fillna(self.dfZDER)
        self.dfMain = pd.concat([self.dfMain, self.dfZDER])
        print(self.dfMain)

    #------------------------------ZDER的操作------------------------------#


    #------------------------------ZOCR的操作------------------------------#
    def InsertZocrOrderValueToDfMain(self):
        self.dfZOCR['Sales Order No.'] = '0' + self.dfZOCR['Sales Order No.']
        self.dfZOCR['Material No.'] = '0000000000' + self.dfZOCR['Material No.']
        # self.dfZOCR = self.dfZOCR.dropna(subset=['Material No.'])

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfZOCR, ['宝洁订单号', '宝洁产品代码'], ['Sales Order No.', 'Material No.'], ['下单数量'], ['Order Value'])

        # self.dfMain.loc[(self.dfMain['宝洁订单号'] == self.dfZOCR['Sales Order No.']) & (self.dfMain['宝洁产品代码'] == self.dfZOCR['Material No.']), '下单数量'] = self.dfZOCR['Order Value']
        # self.dfMain = pd.merge(self.dfMain, self.dfZOCR, left_on=['宝洁订单号', '宝洁产品代码'], right_on=['Sales Order No.', 'Material No.'])

    #------------------------------ZOCR的操作------------------------------#


    #------------------------------ZCCR的操作------------------------------#

    def CalculateNotSatisfiedQty(self):
        # 将字符串列转换为数值类型
        self.dfMain['未满足数量'] = pd.to_numeric(self.dfMain['未满足数量'])
        self.dfMain['下单数量'] = pd.to_numeric(self.dfMain['下单数量'])
        self.dfMain['分货数量'] = pd.to_numeric(self.dfMain['分货数量'])

        # 计算未满足数量
        self.dfMain['未满足数量'] = self.dfMain['下单数量'] - self.dfMain['分货数量']
        pass

    def WriteZccrCutReasonToDfMain(self):
        if self.dfZCCR is None:
            return

        self.CalculateNotSatisfiedQty()

        
    



    #------------------------------ZCCR的操作------------------------------#


    def Main(self):
        # filePath = r'C:\Users\tang.k.5\OneDrive - Procter and Gamble\Desktop\Code Projects\SAMBCOptimization\Input Files\ZDER.xlsx'
        # timeStart = time.time()
        # df = pd.read_excel(filePath, sheet_name="Sheet1", dtype='str')
        # print("读数据的时间为%ss"%(time.time() - timeStart))
        self.ReadFilesToDataFrame()
        self.AppendDfZderToDfMain()
        self.InsertZocrOrderValueToDfMain()
        self.WriteZccrCutReasonToDfMain()
        pass
        # print(df)




if __name__ == '__main__':
    obj = SAMBCOptimization()
    obj.Main()