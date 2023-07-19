import numpy as np
import pandas as pd
import time
import os
from YamlHandler import YamlHandler
from PandasUtils import PandasUtils
from ExcelUtils import ExcelUtils


class SAMBCOptimization:
    def __init__(self) -> None:
        self.Initialize()

        self.mainFolder = '%s\\Input Files' % os.getcwd() if self.isTestMode else os.getcwd()

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
        self.dfZDER = PandasUtils.GetDataFrame(self.mainFolder, 'ZDER.xlsx', 'Sheet1')

        # 将ZOCR.xls转成ZOCR.xlsx，并且读取dfZOCR
        ExcelUtils().ConvertXlsToXlsx(os.path.join(self.mainFolder, 'ZOCR.xls'))
        self.dfZOCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZOCR.xlsx', 'ZOCR')

        # 读取dfZCCR
        self.dfZCCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZCCR.xlsx', 'Sheet1')

    # ------------------------------ZDER的操作Start------------------------------#

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

    # ------------------------------ZDER的操作End------------------------------#

    # ------------------------------ZOCR的操作Start------------------------------#
    def InsertZocrOrderValueToDfMain(self):
        self.dfZOCR['Sales Order No.'] = '0' + self.dfZOCR['Sales Order No.']
        self.dfZOCR['Material No.'] = '0000000000' + self.dfZOCR['Material No.']
        # self.dfZOCR = self.dfZOCR.dropna(subset=['Material No.'])

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfZOCR, ['宝洁订单号', '宝洁产品代码'],
                                            ['Sales Order No.', 'Material No.'], ['下单数量'], ['Order Value'])

        # self.dfMain.loc[(self.dfMain['宝洁订单号'] == self.dfZOCR['Sales Order No.']) & (self.dfMain['宝洁产品代码'] == self.dfZOCR['Material No.']), '下单数量'] = self.dfZOCR['Order Value']
        # self.dfMain = pd.merge(self.dfMain, self.dfZOCR, left_on=['宝洁订单号', '宝洁产品代码'], right_on=['Sales Order No.', 'Material No.'])

    # ------------------------------ZOCR的操作End------------------------------#

    # ------------------------------ZCCR的操作Start------------------------------#

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

        dfUnsatisfiedQty = self.dfMain.loc[self.dfMain['未满足数量'] == np.NaN or self.dfMain['未满足数量'] == '']

        dfUnsatisfiedQty.drop_duplicates(subset=['宝洁订单号', '宝洁产品代码'], keep='first')

        for index, row in dfUnsatisfiedQty.iterrows():
            strSalesOrder = row['宝洁订单号'].lstrip('0')
            strMaterial = row['宝洁产品代码'].lstrip('0')
            dfZCCRCutReason = self.dfZCCR.loc[
                self.dfZCCR['Sales Order'] == strSalesOrder & self.dfZCCR['Material'] == strMaterial]

            # 没有找到记录，继续下一条
            if dfZCCRCutReason.shape[0] == 0:
                continue
            # 找到只有一条记录，赋值，继续下一条
            if dfZCCRCutReason.shape[0] == 1:
                row['未满足原因代码'] = dfZCCRCutReason.at[0, 'Rsn. Code']
                continue


            # ---------------------找到有多条记录Start---------------------#

            # 筛选出cut reason为59的记录
            dfCutReason59 = dfZCCRCutReason.loc[dfZCCRCutReason['Rsn. Code'] == '59']

            # 如果存在59的记录，则赋值59，继续下一条
            if dfCutReason59.shape[0] > 0:
                row['未满足原因代码'] = '59'
                continue

            # --------------不存在59 Start--------------#

            # 如果cut reason只有两行
            if dfZCCRCutReason.shape[0] == 2:
                # 两行分别是01和58, 则赋值58，继续下一条
                if '01' in dfZCCRCutReason['Rsn. Code'].values and '58' in dfZCCRCutReason['Rsn. Code'].values:
                    row['未满足原因代码'] = '58'
                    continue
            # 两行不是01和58，有多行, 按reason code group by
            dfGrouped = dfZCCRCutReason.groupby('Rsn. Code')

            for groupName, dfGroup in dfGrouped:
                # 有同一个reason code有同样数量的正负值
                if any((dfGroup['Cut Quantity'] > 0) & (dfGroup['Cut Quantity'].abs().duplicated())):

                    # 判断dfZCCRCutReason是否包含reason code为D4的记录

                    pass

            # --------------不存在59 End--------------#


            # ---------------------找到有多条记录End---------------------#






    # ------------------------------ZCCR的操作End------------------------------#

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
