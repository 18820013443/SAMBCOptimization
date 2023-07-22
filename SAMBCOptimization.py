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
        self.dfMain = None


        self.mainFolder = '%s\\Input Files' % os.getcwd() if self.isTestMode else os.getcwd()

    def Initialize(self):
        self.settings = YamlHandler(os.path.join(
            os.getcwd(), 'config.yaml')).ReadYaml()

        self.isTestMode = self.settings['isTestMode']
        self.zderRevisedColumns = self.settings['zderRevisedColumns']
        self.zderReservedColumnList = self.zderRevisedColumns.keys()
        self.finalReportFieldList = self.settings['finalReportFieldList']
        self.zeerScreenConditionList = self.settings['zeerScreenConditionList']
        self.zeerReservedColumns = self.settings['zeerReservedColumns']
        self.zeerReservedColumnsList = self.zeerReservedColumns.keys()
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

        # 读取dfVBAK
        # self.dfVBAK = PandasUtils.GetDataFrame(self.mainFolder, 'VBAK.xlsx', 'Sheet1')

        # 读取dfVBAP
        self.dfVBAP = PandasUtils.GetDataFrame(self.mainFolder, 'VBAP.xlsx', 'VBAP')

        # 读取dfLIPS
        self.dfLIPS = PandasUtils.GetDataFrame(self.mainFolder, 'LIPS.xlsx', 'Sheet1')

        # 读取dfOpenAlloment
        self.dfOpenAllotment = PandasUtils.GetDataFrame(self.mainFolder, 'SAMBC Open Allotment List.xlsx', 'Sheet1')

        # 读取dfZEER
        self.dfZEER = PandasUtils.GetDataFrame(self.mainFolder, 'ZEER.xlsx', 'Sheet1')

        # 读取dfCustomerList
        self.dfCustomerList = PandasUtils.GetDataFrame(self.mainFolder, 'SAMBC Customer List.xlsx', 'Sheet1')

        # 读取dfPriceList
        self.dfPriceList = PandasUtils.GetDataFrame(self.mainFolder, 'SAMBC Price List.xlsx', 'Sheet1')

        # 读取dfCutReasonList
        self.dfCutReasonList = PandasUtils.GetDataFrame(self.mainFolder, 'SAMBC Parameters.xlsx', 'Cut Reason List')

    # ------------------------------ZDER的操作Start------------------------------#

    def DeleteUselessColumnsInDfZder(self):
        self.dfZDER = PandasUtils.DeleteColumns(self.dfZDER, self.zderReservedColumnList)

    def RenameFieldNameForDfZder(self):
        self.dfZDER.rename(columns=self.zderRevisedColumns, inplace=True)

    def AppendColumnsToDfZder(self):
        appendColumnList = []

        # 如果columnName不在finalReportFieldList中，则需要将columnName加到dfZDER中
        for columnName in self.finalReportFieldList:
            if not columnName in self.zderReservedColumnList:
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

    def OnlyD4And07Operations(self, dfGrouped, dfZCCRCutReason, row):
        isOnlyD407Contained = PandasUtils.isOnlyD407Contained(dfZCCRCutReason)

        hasNegativeQuantityForD4 = PandasUtils.HasNegativeQuantityForD4(dfZCCRCutReason)

        isZeroSumD4And07 = PandasUtils.IsZeroSumD4And07(dfZCCRCutReason)

        # reason code只包含D4和07
        if isOnlyD407Contained:
            # reason code为D4的Cut Quantity有小于0的记录
            if hasNegativeQuantityForD4:
                row['未满足原因代码'] = '01'
            else:
                row['未满足原因代码'] = '07'
        # reason code不只包含D4和07
        else:
            if isZeroSumD4And07:
                self.Contains04Operations(True, dfGrouped, dfZCCRCutReason, row)
            else:
                if hasNegativeQuantityForD4:
                    row['未满足原因代码'] = '01'
                else:
                    row['未满足原因代码'] = '07'

    def Contains04Operations(self, shouldRemoveSameAbsValue, dfGrouped, dfZCCRCutReason, row):
        if shouldRemoveSameAbsValue:
            # 删除有同样数量和同样reason code的正负值的记录
            dfZCCRCutReasonRemoveValues = PandasUtils.DeleteDuplicatedAbsRows(dfGrouped)
            # 相加reason code相同的'Cut Quantity'的正负值
            dfZCCRCutReasonNew = PandasUtils.SumQtyForSameReasonCode(dfZCCRCutReasonRemoveValues)
        else:
            # 相加reason code相同的'Cut Quantity'的正负值
            dfZCCRCutReasonNew = PandasUtils.SumQtyForSameReasonCode(dfZCCRCutReason)
        # 判断‘Cut Quantity’ 绝对值的最大值是否唯一
        isMaxAbsValueUnique, maxAbsValue = PandasUtils.IsMaxAbsValueUnique(dfZCCRCutReasonNew)

        reasonCode = PandasUtils.FindReasonCodeForMaxAbsValue(maxAbsValue, dfZCCRCutReasonNew)

        # 绝对值的最大值是唯一
        if isMaxAbsValueUnique:
            # reasonCode = PandasUtils.FindReasonCodeForMaxAbsValue(maxAbsValue, dfZCCRCutReasonNew)
            if reasonCode == '01':
                row['未满足原因代码'] = '01'
            else:
                row['未满足原因代码'] = '07'
        # 绝对值的最大值不是是唯一
        else:
            # reasonCode = PandasUtils.FindReasonCodeForMaxAbsValue(maxAbsValue, dfZCCRCutReasonNew)
            row['未满足原因代码'] = reasonCode

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

            hasSamePositiveAndNegativeValueForOneReason = PandasUtils.HasSamePositiveAndNegativeValueForOneReason(
                dfGrouped)

            isD4Contained = PandasUtils.ReasonCodeContainsD4(dfZCCRCutReason)

            # isOnlyD407Contained = PandasUtils.isOnlyD407Contained(dfZCCRCutReason)
            #
            # hasNegativeQuantityForD4 = PandasUtils.HasNegativeQuantityForD4(dfZCCRCutReason)
            #
            # isZeroSumD4And07 = PandasUtils.IsZeroSumD4And07(dfZCCRCutReason)

            # 同一个reason code有同样数量的正负值
            if hasSamePositiveAndNegativeValueForOneReason:
                # reason code是包含D4
                if isD4Contained:
                    self.OnlyD4And07Operations(dfGrouped, dfZCCRCutReason, row)

                # reason code不包含D4
                else:
                    # reason code是否包含D4
                    if isD4Contained:
                        self.Contains04Operations(True, dfGrouped, dfZCCRCutReason, row)


            # 同一个reason code有不同数量的正负值
            else:
                if isD4Contained:
                    self.OnlyD4And07Operations(dfGrouped, dfZCCRCutReason, row)
                else:
                    self.Contains04Operations(False, dfGrouped, dfZCCRCutReason, row)

            # for groupName, dfGroup in dfGrouped:
            #     # 有同一个reason code有同样数量的正负值
            #     if any((dfGroup['Cut Quantity'] > 0) & (dfGroup['Cut Quantity'].abs().duplicated())):
            #
            #         pass
            #     # 不存在同一个reason code有同样数量的正负值
            #     else:
            #         pass

            # --------------不存在59 End--------------#

            # ---------------------找到有多条记录End---------------------#

    # ------------------------------ZCCR的操作End------------------------------#
    def WriteVBAPToDfMain(self):
        self.dfVBAP['VBELN'] = '0' + self.dfVBAP['VBELN']
        self.dfVBAP['MATNR'] = '0000000000' + self.dfVBAP['MATNR']
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfVBAP, ['宝洁订单号', '宝洁产品代码'],
                                            ['VBELN', 'MATNR'], ['客户产品代码'], ['KDMAT'])

    def WriteLIPSToDfMain(self):
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfLIPS, ['交货号', '宝洁产品代码'],
                                            ['DeliveryNo', 'MaterialNo'], ['软转换产品对应新码', '新码实际分货数量'],
                                            ['NewMaterial', 'NewLFIMG'])
        pass

    def WriteOpenAllotmentToDfMain(self):
        # 将日期列转换为 Pandas 的日期时间类型
        self.dfOpenAllotment['操作日期(From)'] = pd.to_datetime(self.dfOpenAllotment['操作日期(From)'])
        self.dfOpenAllotment['操作日期(To)'] = pd.to_datetime(self.dfOpenAllotment['操作日期(To)'])

        dfOpenAllotmentTemp = self.dfOpenAllotment
        dfMainTemp = self.dfMain

        # 进行连接，根据条件设置 '促销装配额开放日缺货Y/N' 列的值为 'Y'
        dfMerged = dfMainTemp.merge(dfOpenAllotmentTemp, left_on='宝洁产品代码', right_on='Item Code', how='left')

        dfMerged['促销装配额开放日缺货Y/N'] = 'N'

        condition = (dfMerged['未满足原因代码'] == '01') & (dfMerged['操作日期(From)'] <= pd.to_datetime('today')) & (
                    dfMerged['操作日期(To)'] >= pd.to_datetime('today'))
        dfMerged.loc[condition, '促销装配额开放日缺货Y/N'] = 'Y'

        self.dfMain['促销装配额开放日缺货Y/N'] = dfMerged['促销装配额开放日缺货Y/N']

    def AppedZeerToDfMain(self):
        dfFiltered = self.dfZEER.loc[(self.dfZEER['Drops Err Message'].isin(self.zeerScreenConditionList)) & (
                    self.dfZEER['Material Quantity'] != 0)]
        
        dfFiltered = PandasUtils.DeleteColumns(dfFiltered, self.zeerReservedColumnsList)

        dfFiltered.rename(columns=self.zeerReservedColumns, inplace=True)

        dfFiltered = PandasUtils.AppendColumnsToDf(dfFiltered, self.finalReportFieldList, dfFiltered.columns.to_list())

        self.dfMain = pd.concat([self.dfMain, dfFiltered])

        self.dfMain.reset_index(drop=True, inplace=True)

    def WriteCustomerListToDfMain(self):
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfCustomerList,
                                            ['付运点代码'],
                                            ['SAP Ship-to Code'],
                                            ['渠道', '区域', '市场', '客户简称'],
                                            ['Channel', 'Division', 'Market', 'Banner/RD Name'])
        pass

    def WritePriceListToDfMain(self):
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfCustomerList,
                                            ['宝洁产品代码'],
                                            ['代码'],
                                            ['品类',	'包装/箱', '不含税价格', 'MSU/箱', '产品条行码'],
                                            ['品类',	'箱规', '200箱不含税价', 'MSU/sale unit', '产品条码'])

    def FillInDfMain(self):

        self.dfMain['订单类型'] = '非提前订单'

        self.dfMain.loc[self.dfMain['AO类型'] != ''] = '提前订单'

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfCustomerList,
                                            ['未满足原因代码'],
                                            ['Cut Reason'],
                                            ['未满足原因代码', '未满足原因中文描述'],
                                            ['砍单原因', '未满足单品补货指引'])

        self.dfMain['下单数量MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(self.dfMain['下单数量'])

        self.dfMain['有效下单数量 MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(self.dfMain['有效下单数量'])

        self.dfMain['分货数量 MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(self.dfMain['分货数量'])

        self.dfMain['未满足数量MSU'] = pd.to_numeric(self.dfMain['下单数量MSU']) - pd.to_numeric(self.dfMain['分货数量 MSU'])

    def Main(self):
        # timeStart = time.time()
        # df = pd.read_excel(filePath, sheet_name="Sheet1", dtype='str')
        # print("读数据的时间为%ss"%(time.time() - timeStart))
        self.ReadFilesToDataFrame()
        self.AppendDfZderToDfMain()
        self.InsertZocrOrderValueToDfMain()
        self.WriteZccrCutReasonToDfMain()
        self.WriteVBAPToDfMain()
        self.WriteLIPSToDfMain()
        self.WriteOpenAllotmentToDfMain()
        self.AppedZeerToDfMain()
        self.WriteCustomerListToDfMain()
        self.WritePriceListToDfMain()
        self.FillInDfMain()
        pass
        # print(df)


if __name__ == '__main__':
    obj = SAMBCOptimization()
    obj.Main()
