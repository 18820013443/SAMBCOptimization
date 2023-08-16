import sys
import json
import numpy as np
import pandas as pd
import time
import os
from YamlHandler import Settings
from PandasUtils import PandasUtils
from Log import Log
import traceback
from datetime import datetime
from ExcelUtils import ExcelUtils


class SAMBCOptimization:
    def __init__(self, isJIT, marketName, strDocumentDate) -> None:
        # self.dirName = self.GetDirName()
        self.dirName = Settings.GetDirName()
        self.settings = Settings.config
        self.Initialize()
        self.dfMain = None

        log = Log()
        log.logPath = self.GetLogPath()
        self.logger = log.GetLog()

        self.marketName = marketName
        self.isJIT = isJIT
        self.strDocumentDate = strDocumentDate
        # self.mainFolder = '%s/Input Files' % os.getcwd() if self.isTestMode else os.getcwd()
        self.mainFolder = '%s/Input Files' % self.dirName if self.isTestMode else self.dirName

    # def GetDirName(self):
    #     scriptPath = os.path.abspath(__file__)
    #     dirName = os.path.dirname(scriptPath)
    #     return dirName
    
    def GetLogPath(self):
        strLogPath = self.settings['logPath']
        logPath = strLogPath if strLogPath != '' else os.path.join(self.dirName, 'log.txt')
        return logPath

    def Initialize(self):
        # self.settings = YamlHandler(os.path.join(
        #     os.getcwd(), 'config.yaml')).ReadYaml()
        # self.settings = YamlHandler(os.path.join(
        #     self.dirName, 'config.yaml')).ReadYaml()
        self.isTestMode = self.settings['isTestMode']
        self.zderRevisedColumns = self.settings['zderRevisedColumns']
        self.zderReservedColumnList = self.zderRevisedColumns.keys()
        self.finalReportFieldList = self.settings['finalReportFieldList']
        self.zeerScreenConditionList = self.settings['zeerScreenConditionList']
        self.zeerReservedColumns = self.settings['zeerReservedColumns']
        self.zeerReservedColumnsList = self.zeerReservedColumns.keys()
        self.priceBraket = self.settings['priceBraket']
        pass

    def ReadFilesToDataFrame(self):

        # 构造dfMain
        self.dfMain = pd.DataFrame(columns=self.finalReportFieldList)

        # 读取dfZDER
        # self.dfZDER = PandasUtils.GetDataFrame(self.mainFolder, 'ZDER.xlsx', 'Sheet1')

        # # 将ZOCR.xls转成ZOCR.xlsx，并且读取dfZOCR
        # ExcelUtils(os.path.join(self.mainFolder, 'ZOCR.xls')).ConvertXlsToXlsx()
        # self.dfZOCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZOCR.xlsx', 'ZOCR')

        # # 读取dfZCCR
        # self.dfZCCR = PandasUtils.GetDataFrame(self.mainFolder, 'ZCCR.xlsx', 'Sheet1')

        # 读取dfVBAK
        self.dfVBAK = PandasUtils.GetDataFrame(self.mainFolder, 'VBAK.xlsx', 'Sheet1', parseDatesList=['ERDAT', 'AUDAT'])
        # self.dfVBAK = PandasUtils.GetDataFrame(self.mainFolder, 'VBAK.xlsx', 'Sheet1')

        # 读取dfVBAP
        self.dfVBAP = PandasUtils.GetDataFrame(self.mainFolder, 'VBAP.csv', 'VBAP')

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

        # 修改Price Braket
        for key in self.priceBraket.keys():
            self.dfMain['固定箱数折扣'].loc[self.dfMain['固定箱数折扣'] == key] = self.priceBraket[key]
        # print(self.dfMain)

    # ------------------------------ZDER的操作End------------------------------#

    # ------------------------------ZOCR的操作Start------------------------------#
    def InsertZocrOrderValueToDfMain(self):
        self.dfZOCR['Sales Order No.'] = '0' + self.dfZOCR['Sales Order No.']
        self.dfZOCR['Material No.'] = '0000000000' + self.dfZOCR['Material No.']
        # self.dfZOCR = self.dfZOCR.dropna(subset=['Material No.'])

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfZOCR,
                                            ['宝洁订单号', '宝洁产品代码'],
                                            ['Sales Order No.', 'Material No.'],
                                            ['分货金额含税（供参考，发票为准）'],
                                            ['Order Value'])

    # ------------------------------ZOCR的操作End------------------------------#

    # ------------------------------ZCCR的操作Start------------------------------#

    def CalculateNotSatisfiedQty(self):
        # 将字符串列转换为数值类型
        # self.dfMain['未满足数量'] = pd.to_numeric(self.dfMain['未满足数量'])
        # self.dfMain['下单数量'] = pd.to_numeric(self.dfMain['下单数量'])
        # self.dfMain['分货数量'] = pd.to_numeric(self.dfMain['分货数量'])

        # 计算未满足数量
        self.dfMain['未满足数量'] = pd.to_numeric(self.dfMain['下单数量']) - pd.to_numeric(self.dfMain['分货数量'])
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
            # 删除有同样数量和同样reason code的正负值的记录
            dfZCCRCutReasonRemoveValues = PandasUtils.DeleteDuplicatedAbsRows(dfGrouped)
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
                # 如果包含'07'那么reason code = 07
                if '07' in dfZCCRCutReasonNew['Rsn. Code'].values:
                    row['未满足原因代码'] = '07'

                # 否则就是最大值的reason code
                else:
                    row['未满足原因代码'] = reasonCode
        # 绝对值的最大值不是是唯一
        else:
            # reasonCode = PandasUtils.FindReasonCodeForMaxAbsValue(maxAbsValue, dfZCCRCutReasonNew)
            row['未满足原因代码'] = reasonCode

    def DfMainDeleteD8Records(self):
        dfMainD8Cut = self.dfMain.loc[self.dfMain['未满足原因代码'] == 'D8']
        if dfMainD8Cut.shape[0] == 0:
            return

        # Remove rows where '未满足原因代码' is equal to 'D8'
        self.dfMain = self.dfMain[self.dfMain['未满足原因代码'] != 'D8']

        # Filter out rows where '客户订单号' contains 'xdc', '_xdc', or '.'
        self.dfMain = self.dfMain[~self.dfMain['客户订单号'].str.contains('xdc|_xdc|\.')]

        self.dfMain.reset_index(drop=True, inplace=True)
        pass

    def WriteZccrCutReasonToDfMain(self):
        # 判断ZCCR是否存在
        if self.dfZCCR is None:
            return

        # 将dfZCCR的Cut Quantity改为float类型
        self.dfZCCR['Cut Quantity'] = pd.to_numeric(self.dfZCCR['Cut Quantity'])

        self.CalculateNotSatisfiedQty()

        dfUnsatisfiedQty = self.dfMain.loc[(self.dfMain['未满足数量'] != np.nan) | (self.dfMain['未满足数量'] != 0)]
        # dfUnsatisfiedQty = self.dfMain.loc[self.dfMain['未满足数量'] == 0]

        dfUnsatisfiedQty.drop_duplicates(subset=['宝洁订单号', '宝洁产品代码'], keep='first')

        # 仅用于测试， test Only
        # if self.isTestMode:
        #     dfUnsatisfiedQty = PandasUtils.GenerateDfForTest()

        for index, row in dfUnsatisfiedQty.iterrows():
            try:
                strSalesOrder = row['宝洁订单号'].lstrip('0')
                strMaterial = row['宝洁产品代码'].lstrip('0')
                dfZCCRCutReason = self.dfZCCR.loc[
                    (self.dfZCCR['Sales Order'] == strSalesOrder) & (self.dfZCCR['Material'] == strMaterial)]
                # print(dfZCCRCutReason.shape[0])
                # 没有找到记录，继续下一条
                if dfZCCRCutReason.shape[0] == 0:
                    continue
                # 找到只有一条记录，赋值，继续下一条
                if dfZCCRCutReason.shape[0] == 1:
                    # row['未满足原因代码'] = dfZCCRCutReason.at[0, 'Rsn. Code']
                    row['未满足原因代码'] = dfZCCRCutReason.iloc[0]['Rsn. Code']
                    dfUnsatisfiedQty.iloc[index] = row
                    continue

                # ---------------------找到有多条记录Start---------------------#

                # 筛选出cut reason为59的记录
                dfCutReason59 = dfZCCRCutReason.loc[dfZCCRCutReason['Rsn. Code'] == '59']

                # 如果存在59的记录，则赋值59，继续下一条
                if dfCutReason59.shape[0] > 0:
                    row['未满足原因代码'] = '59'
                    dfUnsatisfiedQty.iloc[index] = row
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
                        # if isD4Contained:
                        self.Contains04Operations(True, dfGrouped, dfZCCRCutReason, row)

                # 同一个reason code有不同数量的正负值
                else:
                    if isD4Contained:
                        self.OnlyD4And07Operations(dfGrouped, dfZCCRCutReason, row)
                    else:
                        self.Contains04Operations(False, dfGrouped, dfZCCRCutReason, row)

                dfUnsatisfiedQty.iloc[index] = row
            except Exception as e:
                rowData = str(row.to_json(indent=2, force_ascii=False))
                strE = traceback.format_exc()
                # self.logger.debug(('Write ZCCR Cut Reason Error and the row data is: \n %s' % rowData).encode('utf8'))
                self.logger.debug('Write ZCCR Cut Reason Error and the row data is: \n %s' % rowData)
                raise Exception(e)

                    # --------------不存在59 End--------------#

            # ---------------------找到有多条记录End---------------------#
        

        # 将dfMain的类型转成字符类型
        self.dfMain = self.dfMain.astype('object')

        # 将dfUnsatisfiedQty中的cutReason回写到self.dfMain中
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, dfUnsatisfiedQty,
                                            ['宝洁订单号', '宝洁产品代码'],
                                            ['宝洁订单号', '宝洁产品代码'],
                                            ['未满足原因代码'],
                                            ['未满足原因代码'],)

    # ------------------------------ZCCR的操作End------------------------------#

    def WriteVBAKToDfMain(self):
        if self.dfVBAK is None:
            return

        self.dfVBAK['VBELN'] = '0' + self.dfVBAK['VBELN']
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfVBAK,
                                            ['宝洁订单号'],
                                            ['VBELN'],
                                            ['订单生成日'],
                                            ['ERDAT']
                                            )

    def WriteVBAPToDfMain(self):
        if self.dfVBAP is None:
            return

        # self.dfVBAP['VBELN'] = '0' + self.dfVBAP['VBELN']
        # self.dfVBAP['MATNR'] = '0000000000' + self.dfVBAP['MATNR']
        self.dfVBAP.drop_duplicates(['VBELN', 'MATNR'], keep='first', inplace=True)
        self.dfVBAP.reset_index(drop=True, inplace=True)
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfVBAP,
                                            ['宝洁订单号', '宝洁产品代码'],
                                            ['VBELN', 'MATNR'],
                                            ['客户产品代码'],
                                            ['KDMAT'])

    def WriteLIPSToDfMain(self):
        if self.dfLIPS is None:
            return

        dfLIPSNotNull = self.dfLIPS.loc[self.dfLIPS['NewMaterial'].notnull()]

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, dfLIPSNotNull,
                                            ['交货号', '宝洁产品代码'],
                                            ['DeliveryNo', 'MaterialNo'],
                                            ['软转换产品对应新码', '新码实际分货数量'],
                                            ['NewMaterial', 'NewLFIMG'])
        pass

    def WriteOpenAllotmentToDfMain(self):
        # 将日期列转换为 Pandas 的日期时间类型
        self.dfOpenAllotment['操作日期(From)'] = pd.to_datetime(self.dfOpenAllotment['操作日期(From)'])
        self.dfOpenAllotment['操作日期(To)'] = pd.to_datetime(self.dfOpenAllotment['操作日期(To)'])

        # 加前置0
        self.dfOpenAllotment['Item Code'] = '0000000000' + self.dfOpenAllotment['Item Code']

        dfOpenAllotmentTemp = self.dfOpenAllotment
        dfMainTemp = self.dfMain

        # 进行连接，根据条件设置 '促销装配额开放日缺货Y/N' 列的值为 'Y'
        dfMerged = dfMainTemp.merge(dfOpenAllotmentTemp, left_on='宝洁产品代码', right_on='Item Code', how='left')

        dfMerged['促销装配额开放日缺货Y/N'] = 'N'

        # condition = (dfMerged['未满足原因代码'] == '01') & (dfMerged['操作日期(From)'] <= pd.to_datetime('today')) & (
        #         dfMerged['操作日期(To)'] >= pd.to_datetime('today'))
        currentDay = self.dfMain.loc[0, '分货日']
        condition = (dfMerged['未满足原因代码'] == '01') & (dfMerged['操作日期(From)'] <= currentDay) & (
                dfMerged['操作日期(To)'] >= currentDay)
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
        if self.dfPriceList is None:
            return

        # self.dfPriceList['代码'] = '0' + self.dfPriceList['代码']
        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfPriceList,
                                            ['宝洁产品代码'],
                                            ['Material'],
                                            ['品类', '箱规', '200箱不含税价', 'MSU/sale unit', '产品条码'],
                                            ['品类', '包装/箱', '不含税价格', 'MSU/箱', '产品条行码'])

    def FillInDfMain(self):

        self.dfMain['订单类型'] = '非提前订单'

        # self.dfMain['订单类型'].loc[self.dfMain['AO类型'] != ''] = '提前订单'
        self.dfMain.loc[self.dfMain['AO类型'].notnull(), '订单类型'] = '提前订单'

        # self.dfMain.loc[self.dfMain['AO类型'] != '']['订单类型'] = '提前订单'
        
        # 将dfMain的类型转成字符类型
        self.dfMain = self.dfMain.astype('object')

        PandasUtils.UpdateDfMainFromDfOther(self.dfMain, self.dfCutReasonList,
                                            ['未满足原因代码'],
                                            ['Cut Reason'],
                                            ['未满足原因中文描述', '未满足单品补货指引'],
                                            ['砍单原因', '未满足单品补货指引'])

        self.dfMain['下单数量MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(
            self.dfMain['下单数量'])

        self.dfMain['有效下单数量 MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(
            self.dfMain['有效下单数量'])

        self.dfMain['分货数量 MSU'] = pd.to_numeric(self.dfMain['MSU/sale unit']) * pd.to_numeric(
            self.dfMain['分货数量'])

        self.dfMain['未满足数量MSU'] = pd.to_numeric(self.dfMain['下单数量MSU']) - pd.to_numeric(
            self.dfMain['分货数量 MSU'])

    def FormatDfMain(self):
        # 订单生成日，分货日，到货日的格式改为yyyy/MM/dd
        # self.dfMain['订单生成日'] = self.dfMain['订单生成日'].str.slice(0, 10).str.replace('-', '/')
        # self.dfMain['分货日'] = self.dfMain['分货日'].str.slice(0, 10).str.replace('-', '/')
        # self.dfMain['到货日'] = self.dfMain['到货日'].str.slice(0, 10).str.replace('-', '/')

        # 宝洁订单号去除前置0
        self.dfMain['宝洁订单号'] = self.dfMain['宝洁订单号'].str.lstrip('0')

        # 宝洁产品代码去除前置0
        self.dfMain['宝洁产品代码'] = self.dfMain['宝洁产品代码'].str.lstrip('0')

        # 交货号去除前置0
        self.dfMain['交货号'] = self.dfMain['交货号'].str.lstrip('0')

        # 软转换产品对应新码去除前置0
        self.dfMain['软转换产品对应新码'] = self.dfMain['软转换产品对应新码'].str.lstrip('0')

        pass

    def GenerateDailyReport(self):

        # 构造daily report name
        date = datetime.strptime(self.strDocumentDate, '%d.%m.%Y')
        strDate = date.strftime('%Y%m%d')

        strJIT = ' JIT' if isJIT else ''
        strDailyReportName = '{} Daily SAMBC Report {}{}.xlsx'.format(self.marketName, strDate, strJIT)

        # 将剪切板的数据插入到Final Report Template.xlsx中, strDocumentDate = 'dd.MM.yyyy'
        filePath = os.path.join(self.mainFolder, 'Final Report Template.xlsx')

        # 将dfMain写入clipboard, 并且写入excel保存
        # self.dfMain.to_clipboard(index=False, header=False, sep='\t', encoding='utf-8')
        # ExcelUtils(filePath).SaveXlsx('Sheet1', strDailyReportName)

        ExcelUtils(filePath).SaveXlsxWings('Sheet1', strDailyReportName, self.dfMain)
        return strDailyReportName

    def CombineDataLogic(self):
        self.ReadFilesToDataFrame()
        self.logger.info('Files have been read to dataframe')

        self.AppendDfZderToDfMain()
        self.logger.info('Append ZDER to dfMain')

        self.InsertZocrOrderValueToDfMain()
        self.logger.info('Insert ZOCR to dfMain')

        self.WriteZccrCutReasonToDfMain()
        self.logger.info('Write ZCCR Cut Reason to dfMain')

        self.DfMainDeleteD8Records()
        self.logger.info('Delete D8')

        self.WriteVBAKToDfMain()
        self.logger.info('Write VBAK to dfMain')

        self.WriteVBAPToDfMain()
        self.logger.info('Write VBAP to dfMain')

        self.WriteLIPSToDfMain()
        self.logger.info('Write LIPS to dfMain')

        self.WriteOpenAllotmentToDfMain()
        self.logger.info('Write Open Allotment to dfMain')

        self.AppedZeerToDfMain()
        self.logger.info('Append ZEER to dfMain')

        self.WriteCustomerListToDfMain()
        self.logger.info('Write Customer List to dfMain')

        self.WritePriceListToDfMain()
        self.logger.info('Write Price List to dfMain')

        self.FillInDfMain()
        self.logger.info('Fill in dfMain')

        self.FormatDfMain()
        self.logger.info('Format dfMain')

        # self.dfMain.to_excel('%s output.xlsx' % self.marketName, index=False)
        # self.logger.info('Write %s output.xlsx' % self.marketName)

        strDailyReportName = self.GenerateDailyReport()
        self.logger.info('Generate report %s' % strDailyReportName)
        pass

    def Main(self):

        timeStart = time.time()
        
        startParameters = {
            'isJIT': self.isJIT,
            'marketName': self.marketName,
            'strDocumentDate': self.strDocumentDate
        }

        # dataJson = json.dumps(startParameters).encode('utf8')
        dataJson = str(startParameters)

        # self.logger.info(('Start with paramaters: %s' % dataJson).encode('utf8'))
        self.logger.info('Start with paramaters: %s' % dataJson)

        try:
            self.CombineDataLogic()
        except Exception as e:
            strE = traceback.format_exc()
            self.logger.debug(strE)
            sys.exit(1)
        
        self.logger.info('Total run time is %ds' % int(time.time() - timeStart))


if __name__ == '__main__':
    
    isTestMode = Settings.config.get('isTestMode')

    if not isTestMode:

        # 定义执行文件的入参
        args = sys.argv
        isJIT = args[1].lower() == 'true'
        marketName = args[2]
        strDocumentDate = args[3]
    else:

        # 测试配置
        isJIT = False
        marketName = 'GBJ'
        strDocumentDate = '16.08.2023'

    # 实例化对象并且执行Main方法
    obj = SAMBCOptimization(isJIT, marketName, strDocumentDate)
    obj.Main()
