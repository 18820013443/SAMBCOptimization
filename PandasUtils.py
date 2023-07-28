import os

import numpy as np
import pandas as pd


class PandasUtils:
    def __init__(self) -> None:
        pass

    @staticmethod
    def GetDataFrame(folder, fileName, sheetName):
        df = None
        filePath = os.path.join(folder, fileName)
        fileName, extension = os.path.splitext(os.path.basename(filePath))
        if not os.path.exists(filePath):
            return None

        if '.xls' in extension:
            df = pd.read_excel(filePath, sheet_name=sheetName, engine='openpyxl', dtype='str')
        elif '.csv' in extension:
            df = pd.read_csv(filePath, dtype='str')
        else:
            raise Exception('The file name is not like .xlsx, .xls, or .csv, please make sure the file name is correct!')

        df.dropna(axis=0, how='all', inplace=True)
        return df

    @staticmethod
    def GetFieldList(df):
        return df.columns.to_list()

    @staticmethod
    # columnList         --> df中所有的column列表
    # reservedColumnList --> 需要保留的列表
    def DeleteColumns(df, reservedColumnList):
        columnList = df.columns.to_list()
        dropColumnList = []

        for column in columnList:
            if column not in reservedColumnList:
                dropColumnList.append(column)
        df = df.drop(columns=dropColumnList)
        return df

    @staticmethod
    def AppendColumnsToDf(df, finalReportFieldList, reservedColumnList):
        appendColumnList = []

        # 如果columnName不在finalReportFieldList中，则需要将columnName加到dfZDER中
        for columnName in finalReportFieldList:
            if not columnName in reservedColumnList:
                appendColumnList.append(columnName)

        df[appendColumnList] = ""

    @staticmethod
    # 将dfOther中的几列数据跟新到dfMain中
    def UpdateDfMainFromDfOther(dfMain, dfOther, leftConditionColumnList, rightConditionColumnList, leftUpdateColumnList, rightUpdateColumnList, ):
        dfMerged = pd.merge(dfMain, dfOther, left_on=leftConditionColumnList, right_on=rightConditionColumnList, how='left', suffixes=('', '_x'))

        for i in range(0, len(leftUpdateColumnList)):
            if leftUpdateColumnList[i] != rightUpdateColumnList[i]:
                dfMerged[leftUpdateColumnList[i]] = dfMerged[rightUpdateColumnList[i]].fillna(dfMerged[leftUpdateColumnList[i]])
            else:
                dfMerged[leftUpdateColumnList[i]] = dfMerged[rightUpdateColumnList[i] + '_x'].fillna(
                    dfMerged[leftUpdateColumnList[i]])
            dfMain[leftUpdateColumnList[i]] = dfMerged[leftUpdateColumnList[i]]
        # dfMerged['下单数量'] = dfMerged['Order Value'].fillna(dfMerged['下单数量'])
        # dfMerged['分货号码'] = dfMerged['delivery number'].fillna(dfMerged['分货号码'])
        # dfMain['下单数量'] = dfMerged['下单数量']
        # dfMain['分货号码'] = dfMerged['分货号码']
        # return dfMain

    @staticmethod
    def HasSamePositiveAndNegativeValueForOneReason(dfGrouped):
        hasSamePositiveAndNegativeValue = False
        for groupName, dfGroup in dfGrouped:
            dfGroup['Cut Quantity'] = pd.to_numeric(dfGroup['Cut Quantity'])
            if any((dfGroup['Cut Quantity'] > 0) & (dfGroup['Cut Quantity'].abs().duplicated())):
                hasSamePositiveAndNegativeValue = True
                break
        return hasSamePositiveAndNegativeValue

    @staticmethod
    def DeleteDuplicatedAbsRows(dfGrouped):

        reserveList = []

        for groupName, dfGroup in dfGrouped:
            dfGroup['Cut Quantity'] = pd.to_numeric(dfGroup['Cut Quantity'])
            dfGroup['absValue'] = dfGroup['Cut Quantity'].abs()
            dfGroup['isNegative'] = np.where(dfGroup['Cut Quantity'] < 0, True, False)

            # 拿到absValue，并且去重
            absValueList = list(set(dfGroup['absValue'].values.tolist()))

            for absValue in absValueList:
                dfPositive = dfGroup.loc[(dfGroup['isNegative'] == False) & (dfGroup['absValue'] == absValue)]
                dfNegative = dfGroup.loc[(dfGroup['isNegative'] == True) & (dfGroup['absValue'] == absValue)]

                numPositiveRows = dfPositive.shape[0]
                numNegativeRows = dfNegative.shape[0]

                # 如果这个absValue没有正负值相等的行，则将这个absValue的所有行都保留，并且跳到下一个dfGroup
                if numPositiveRows == 0 or numNegativeRows == 0:
                    dfAbsValue = dfGroup.loc[dfGroup['absValue'] == absValue]
                    reserveList.append(dfAbsValue)
                    continue

                numDeleteRows = min(numPositiveRows, numNegativeRows)

                dfPositiveReserve = dfPositive.iloc[numDeleteRows:]
                dfNegativeReserve = dfNegative.iloc[numDeleteRows:]

                reserveList.append(dfPositiveReserve)
                reserveList.append(dfNegativeReserve)

        df = pd.concat(reserveList)
        df.reset_index(drop=True, inplace=True)
        df = df.drop(columns=['absValue', 'isNegative'])

        return df

    @staticmethod
    def ReasonCodeContainsD4(dfZCCRCutReason):
        isD4Contained = False

        if 'D4' in dfZCCRCutReason['Rsn. Code'].values:
            isD4Contained = True

        return isD4Contained

    @staticmethod
    def isOnlyD407Contained(dfZCCRCutReason):
        isOnlyD407Contained = set(dfZCCRCutReason['Rsn. Code']) == {'D4', '07'}

        return isOnlyD407Contained

    @staticmethod
    def HasNegativeQuantityForD4(dfZCCRCutReason):
        hasNegativeQuantityForD4 = (dfZCCRCutReason['Rsn. Code'] == 'D4') & (dfZCCRCutReason['Cut Quantity'] < 0)

        return hasNegativeQuantityForD4

    @staticmethod
    def IsZeroSumD4And07(dfZCCRCutReason):

        numSumD4And07 = dfZCCRCutReason.loc[(dfZCCRCutReason['Rsn. Code'] == 'D4') | (dfZCCRCutReason['Rsn. Code'] == '07'), 'Cut Quantity'].sum()
        return numSumD4And07 == 0

    @staticmethod
    def SumQtyForSameReasonCode(dfZCCRCutReason):
        df = dfZCCRCutReason.groupby('Rsn. Code')['Cut Quantity'].sum().reset_index()
        return df

    @staticmethod
    def IsMaxAbsValueUnique(dfZCCRCutReason):
        # 获取'Cut Quantity'列的绝对值
        dfZCCRCutReason['Cut Quantity'] = pd.to_numeric(dfZCCRCutReason['Cut Quantity'])
        absCutQuantity = dfZCCRCutReason['Cut Quantity'].abs()

        # 找到绝对值的最大值
        maxAbsValue = absCutQuantity.max()

        # 统计每个绝对值出现的次数
        absValueCounts = absCutQuantity.value_counts()

        # 判断绝对值的最大值是否唯一
        isMaxAbsValueUnique = absValueCounts[maxAbsValue] == 1
        return isMaxAbsValueUnique, maxAbsValue

    @staticmethod
    def FindReasonCodeForMaxAbsValue(maxAbsValue, df):
        df['Cut Quantity'] = pd.to_numeric(df['Cut Quantity'])
        df = df.loc[df['Cut Quantity'].abs() == maxAbsValue]
        df.reset_index(drop=True, inplace=True)
        numRows = df.shape[0]
        reasonCode = df.loc[numRows - 1, ['Rsn. Code']][0]
        return reasonCode

























