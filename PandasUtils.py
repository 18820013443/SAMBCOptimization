import os
import pandas as pd


class PandasUtils:
    def __init__(self) -> None:
        pass

    @staticmethod
    def GetDataFrame(folder, fileName, sheetName):
        filePath = os.path.join(folder, fileName)
        if not os.path.exists(filePath):
            return None
        df = pd.read_excel(filePath, sheet_name=sheetName, engine='openpyxl', dtype='str')
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
    # 将dfOther中的几列数据跟新到dfMain中
    def UpdateDfMainFromDfOther(dfMain, dfOther, leftConditionColumnList, rightConditionColumnList, leftUpdateColumnList, rightUpdateColumnList):
        dfMerged = pd.merge(dfMain, dfOther, left_on=leftConditionColumnList, right_on=rightConditionColumnList, how='left')

        for i in range(len(leftUpdateColumnList) - 1):
            dfMerged[leftUpdateColumnList(i)] = dfMerged[rightUpdateColumnList(i)].fillna(dfMerged[leftUpdateColumnList(i)])
            dfMain[leftUpdateColumnList(i)] = dfMerged[leftUpdateColumnList(i)]
        # dfMerged['下单数量'] = dfMerged['Order Value'].fillna(dfMerged['下单数量'])
        # dfMerged['分货号码'] = dfMerged['delivery number'].fillna(dfMerged['分货号码'])
        # dfMain['下单数量'] = dfMerged['下单数量']
        # dfMain['分货号码'] = dfMerged['分货号码']