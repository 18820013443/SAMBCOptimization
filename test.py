import pandas as pd
import traceback
from PandasUtils import PandasUtils


def test1():  # 创建包含数据的DataFrame
    df = pd.DataFrame({
        'Rsn. Code': ['01', '01', '01', '02', '02', '03', '03'],
        'Cut Quantity': [500, -500, 500, 200, -200, -200, 300]
    })

    # 根据'Rsn. Code'分组，并检查每个组内是否存在具有相同正负值的'Cut Quantity'
    grouped = df.groupby('Rsn. Code')

    for group_name, group_df in grouped:
        if any((group_df['Cut Quantity'] > 0) & (group_df['Cut Quantity'].abs().duplicated())):
            print(f"'Rsn. Code'为 {group_name} 的行存在具有相同正负值的'Cut Quantity'")
        else:
            print(f"'Rsn. Code'为 {group_name} 的行不存在具有相同正负值的'Cut Quantity'")


def test2():
    # 创建包含数据的DataFrame
    df = pd.DataFrame({
        'Rsn. Code': ['01', '01', '01', '02', '02', '03', '03', '03', '03'],
        'Cut Quantity': [500, -500, -500, 200, -200, -200, 300, 200, 400]
    })

    dfGrouped = df.groupby(df['Rsn. Code'])

    dfResult = PandasUtils.DeleteDuplicatedAbsRows(dfGrouped)

    dfList = []
    for groupName, dfGroup in dfGrouped:
        # 判断是否存在具有相同正负值的'Cut Quantity'
        has_duplicates = dfGroup.groupby(['Rsn. Code', dfGroup['Cut Quantity'].abs()]).filter(lambda x: len(x) == 2)

        # 删除具有相同正负值的行
        filtered_df = dfGroup.drop(has_duplicates.index)

        # 重新设置索引
        filtered_df.reset_index(drop=True, inplace=True)
        dfList.append(filtered_df)
        pass

    # # 判断是否存在具有相同正负值的'Cut Quantity'
    # has_duplicates = df.groupby(['Rsn. Code', df['Cut Quantity'].abs()]).filter(lambda x: len(x) == 2)
    #
    # # 删除具有相同正负值的行
    # filtered_df = df.drop(has_duplicates.index)
    #
    # # 重新设置索引
    # filtered_df.reset_index(drop=True, inplace=True)

    # 打印结果
    # print(filtered_df)


def test3():
    # 创建包含数据的DataFrame
    df = pd.DataFrame({
        'Rsn. Code': ['D4', '07', '01'],
        'Cut Quantity': [500, -500, -700]
    })

    maxAbsValue = df['Cut Quantity'].abs().max()

    a = PandasUtils.FindReasonCodeForMaxAbsValue(maxAbsValue, df)
    a = PandasUtils.IsZeroSumD4And07(df)

    # 利用transform将每个分组的绝对值进行转换
    df['abs_quantity'] = df.groupby('Rsn. Code')['Cut Quantity'].transform(lambda x: abs(x))

    # 判断是否存在正负值相同的行
    has_duplicates = df.duplicated(subset=['Rsn. Code', 'abs_quantity'], keep=False)

    # 提取正负值相同的行
    result = df[has_duplicates]

    # 打印结果
    print(result)


def test4():
    try:
        try:
            1 / 0
        except Exception as e:
            strE = traceback.format_exc()
            raise Exception(e)
    except Exception as e:
        strE = traceback.format_exc()
        print(strE)
        pass


if __name__ == '__main__':
    test4()
