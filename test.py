import pandas as pd

# 创建包含数据的DataFrame
df = pd.DataFrame({
    'Rsn. Code': ['01', '01', '01', '02', '03', '03'],
    'Cut Quantity': [500, -500, 500, 200, -200, 300]
})

# 根据'Rsn. Code'分组，并检查每个组内是否存在具有相同正负值的'Cut Quantity'
grouped = df.groupby('Rsn. Code')

for group_name, group_df in grouped:
    if any((group_df['Cut Quantity'] > 0) & (group_df['Cut Quantity'].abs().duplicated())):
        print(f"'Rsn. Code'为 {group_name} 的行存在具有相同正负值的'Cut Quantity'")
    else:
        print(f"'Rsn. Code'为 {group_name} 的行不存在具有相同正负值的'Cut Quantity'")