import pandas as pd
import numpy as np
import openpyxl
from scipy import stats


input_path = "../data/学生成绩表_缺失值.xlsx"
output_path = "../output/缺失值处理.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

subjects = df2.columns[4:]
null_count = df2[subjects].isnull().sum()
df_null_count = pd.DataFrame({
    '缺失值数量': null_count,
    '缺失值占比': round(null_count/len(df)*100,2).astype(str)+'%'})

print("="*50)
print("以下是每一列的缺失数数量与占比：")
print("="*50)
print(df_null_count)

#处理方法1：直接删去含有缺失值的行
df_delete = df2.copy()
df_delete= df_delete.dropna()

#处理方法2：均值填充
df_mean = df2.copy()
df_mean[subjects] = df_mean[subjects].fillna(df_mean[subjects].mean())
comparison_mean = pd.DataFrame({
    '直接删去_均值': df_delete[subjects].mean(),
    '均值填充_均值': df_mean[subjects].mean(),
    '直接删去_标准差': df_delete[subjects].std(),
    '均值填充_标准差': df_mean[subjects].std()
})

print("="*50)
print("直接删去与均值填充的对比：")
print("="*50)
print(comparison_mean.round(2))

#处理方法3：中位数填充
df_median = df2.copy()
df_median[subjects] = df_median[subjects].fillna(df_median[subjects].median())
comparison_median = pd.DataFrame({
    '直接删去_均值': df_delete[subjects].mean(),
    '中位数填充_均值': df_median[subjects].mean(),
    '直接删去_标准差': df_delete[subjects].std(),
    '中位数填充_标准差': df_median[subjects].std()
})

print("="*50)
print("直接删去与中位数填充的对比：")
print("="*50)
print(comparison_median.round(2))


#处理方法4：分组均值填充
df_group_mean = df2.copy()
for subject in subjects:
    df_group_mean[subject] = df_group_mean.groupby('班级')[subject].transform(
        lambda x: x.fillna(x.mean())
    )

comparison_group_mean = pd.DataFrame({
    '直接删去_均值': df_delete[subjects].mean(),
    '分组均值填充_均值': df_group_mean[subjects].mean(),
    '直接删去_标准差': df_delete[subjects].std(),
    '分组均值填充_标准差': df_group_mean[subjects].std()
})

print("="*50)
print("直接删去与分组均值填充的对比：")
print("="*50)
print(comparison_group_mean.round(2))

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df2.to_excel(writer, sheet_name='原数据表', index=False)
    df_delete.to_excel(writer, sheet_name='直接删除', index=False)
    df_mean.to_excel(writer, sheet_name='均值填充', index=False)
    df_median.to_excel(writer, sheet_name='中位数填充', index=False)
    df_group_mean.to_excel(writer, sheet_name='分组均值填充', index=False)

print(f"已成功保存至{output_path}")




