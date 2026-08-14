import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/剔除异常值前后对比.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()



q1 = df['数学'].quantile(0.25)
q3 = df['数学'].quantile(0.75)
iqr = q3 - q1

lower_bound = q1-1.5*iqr
upper_bound = q3+1.5*iqr

print()


df2_outlier= df2[(df2['数学']<lower_bound) | (df2['数学']>upper_bound)]
df2_cleaned = df2[(df2['数学']>=lower_bound) & (df2['数学'] <=upper_bound)]

print("以下是数学成绩为异常值的学生名单：")
print(df2_outlier['姓名'])

avg_before = round(df2['数学'].mean(), 2)
std_before = round(df2['数学'].std(), 2)
avg_after = round(df2_cleaned['数学'].mean(), 2)
std_after = round(df2_cleaned['数学'].std(), 2)

summary_table = pd.DataFrame({
    '平均分': [avg_before, avg_after],
    '标准差': [std_before, std_after]
}, index=['剔除前', '剔除后'])

print('='*50)
print('剔除异常值前后数学平均分、标准差对比')
print('='*50)
print(summary_table)

summary_table.to_excel(output_path, index = False)






