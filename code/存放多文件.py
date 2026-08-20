import pandas as pd
import numpy as np
import openpyxl

input_path = "../data/学生成绩表.xlsx"
output_path = "../output/学生成绩分析报告.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")
sheet1 = df.copy()

#sheet2
sheet2 = sheet1.groupby('班级').agg({
    '语文': ['mean', 'median', 'std', 'min', 'max'],
    '数学': ['mean', 'median', 'std', 'min', 'max'],
    '英语': ['mean', 'median', 'std', 'min', 'max']
}).round(2)
sheet2.columns = ['_'.join(col).strip() for col in sheet2.columns.values]
sheet2 = sheet2.reset_index()

#sheet3
bins = [0, 60, 70, 80, 90, 100]
labels = ['0-60', '60-70', '70-80', '80-90', '90-100']

sheet3 = []
for sub in ['语文', '数学', '英语']:
    box = pd.cut(df[sub], bins=bins, labels=labels, include_lowest=True)
    range_count = df.groupby(box, observed=False)['姓名'].count().reset_index()
    range_count.columns = ['分数区间', '频数']
    range_count['科目'] = sub
    sheet3.append(range_count)

sheet3 = pd.concat(sheet3, ignore_index=True)


#sheet4
q1 = sheet1['数学'].quantile(0.25)
q3 = sheet1['数学'].quantile(0.75)
iqr = q3 - q1

lower_bound = q1-1.5*iqr
upper_bound = q3+1.5*iqr

sheet4= sheet1[(sheet1['数学']<lower_bound) | (sheet1['数学']>upper_bound)]


# 写入Excel
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    sheet1.to_excel(writer, sheet_name='原始数据', index=False)
    sheet2.to_excel(writer, sheet_name='各班汇总', index=False)
    sheet3.to_excel(writer, sheet_name='频数分布', index=False)
    sheet4.to_excel(writer, sheet_name='异常值清单', index=False)

print(f"已保存至{output_path}")



