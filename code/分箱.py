import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/各分数段学生统计.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

bins = [0,60,70,80,90,100]
labels = ['0‑60','60‑70','70‑80','80‑90','90‑100']

df2['分箱'] = pd.cut(df2['数学'],
                   bins = bins,
                   labels = labels,
                   include_lowest = True)

range_count = df2.pivot_table(index = '分箱',
                values = '姓名',
                aggfunc = 'count',
                fill_value = 0).reset_index()

print('='*50)
print('各个分数区间内学生数')
print('='*50)
print(range_count)

range_count.to_excel(output_path,index = False)
