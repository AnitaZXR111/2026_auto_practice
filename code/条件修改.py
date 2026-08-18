import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/不及格标记.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

df2['平均分'] = round(df2[['语文', '数学', '英语']].mean(axis = 1),1)
df2['是否优等生'] = np.where(df2['平均分'] >= 85, True, False)

f_good = df2[(df2['性别'] == '女') & (df2['是否优等生'] == True)]

print('='*50)
print('以下是所有的优等生女生：')
print('='*50)
print(f_good)

def highlight_fail(row):
    return [f'color: red' if val < 60 else 'color: black' for val in row]

df2[['语文', '数学', '英语']] = df2[['语文', '数学', '英语']].where(
    df2[['语文', '数学', '英语']] >= 60,
    df2[['语文', '数学', '英语']].astype(str) + '(不及格)'
)
df2.to_excel(output_path, index = False)
