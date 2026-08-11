import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/透视表.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

df2['平均分'] = df2[['语文', '数学', '英语']].mean(axis = 1)

df_pivot = df2.pivot_table(index = '班级',
                           values = '平均分',
                           aggfunc = 'mean').round(2).reset_index()

df_pivot['等级'] = np.where(df_pivot['平均分'] >= 90, '优秀',
               np.where(df_pivot['平均分'] >= 80, '良好',
                        np.where(df_pivot['平均分'] >= 70, '中等',
                                 np.where(df_pivot['平均分'] >= 60, '及格','不及格'))))
    

print('='*50)
print('平均分均值透视表')
print('='*50)
print(df_pivot)

df_pivot.to_excel(output_path, index = False)
