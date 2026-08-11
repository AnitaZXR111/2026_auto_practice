import openpyxl
import pandas as pd
import numpy as np


input_path = "../data/sales_data.xlsx"
output_path = "../output/销量计算.xlsx"
sheet = "sales_data"

df = pd.read_excel(input_path)
df2 = df.copy()

df2['日期'] = pd.to_datetime(df2['日期']).dt.date

df2['金额'] = df2['销量']*df2['单价']
df2['等级'] = np.where(df2['销量'] >= 45, '优秀',
               np.where(df2['销量'] >= 20, '良好', '不及格'))

with pd.ExcelWriter(output_path,engine='openpyxl') as writer:
    df2.to_excel(writer,sheet_name = '金额与销量等级',index = False)

print("="*50)
print(f"已保存至: {output_path}")
print("="*50)
