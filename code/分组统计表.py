import openpyxl
import pandas as pd
import numpy as np


input_path = "../data/sales_data.xlsx"
output_path = "../output/分组统计表.xlsx"
sheet = "sales_data"

df = pd.read_excel(input_path)
df2 = df.copy()

df_group = df2.groupby('区域')['销量'].mean().reset_index()
df_group.rename(columns={'销量': '平均销量'}, inplace=True)

sum_group = df2.groupby(['区域', '产品'])['销量'].sum().reset_index()
sum_group_sorted = sum_group.sort_values(['区域', '销量'])

highest = sum_group_sorted.groupby('区域').tail(1)[['销量']].reset_index(drop=True)
lowest = sum_group_sorted.groupby('区域').head(1)[['销量']].reset_index(drop=True)
df_group["最高产品销量"] = highest
df_group["最低产品销量"] = lowest

with pd.ExcelWriter(output_path,engine = 'openpyxl') as writer:
    df_group.to_excel(writer,sheet_name = '分组统计',index = False)
    

print("="*50)
print(f"已保存至: {output_path}")
print("="*50)
