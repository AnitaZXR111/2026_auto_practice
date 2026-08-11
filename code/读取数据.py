import pandas as pd
import numpy as np
import openpyxl

input_path = "../data/sales_data.xlsx"

df_origin = pd.read_excel(input_path, sheet_name="Sheet1")
df_origin2 = df_origin.copy()
print("="*50)
print("Sheet1中第五行的数据为:")
print(df_origin2.iloc[3,:])
print("="*50)

df_result = pd.read_excel(input_path, sheet_name="Sheet2")
df_result2 = df_result.copy()
highest_sales = df_result2[df_result2['平均销量']==max(df_result2['平均销量'])]
print(f"Sheet2中，平均销量最高的区域为{highest_sales.iloc[0,0]}。")
print(f"其平均销量为{highest_sales.iloc[0,1]}，最高和最低销量分别为{highest_sales.iloc[0,2]}和{highest_sales.iloc[0,3]}。")
