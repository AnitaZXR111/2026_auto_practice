import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns

input_path1 = "../data/学生成绩表_离群值.xlsx"
input_path2 = "../output/离群值处理.xlsx"
output_path = "../output/重合异常对比表.xlsx"


df = pd.read_excel(input_path1, sheet_name="Sheet1")
df2 = df.copy()



def get_z_value(df, sub):
    avg = df[sub].mean()
    std = df[sub].std()
    z = (df[sub]-avg)/std
    return(z)

df2["语文z值"] = get_z_value(df2, "语文")
df2["英语z值"] = get_z_value(df2, "英语")

df_outlier_z = df2[(abs(df2['语文z值'])>3) | (abs(df2['英语z值'])>3)]

df_outlier_iqr = pd.read_excel(input_path2, sheet_name="离群值清单")
common_rows = df_outlier_z.merge(df_outlier_iqr, how='inner')
print("="*50)
print("以下是IQR和Z值测算出的重合异常名单：")
print("="*50)
print(common_rows)
print("="*50)

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    common_rows.to_excel(writer, sheet_name = '汇总对比表', index = False)
print(f'共{len(common_rows)}个重合异常样本。已将汇总对比表导出至{output_path}')
