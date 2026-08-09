import pandas as pd
import numpy as np
import openpyxl

input_path = "../data/sales_data.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()
df2['日期'] = pd.to_datetime(df2['日期']).dt.date

citys = df2['区域'].unique()

for city in citys:
    df_city = df2[df2['区域'] == city]
    
    output_path = f'../data/divide/{city}.xlsx'
    with pd.ExcelWriter(output_path,engine = 'openpyxl') as writer:
        df_city.to_excel(writer,sheet_name = '拆分数据',index = False)
    

print("="*50)
print(f"城市分组excel已保存至divide文件夹。")
print("="*50)


output_path = "../output/merge.xlsx"
all_data = []
for city in citys:
    filepath = f"../data/divide/{city}.xlsx"
    df = pd.read_excel(filepath, sheet_name="拆分数据")
    df2 = df.copy()
    df2['日期'] = pd.to_datetime(df2['日期']).dt.date
    all_data.append(df2)

df_merged = pd.concat(all_data, ignore_index=True)
total_row = {}
for col in df_merged.columns:
    if col == '销量': 
        total_row[col] = df_merged[col].sum()
    elif col == '产品':
        total_row[col] = '合计'
    else:
        total_row[col] = ''

df_merged = pd.concat([df_merged, pd.DataFrame([total_row])], ignore_index=True)
with pd.ExcelWriter(output_path,engine = 'openpyxl') as writer:
    df_merged.to_excel(writer,sheet_name = '合并数据',index = False)

print("="*50)
print(f"合并完成。总行数: {len(df_merged)} 行。")
print(f"文件保存至: {output_path}")
print("="*50)
