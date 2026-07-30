import openpyxl
import pandas as pd


input_path = "../data/sales_data.xlsx"
output_path = "../output/重复与空白处理.xlsx"

df = pd.read_excel(input_path)
df2 = df.copy()

df2['日期'] = pd.to_datetime(df2['日期']).dt.date

df2 = df2.drop_duplicates()
df2['销量'].fillna(0, inplace = True)
df2['单价'].fillna(0, inplace = True)


with pd.ExcelWriter(output_path,engine = 'openpyxl') as writer:
    df2.to_excel(writer,sheet_name = '重复与空白处理',index = False)

print('='*50)
print(f"已保存至：{output_path}")
print('='*50)

