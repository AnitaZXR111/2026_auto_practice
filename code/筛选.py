import openpyxl
import pandas as pd

input_path = "../data/sales_data.xlsx"
output_path = "../output/销售数据筛选结果.xlsx"
sheet = "sales_data"

df = pd.read_excel(input_path,sheet_name=sheet)
df_30 = df[df['销量']>=30]
df_a = df[df['产品']=='A']

df_5 = df.copy()
df_5['异常值标记'] = ''
df_5.loc[df_5['销量']<5,'异常值标记']='销量低于5'

with pd.ExcelWriter(output_path,engine='openpyxl') as writer:
    df_30.to_excel(writer,sheet_name = '销量≥30的记录',index = False)
    df_a.to_excel(writer,sheet_name = '产品A的所有记录',index = False)
    df_5.to_excel(writer,sheet_name = '销量<5标记结果',index = False)

print("="*50)
print(f"已保存至: {output_path}")
print("=" * 50)

