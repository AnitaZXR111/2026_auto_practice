import pandas as pd
import numpy as np
import openpyxl
from scipy import stats


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/day21_result.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()


#多层布尔条件(Multiple Boolean Conditions)
df_new = df2[(df2['数学']>60)
             &(df2['物理']>60)
             &(df2['化学']>60)
             &(df2['班级']==1)].reset_index()

print("="*50)
print('以下是一班数理化均及格的同学名单：')
print("="*50)
print(df_new)

#复杂切片
df_1 = df.query('班级 == 1') #一班
df1_score = df_1.iloc[:,4:] #所有学科成绩
df1_math = df1_score.loc[:,'数学'] #数学成绩
df1_math = pd.DataFrame({'姓名': df_1.loc[:,'姓名'],'数学成绩':df1_math})
print("="*50)
print('以下是一班所有同学的数学成绩：')
print("="*50)
print(df1_math)


#分段标记（分数区间）
df2['平均分'] = df2[df2.columns[4:]].mean(axis=1)
df2['分数区间'] = np.where(df2['平均分']>=90, "A",
                     np.where(df2['平均分']>=80,'B',
                              np.where(df2['平均分']>=70,'C',
                                       np.where(df2['平均分']>=60,'D','F'))))
print(f"已完成分分数区间标记，不及格共{len(df2[df2['分数区间']=='F'])}人。")


with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_new.to_excel(writer, sheet_name='多层布尔条件', index=False)
    df1_math.to_excel(writer, sheet_name='复杂切片', index=False)
    df2.to_excel(writer, sheet_name='分段标记', index=False)

print(f"已成功保存至{output_path}")
