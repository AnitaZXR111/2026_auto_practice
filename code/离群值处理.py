import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns

input_path = "../data/学生成绩表_离群值.xlsx"
output_path1 = "../output/语文英语箱线图.png"
output_path2 = "../output/离群值处理.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

def get_iqr(df,sub):
    q1 = df[sub].quantile(0.25)
    q3 = df[sub].quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1-1.5*iqr
    upper_bound = q3+1.5*iqr
    return(q1,q3,iqr,lower_bound,upper_bound)

summary_subject = pd.DataFrame({
    '语文': get_iqr(df2,'语文'),
    '英语': get_iqr(df2,'英语')}).T
summary_subject.columns = ['Q1','Q3','IQR','下界阈值','上界阈值']
print('='*50)
print('以下是语文和英语的Q1、Q3、IQR和上下界阈值：')
print('='*50)
print(summary_subject)

for subject in ['语文','英语']:
    lower = summary_subject.loc[subject, '下界阈值']
    upper = summary_subject.loc[subject, '上界阈值']
    
    df2[f'{subject}_is_outlier'] = (df2[subject] < lower) | (df2[subject] > upper)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
sns.boxplot(data=df2, x='班级', y='语文',ax = ax1)
ax1.set_xlabel('Class')
ax1.set_ylabel('Chinese')

sns.boxplot(data=df2, x='班级', y='英语',ax = ax2)
ax2.set_xlabel('Class')
ax2.set_ylabel('English')

plt.savefig(output_path1)
plt.show()


df_outlier = df2[(df2['语文_is_outlier']==True) | (df2['英语_is_outlier']==True)]

with pd.ExcelWriter(output_path2, engine='openpyxl') as writer:
    summary_subject.to_excel(writer, sheet_name = '上下界阈值', index = True)
    df_outlier.to_excel(writer, sheet_name='异常样本清单', index=False)

print(f"已成功保存至{output_path2}")




