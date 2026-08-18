import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns


input_path = "../data/学生成绩表.xlsx"
output_path1 = "../output/数学箱线图.png"
output_path2 = "../output/平均分柱状图.png"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

df12 = df2[df2['班级'].isin([1, 2])]


sns.boxplot(data=df12, x='班级', y='数学')
plt.xlabel('Class')
plt.ylabel('Math')
plt.savefig(output_path1)
plt.show()


df_mean = df12.groupby('班级')[['语文', '数学', '英语']].mean().reset_index()
df_long = df_mean.melt(id_vars='班级', var_name='科目', value_name='平均分')
df_long['科目'] = df_long['科目'].map({'语文': 'Chinese', '数学': 'Math', '英语': 'English'})

sns.barplot(data=df_long, x='班级', y='平均分', hue='科目')
plt.xlabel('Class')
plt.ylabel('Avg_Score')
plt.legend(title='Subject')
plt.savefig(output_path2)
plt.show()
