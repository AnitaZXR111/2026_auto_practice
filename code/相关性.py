import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns



input_path = "../data/学生成绩表.xlsx"
output_path = "../output/科目相关性热力图.png"

df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

subject = df2[['语文', '数学', '英语']]
sub_corr = subject.corr()
print("="*50)
print("各科目Pearson相关系数矩阵：")
print("="*50)
print(sub_corr)
upper_tri = sub_corr.where(np.triu(np.ones(sub_corr.shape), k=1).astype(bool))
max_corr = upper_tri.abs().max().max()
max_idx = upper_tri.abs().stack().idxmax()

print(f'呈现相对较强的相关性的两个学科：{max_idx},相关系数为{round(max_corr, 2)}')

subject.columns = ['Chinese', 'Math', 'English']
sns.heatmap(subject.corr(),cmap='Blues')
plt.savefig(output_path)
plt.show()
