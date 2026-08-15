import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import seaborn as sns

input_path = "../data/学生成绩表.xlsx"
output_path1 = "../output/数学成绩直方图.png"
output_path2 = "../output/语英成绩直方图.png"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

sns.histplot(data=df2, x="数学")
plt.xlabel('Math')
plt.savefig(output_path1)
plt.show()

fig, (ax1, ax2) = plt.subplots(1, 2)
sns.histplot(data=df2, x="语文", ax=ax1)
ax1.set_xlabel("Chinese") 
sns.histplot(data=df2, x="英语", ax=ax2)
ax2.set_xlabel("English") 
plt.tight_layout()
plt.savefig(output_path2)
plt.show()
