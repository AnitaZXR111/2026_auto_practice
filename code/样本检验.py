import pandas as pd
import numpy as np
import openpyxl
from scipy import stats


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/各分数段学生统计.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
math_1 = df[df['班级']==1]["数学"]
math_2 = df[df['班级']==2]["数学"]

p_value = stats.ttest_ind(math_1,math_2)[1]
print(f"检验所得的p-value为：{p_value:.4f}。")
if p_value<=0.05:
    print("因p_value<=0.05，两个班级数学平均分存在显著差异。")
else:
    print("因p_value>0.05，两个班级数学平均分不存在显著差异。")
