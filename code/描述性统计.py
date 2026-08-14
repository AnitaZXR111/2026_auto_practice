import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/描述性统计.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

score = df2.iloc[:,3:6]
print('='*50)
print('成绩基础指标汇总')
print('='*50)
print(score.describe().round(2))

score_summary_additional = pd.DataFrame({"median":score.median(),
                              "mode":score.apply('mode').iloc[0],
                              "range":score.max() - score.min(),
                              "variance":score.var().round(2)}).T


print('='*50)
print('补充成绩描述统计汇总')
print('='*50)
print(score_summary_additional)

score_summary = pd.concat([score.describe().round(2),score_summary_additional])
score_summary.to_excel(output_path, index = True)


class_name = df['班级'].unique()
for c in class_name:
    class_score = df[df['班级']==c].iloc[:,3:6]
    score1 = class_score.describe().round(2)
    score2 = pd.DataFrame({"median":class_score.median(),
                              "mode":class_score.apply('mode').iloc[0],
                              "range":class_score.max() - class_score.min(),
                              "variance":class_score.var().round(2)}).T
    score_sum = pd.concat([score1,score2])
    print(f"{c}班成绩描述统计汇总：")
    print(score_sum)
    




    
