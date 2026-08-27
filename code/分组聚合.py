import pandas as pd
import numpy as np
import openpyxl


input_path = "../data/学生成绩表.xlsx"
output_path = "../output/各班各科统计汇总.xlsx"


df = pd.read_excel(input_path, sheet_name="Sheet1")
df2 = df.copy()

def get_summary(df,sub):
    summary = df.groupby("班级")[sub].agg(['count',
                             'mean',
                             'var',
                             'std',
                             'median',
                             'max',
                             'min'])
    print(f"各班{sub}成绩统计汇总：")
    print(summary)
    return(summary)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for subject in df2.columns[4:]:
        summary_subject = get_summary(df2,subject)
        summary_subject.to_excel(writer, sheet_name = subject, index = True)

print("已将统计汇总表导出至{output_path}'")
