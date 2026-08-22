import pandas as pd
import numpy as np
import openpyxl
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns


input_path = "../data/学生成绩表.xlsx"
output_path = "../final_output"

def translate_subject(subject_name_original):
    subject_name_translated = ''
    if subject_name_original=="语文":
        subject_name_translated = 'Chinese'
    elif subject_name_original=="数学":
        subject_name_translated = 'Math'
    elif subject_name_original=="英语":
        subject_name_translated = 'English'
    elif subject_name_original=="物理":
        subject_name_translated = 'Physics'
    elif subject_name_original=="化学":
        subject_name_translated = 'Chemistry'
    else:
       subject_name_translated = 'Other'
    return(subject_name_translated)


def import_data(in_path):
    df = pd.read_excel(in_path)
    df2 = df.copy()
    subjects = df.columns[4:].tolist()
    translated_subjects = {col: translate_subject(col) for col in subjects}
    df2.rename(columns=translated_subjects, inplace=True)
    print(f"已导入数据，共{len(df2)}行。")
    return(df2)



def clean_data(df):
    df = df.drop_duplicates()#剔除重复行
    subjects = df.columns[4:]
    df[subjects] = df[subjects].fillna(df[subjects].mean()) #空白格填充平均值
    print(f"已完成对缺失与重复值的处理，处理后共{len(df)}行。")
    return(df)

def add_avg_total(df):
    subjects = df.columns[4:]
    df['平均分'] = df[subjects].mean(axis = 1).round(2)
    df['总分'] = df[subjects].sum(axis = 1)
    print("已完成对学生成绩平均分与总分的计算。")
    return(df)

def box_sort(df):
    bins = [0,60,70,80,90,100]
    labels = ['0‑60','60‑70','70‑80','80‑90','90‑100']

    df['分箱'] = pd.cut(df['平均分'],
                   bins = bins,
                   labels = labels,
                   include_lowest = True)
    print("已完成对学生成绩平均分的分箱处理。")
    return(df)

def get_summary(df):
    df = box_sort(df)
    range_count = df.pivot_table(index = '分箱',
                                 columns = "班级",
                                 values = '姓名',
                                 aggfunc = 'count',
                                 fill_value = 0).reset_index()
    print("已完成各班分箱统计：")
    print(range_count)
    return(range_count)

def drop_outlier(df):
    subjects = df.columns[4:]
    exclude_cols = ['分箱', '平均分', '总分']
    subjects = [col for col in subjects if col not in exclude_cols]   
    df_new = df.copy()   
    for sub in subjects:
        q1 = df_new[sub].quantile(0.25)
        q3 = df_new[sub].quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        df_new[sub] = df_new[sub].clip(lower=lower_bound, upper=upper_bound)    
    print(f"已处理异常值（替换为边界值），处理后共{len(df_new)}行。")
    return(df_new)

def calculate_corr(df):
    subjects = df.columns[4:]
    exclude_cols = ['分箱', '平均分', '总分']
    subjects = [col for col in subjects if col not in exclude_cols] 
    sub_corr = df[subjects].corr()

    upper_tri = sub_corr.where(np.triu(np.ones(sub_corr.shape), k=1).astype(bool))
    max_corr = upper_tri.abs().max().max()
    max_idx = upper_tri.abs().stack().idxmax()
    print(f'呈现相对较强的相关性的两个学科：{max_idx},相关系数为{round(max_corr, 2)}')
    return(sub_corr)

def corr_heatmap(corr_matrix,out_path):
    sns.heatmap(corr_matrix,cmap='Blues')
    plt.savefig(out_path)
    plt.show()

def create_histogram(df,sub,out_path):
    sns.histplot(data=df, x=sub)
    plt.tight_layout()
    plt.savefig(out_path)
    plt.show()
    
def create_boxplot(df,sub,out_path):
    sns.boxplot(data=df, x='班级', y=sub)
    plt.xlabel('Class')
    plt.savefig(out_path)
    plt.show()

def create_barplot(df,out_path):
    subjects = df.columns[4:]
    exclude_cols = ['分箱', '平均分', '总分']
    subjects = [col for col in subjects if col not in exclude_cols] 
    df_mean = df.groupby('班级')[subjects].mean().reset_index()
    df_long = df_mean.melt(id_vars='班级', var_name='科目', value_name='平均分')

    sns.barplot(data=df_long, x='班级', y='平均分', hue='科目')
    plt.xlabel('Class')
    plt.ylabel('Avg_Score')
    plt.legend(title='Subject')
    plt.savefig(out_path)
    plt.show()

def export_to_excel(df_original, df_summary, corr_matrix, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_original.to_excel(writer, sheet_name='原始数据', index=False)
        df_summary.to_excel(writer, sheet_name='分箱统计', index=False)
        corr_matrix.to_excel(writer, sheet_name='相关性矩阵')
    print(f"分析报告已导出至: {output_path}")
        

if __name__ == "__main__":
    # 读取excel
    df = import_data(input_path)
    # 清洗缺失、重复值
    df = clean_data(df)
    # 新增总分平均分
    df = add_avg_total(df)
    # 分箱分级
    df = box_sort(df)
    # 按班级分组统计
    summary = get_summary(df)
    # 异常值排查 
    df_clean = drop_outlier(df)
    # 相关性计算
    corr_matrix = calculate_corr(df_clean)
    corr_heatmap(corr_matrix, f"{output_path}/correlation_heatmap.png")
    
    subjects = df.columns[4:]
    exclude_cols = ['分箱', '平均分', '总分']
    subjects = [col for col in subjects if col not in exclude_cols] 
    # 直方图
    for sub in subjects:
        create_histogram(df_clean, sub, f"{output_path}/{sub}_histogram.png")
    # 箱线图
    for sub in subjects:
        create_boxplot(df_clean, sub, f"{output_path}/{sub}_boxplot.png")
    # 柱状图
    create_barplot(df_clean, f"{output_path}/subject_avg_barplot.png")
    
    # 多sheet导出分析报告
    export_to_excel(df_clean, summary, corr_matrix, f"{output_path}/分析报告.xlsx")


    
                                                     
